"""
Núcleo da consulta de CT-e: certificados, sincronização com a SEFAZ, cache local
e extração dos dados (tomador, destinatário, valor, data, notas...).

A SEFAZ não permite buscar um CT-e por chave no serviço de Distribuição DF-e
(CTeDistribuicaoDFe). Ele entrega uma fila numerada (NSU) com todos os documentos
em que o CNPJ do certificado é parte. Por isso o app sincroniza essa fila de cada
certificado cadastrado para um cache local (SQLite) e a busca por chave consulta
esse cache.
"""

import base64
import gzip
import json
import os
import re
import shutil
import sqlite3
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime, timezone

import requests
from lxml import etree
from cryptography import x509
from cryptography.hazmat.primitives.serialization import (
    pkcs12, Encoding, PrivateFormat, NoEncryption,
)
from cryptography.x509.oid import NameOID

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

try:
    import keyring
except ImportError:
    keyring = None

NS = "http://www.portalfiscal.inf.br/cte"
WSDL_NS = "http://www.portalfiscal.inf.br/cte/wsdl/CTeDistribuicaoDFe"
ACTION = "http://www.portalfiscal.inf.br/cte/wsdl/CTeDistribuicaoDFe/cteDistDFeInteresse"

NS_NFE = "http://www.portalfiscal.inf.br/nfe"
WSDL_NFE = "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe"
URLS_NFE_DIST = {
    "producao":    "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx",
    "homologacao": "https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx",
}

URLS_DIST = {
    "producao":    "https://www1.cte.fazenda.gov.br/CTeDistribuicaoDFe/CTeDistribuicaoDFe.asmx",
    "homologacao": "https://hom1.cte.fazenda.gov.br/CTeDistribuicaoDFe/CTeDistribuicaoDFe.asmx",
}

UF_SIGLA = {
    "11": "RO", "12": "AC", "13": "AM", "14": "RR", "15": "PA", "16": "AP", "17": "TO",
    "21": "MA", "22": "PI", "23": "CE", "24": "RN", "25": "PB", "26": "PE", "27": "AL",
    "28": "SE", "29": "BA", "31": "MG", "32": "ES", "33": "RJ", "35": "SP", "41": "PR",
    "42": "SC", "43": "RS", "50": "MS", "51": "MT", "52": "GO", "53": "DF",
}
UF_CODIGO = {v: k for k, v in UF_SIGLA.items()}

APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ConsultaCTe")
CERTS_DIR = os.path.join(APP_DIR, "certs")
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
DB_PATH = os.path.join(APP_DIR, "cache.db")
KEYRING_SERVICE = "ConsultaCTe"


# ── Certificados ─────────────────────────────────────────────────────────────

class CertificadoErro(Exception):
    pass


def ler_certificado(pfx_path: str, senha: str) -> dict:
    """Valida o .pfx e devolve nome, CNPJ (da extensão SAN), UF e validade."""
    try:
        with open(pfx_path, "rb") as f:
            chave, cert, _ = pkcs12.load_key_and_certificates(f.read(), senha.encode("utf-8"))
    except Exception as e:
        raise CertificadoErro(f"Não foi possível abrir o certificado (senha incorreta?): {e}")
    if cert is None or chave is None:
        raise CertificadoErro("O arquivo não contém certificado e chave privada.")

    cnpj = None
    try:
        san = cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
        for outro in san.get_values_for_type(x509.OtherName):
            # 2.16.76.1.3.3 = CNPJ da empresa titular (ICP-Brasil). O "OU" do
            # subject é da autoridade de registro e não deve ser usado.
            if outro.type_id.dotted_string == "2.16.76.1.3.3":
                m = re.search(rb"\d{14}", outro.value)
                if m:
                    cnpj = m.group(0).decode()
    except x509.ExtensionNotFound:
        pass
    cn = cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)
    nome = cn[0].value if cn else "(sem nome)"
    if not cnpj:
        m = re.search(r":(\d{14})\b", nome)
        cnpj = m.group(1) if m else None
    if not cnpj:
        raise CertificadoErro("Não encontrei o CNPJ no certificado (precisa ser e-CNPJ A1).")

    uf_attr = cert.subject.get_attributes_for_oid(NameOID.STATE_OR_PROVINCE_NAME)
    uf = UF_CODIGO.get(uf_attr[0].value.upper(), "43") if uf_attr else "43"

    return {
        "nome": re.sub(r":\d{14}$", "", nome),
        "cnpj": cnpj,
        "uf": uf,
        "validade": cert.not_valid_after_utc.date().isoformat(),
        "vencido": cert.not_valid_after_utc < datetime.now(timezone.utc),
    }


@contextmanager
def pem_temporarios(pfx_path: str, senha: str):
    """Converte o .pfx em PEM temporários (requests exige arquivos) e os apaga ao sair."""
    with open(pfx_path, "rb") as f:
        chave, cert, cadeia = pkcs12.load_key_and_certificates(f.read(), senha.encode("utf-8"))
    pem_cert = cert.public_bytes(Encoding.PEM) + b"".join(
        c.public_bytes(Encoding.PEM) for c in (cadeia or []))
    pem_key = chave.private_bytes(Encoding.PEM, PrivateFormat.TraditionalOpenSSL, NoEncryption())
    arquivos = []
    try:
        for dados in (pem_cert, pem_key):
            tf = tempfile.NamedTemporaryFile(suffix=".pem", delete=False)
            tf.write(dados)
            tf.close()
            arquivos.append(tf.name)
        yield tuple(arquivos)
    finally:
        for a in arquivos:
            try:
                os.unlink(a)
            except OSError:
                pass


# ── Configuração (lista de certificados) ─────────────────────────────────────

def carregar_config() -> dict:
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"certificados": []}


def salvar_config(cfg: dict) -> None:
    os.makedirs(APP_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def adicionar_certificado(pfx_path: str, senha: str, lembrar: bool) -> dict:
    """Valida, copia o .pfx para a pasta do app e registra. Substitui se o CNPJ já existir."""
    info = ler_certificado(pfx_path, senha)
    os.makedirs(CERTS_DIR, exist_ok=True)
    destino = os.path.join(CERTS_DIR, f"{info['cnpj']}.pfx")
    if os.path.abspath(pfx_path) != os.path.abspath(destino):
        shutil.copyfile(pfx_path, destino)

    cfg = carregar_config()
    anterior = next((c for c in cfg["certificados"] if c["cnpj"] == info["cnpj"]), None)
    certs = [c for c in cfg["certificados"] if c["cnpj"] != info["cnpj"]]
    entrada = {**info, "arquivo": destino, "situacao": "não testado", "ultima_sync": None,
               "estabelecimentos": (anterior or {}).get("estabelecimentos")
               or {info["cnpj"]: _novo_estab()}}
    certs.append(entrada)
    cfg["certificados"] = certs
    salvar_config(cfg)

    if lembrar and keyring:
        try:
            keyring.set_password(KEYRING_SERVICE, info["cnpj"], senha)
        except Exception:
            pass    # sem Gerenciador de Credenciais: a senha só vale nesta sessão
    return entrada


def remover_certificado(cnpj: str) -> None:
    cfg = carregar_config()
    for c in cfg["certificados"]:
        if c["cnpj"] == cnpj:
            try:
                os.unlink(c["arquivo"])
            except OSError:
                pass
    cfg["certificados"] = [c for c in cfg["certificados"] if c["cnpj"] != cnpj]
    salvar_config(cfg)
    if keyring:
        try:
            keyring.delete_password(KEYRING_SERVICE, cnpj)
        except Exception:
            pass


def senha_salva(cnpj: str) -> str | None:
    if not keyring:
        return None
    try:
        return keyring.get_password(KEYRING_SERVICE, cnpj)
    except Exception:
        return None


def lembrar_senha(cnpj: str, senha: str) -> None:
    if keyring:
        try:
            keyring.set_password(KEYRING_SERVICE, cnpj, senha)
        except Exception:
            pass


# ── Cache local ──────────────────────────────────────────────────────────────

class Cache:
    def __init__(self, caminho: str = DB_PATH):
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        self.caminho = caminho
        with self._con() as con:
            con.execute("CREATE TABLE IF NOT EXISTS cte ("
                        "chave TEXT PRIMARY KEY, xml BLOB NOT NULL, cnpj_cert TEXT)")
            con.execute("CREATE TABLE IF NOT EXISTS nfe (chave TEXT PRIMARY KEY, xml BLOB NOT NULL)")

    def _con(self):
        return sqlite3.connect(self.caminho)

    def salvar(self, chave: str, xml: bytes, cnpj_cert: str) -> bool:
        """True se era novo."""
        with self._con() as con:
            cur = con.execute("INSERT OR IGNORE INTO cte VALUES (?,?,?)", (chave, xml, cnpj_cert))
            return cur.rowcount > 0

    def obter(self, chave: str) -> bytes | None:
        con = self._con()
        try:
            linha = con.execute("SELECT xml FROM cte WHERE chave=?", (chave,)).fetchone()
        finally:
            con.close()
        return linha[0] if linha else None

    def salvar_nfe(self, chave: str, xml: bytes) -> None:
        with self._con() as con:
            con.execute("INSERT OR REPLACE INTO nfe VALUES (?,?)", (chave, xml))

    def obter_nfe(self, chave: str) -> bytes | None:
        con = self._con()
        try:
            linha = con.execute("SELECT xml FROM nfe WHERE chave=?", (chave,)).fetchone()
        finally:
            con.close()
        return linha[0] if linha else None

    def total(self) -> int:
        con = self._con()
        try:
            return con.execute("SELECT COUNT(*) FROM cte").fetchone()[0]
        finally:
            con.close()


# ── Sincronização com a SEFAZ ────────────────────────────────────────────────

def _pagina(cert_pem, key_pem, cnpj, uf, ult_nsu, ambiente):
    tp_amb = "1" if ambiente == "producao" else "2"
    dist = (f'<distDFeInt versao="1.00" xmlns="{NS}"><tpAmb>{tp_amb}</tpAmb>'
            f"<cUFAutor>{uf}</cUFAutor><CNPJ>{cnpj}</CNPJ>"
            f"<distNSU><ultNSU>{ult_nsu:015d}</ultNSU></distNSU></distDFeInt>")
    envelope = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soapenv:Envelope xmlns:soapenv="http://www.w3.org/2003/05/soap-envelope">'
        f'<soapenv:Header/><soapenv:Body><cteDistDFeInteresse xmlns="{WSDL_NS}">'
        f"<cteDadosMsg>{dist}</cteDadosMsg></cteDistDFeInteresse></soapenv:Body></soapenv:Envelope>"
    )
    resp = requests.post(
        URLS_DIST[ambiente], data=envelope.encode("utf-8"),
        headers={"Content-Type": f'application/soap+xml;charset=UTF-8;action="{ACTION}"'},
        cert=(cert_pem, key_pem), timeout=60)
    try:
        raiz = etree.fromstring(resp.content)
    except etree.XMLSyntaxError:
        raise RuntimeError(f"HTTP {resp.status_code}: resposta inválida da SEFAZ")
    ret = raiz.find(f".//{{{NS}}}retDistDFeInt")
    if ret is None:
        if resp.status_code == 403:
            raise RuntimeError("SEFAZ recusou o certificado (403) — pode estar revogado.")
        raise RuntimeError(f"HTTP {resp.status_code}: " + " ".join(raiz.itertext()).strip()[:200])
    return ret


def _novo_estab() -> dict:
    return {"ult_nsu": 0, "max_nsu": 0, "proxima": 0}


def estabelecimentos_de(cert_cfg: dict) -> dict:
    """CNPJs consultados com este certificado ({cnpj: estado da fila})."""
    est = cert_cfg.get("estabelecimentos")
    if not est:
        est = {cert_cfg["cnpj"]: {"ult_nsu": cert_cfg.get("ult_nsu", 0),
                                  "max_nsu": cert_cfg.get("max_nsu", 0), "proxima": 0}}
    return est


def definir_estabelecimentos(cnpj_cert: str, cnpjs: list[str]) -> tuple[list[str], list[str]]:
    """
    Define os estabelecimentos (filiais) consultados com o certificado. A SEFAZ só aceita
    CNPJs da mesma raiz (8 primeiros dígitos) do certificado; o CNPJ do próprio certificado
    sempre fica na lista. Retorna (aceitos, recusados).
    """
    cfg = carregar_config()
    cert = next(c for c in cfg["certificados"] if c["cnpj"] == cnpj_cert)
    atuais = estabelecimentos_de(cert)
    aceitos, recusados = [], []
    for texto in [cnpj_cert] + list(cnpjs):
        cnpj = re.sub(r"\D", "", texto or "")
        if not cnpj:
            continue
        if len(cnpj) == 14 and cnpj[:8] == cnpj_cert[:8]:
            if cnpj not in aceitos:
                aceitos.append(cnpj)
        else:
            recusados.append(texto.strip())
    cert["estabelecimentos"] = {c: atuais.get(c) or _novo_estab() for c in aceitos}
    salvar_config(cfg)
    return aceitos, recusados


def _gravar_estado(cnpj_cert: str, estabs: dict, situacao: str | None = None) -> None:
    cfg = carregar_config()
    for c in cfg["certificados"]:
        if c["cnpj"] == cnpj_cert:
            c["estabelecimentos"] = estabs
            if situacao is not None:
                c["situacao"] = situacao
                c["ultima_sync"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    salvar_config(cfg)


def sincronizar(cert_cfg: dict, senha: str, cache: Cache, ambiente: str = "producao",
                progresso=None, max_paginas: int = 500) -> dict:
    """
    Baixa os documentos novos da fila de cada estabelecimento do certificado para o cache.
    Retorna {'novos': n, 'situacao': str}.
    """
    novos = 0
    problemas: list[str] = []
    adiados = 0
    estabs = estabelecimentos_de(cert_cfg)
    try:
        with pem_temporarios(cert_cfg["arquivo"], senha) as (c, k):
            for cnpj, est in estabs.items():
                if est.get("proxima", 0) > time.time():
                    adiados += 1                # SEFAZ: aguardar 1h após "nenhum documento"
                    continue
                for _ in range(max_paginas):
                    ret = _pagina(c, k, cnpj, cert_cfg["uf"], est["ult_nsu"], ambiente)
                    c_stat = (ret.findtext(f"{{{NS}}}cStat") or "").strip()
                    motivo = (ret.findtext(f"{{{NS}}}xMotivo") or "").strip()
                    if c_stat == "138":
                        for dz in ret.iter(f"{{{NS}}}docZip"):
                            xml = gzip.decompress(base64.b64decode(dz.text))
                            raiz = etree.fromstring(xml)
                            if raiz.tag.endswith("cteProc"):
                                chave = raiz.findtext(f".//{{{NS}}}infProt/{{{NS}}}chCTe")
                                if chave and cache.salvar(chave, xml, cnpj):
                                    novos += 1
                        est["ult_nsu"] = int(ret.findtext(f"{{{NS}}}ultNSU"))
                        est["max_nsu"] = int(ret.findtext(f"{{{NS}}}maxNSU"))
                        if progresso:
                            progresso(f"{formatar_cnpj(cnpj)}: NSU {est['ult_nsu']}/{est['max_nsu']} "
                                      f"(+{novos} CT-e)")
                        if est["ult_nsu"] >= est["max_nsu"]:
                            break
                    elif c_stat == "137":       # nada novo
                        est["max_nsu"] = int(ret.findtext(f"{{{NS}}}maxNSU") or est["max_nsu"])
                        est["proxima"] = time.time() + 3600
                        break
                    elif c_stat == "656":       # consumo indevido
                        est["proxima"] = time.time() + 3600
                        problemas.append("limite da SEFAZ atingido (aguarde 1h)")
                        break
                    else:
                        problemas.append(f"{formatar_cnpj(cnpj)}: {c_stat} — {motivo}")
                        break
                _gravar_estado(cert_cfg["cnpj"], estabs)
    except RuntimeError as e:
        problemas.append(str(e))
    except requests.RequestException as e:
        problemas.append(f"erro de conexão: {e}")

    if problemas:
        situacao = problemas[0] + (f" (+{len(problemas) - 1} avisos)" if len(problemas) > 1 else "")
    elif adiados == len(estabs):
        situacao = "ok — nada novo (nova consulta liberada em até 1h)"
    else:
        situacao = "ok"
    _gravar_estado(cert_cfg["cnpj"], estabs, situacao)
    cert_cfg["situacao"] = situacao
    return {"novos": novos, "situacao": situacao}


# ── NF-e: transportadora da nota ─────────────────────────────────────────────

class NfeNaoEncontrada(Exception):
    pass


def _consulta_nfe(cert_pem, key_pem, cnpj, uf, chave, ambiente):
    tp_amb = "1" if ambiente == "producao" else "2"
    dist = (f'<distDFeInt versao="1.01" xmlns="{NS_NFE}"><tpAmb>{tp_amb}</tpAmb>'
            f"<cUFAutor>{uf}</cUFAutor><CNPJ>{cnpj}</CNPJ>"
            f"<consChNFe><chNFe>{chave}</chNFe></consChNFe></distDFeInt>")
    envelope = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soap12:Envelope xmlns:soap12="http://www.w3.org/2003/05/soap-envelope"><soap12:Body>'
        f'<nfeDistDFeInteresse xmlns="{WSDL_NFE}"><nfeDadosMsg>{dist}</nfeDadosMsg>'
        "</nfeDistDFeInteresse></soap12:Body></soap12:Envelope>"
    )
    resp = requests.post(URLS_NFE_DIST[ambiente], data=envelope.encode("utf-8"),
                         headers={"Content-Type": "application/soap+xml;charset=UTF-8"},
                         cert=(cert_pem, key_pem), timeout=60)
    try:
        raiz = etree.fromstring(resp.content)
    except etree.XMLSyntaxError:
        raise RuntimeError(f"HTTP {resp.status_code}: resposta inválida da SEFAZ")
    ret = raiz.find(f".//{{{NS_NFE}}}retDistDFeInt")
    if ret is None:
        raise RuntimeError(f"HTTP {resp.status_code}: " + " ".join(raiz.itertext()).strip()[:200])
    return ret


_MOD_FRETE = {"0": "Remetente (CIF)", "1": "Destinatário (FOB)", "2": "Terceiros",
              "3": "Próprio — remetente", "4": "Próprio — destinatário", "9": "Sem frete"}


def _t(no, caminho: str) -> str:
    if no is None:
        return ""
    el = no.find("/".join(f"{{{NS_NFE}}}{p}" for p in caminho.split("/")))
    return (el.text or "").strip() if el is not None else ""


def extrair_nfe(nfe_xml: etree._Element) -> dict:
    """Dados da NF-e (nfeProc/NFe), com foco na transportadora declarada."""
    inf = nfe_xml.find(f".//{{{NS_NFE}}}infNFe")
    if inf is None:
        raise ValueError("XML sem infNFe")
    ide, emit, dest = (inf.find(f"{{{NS_NFE}}}{t}") for t in ("ide", "emit", "dest"))
    transp = inf.find(f"{{{NS_NFE}}}transp")
    tr = transp.find(f"{{{NS_NFE}}}transporta") if transp is not None else None
    veic = transp.find(f"{{{NS_NFE}}}veicTransp") if transp is not None else None
    volumes = transp.findall(f"{{{NS_NFE}}}vol") if transp is not None else []
    prot = nfe_xml.find(f".//{{{NS_NFE}}}infProt")
    mod = _t(transp, "modFrete")

    def soma(tag):
        total = 0.0
        for v in volumes:
            try:
                total += float(_t(v, tag))
            except ValueError:
                pass
        return total

    return {
        "chave": (inf.get("Id") or "").replace("NFe", ""),
        "numero": _t(ide, "nNF"),
        "serie": _t(ide, "serie"),
        "data": _t(ide, "dhEmi") or _t(ide, "dEmi"),
        "natureza": _t(ide, "natOp"),
        "valor": _t(inf, "total/ICMSTot/vNF"),
        "emitente": {"nome": _t(emit, "xNome"), "doc": _t(emit, "CNPJ") or _t(emit, "CPF")},
        "destinatario": {"nome": _t(dest, "xNome"), "doc": _t(dest, "CNPJ") or _t(dest, "CPF")},
        "mod_frete": mod,
        "mod_frete_txt": _MOD_FRETE.get(mod, mod),
        "transportadora": {
            "nome": _t(tr, "xNome"),
            "doc": _t(tr, "CNPJ") or _t(tr, "CPF"),
            "ie": _t(tr, "IE"),
            "endereco": _t(tr, "xEnder"),
            "municipio": f'{_t(tr, "xMun")}/{_t(tr, "UF")}' if tr is not None and _t(tr, "xMun") else "",
        },
        "veiculo": f'{_t(veic, "placa")} {_t(veic, "UF")}'.strip(),
        "volumes": int(soma("qVol")),
        "peso_bruto": soma("pesoB"),
        "peso_liquido": soma("pesoL"),
        "protocolo": _t(prot, "nProt"),
        "status": _t(prot, "cStat"),
        "motivo": _t(prot, "xMotivo"),
    }


def consultar_nfe(chave: str, certificados: list[tuple[dict, str]], cache: Cache,
                  ambiente: str = "producao", progresso=None) -> dict:
    """
    Obtém a NF-e pela chave (Distribuição DF-e da NF-e) e devolve extrair_nfe().
    `certificados` é a lista (cadastro, senha). A SEFAZ responde 138 para o certificado cuja
    empresa (raiz de CNPJ) é parte da nota e 640 para os demais, então testa um por um.
    O resultado fica no cache; a próxima consulta da mesma chave não vai à SEFAZ.
    """
    xml = cache.obter_nfe(chave)
    if xml is not None:
        return extrair_nfe(etree.fromstring(xml))

    resumo = False
    erros: list[str] = []
    for cert, senha in certificados:
        if progresso:
            progresso(f"Consultando com {cert['nome']}…")
        try:
            with pem_temporarios(cert["arquivo"], senha) as (c, k):
                ret = _consulta_nfe(c, k, cert["cnpj"], cert["uf"], chave, ambiente)
        except (RuntimeError, requests.RequestException) as e:
            erros.append(f"{cert['nome']}: {e}")
            continue
        c_stat = (ret.findtext(f"{{{NS_NFE}}}cStat") or "").strip()
        motivo = (ret.findtext(f"{{{NS_NFE}}}xMotivo") or "").strip()
        if c_stat == "138":
            for dz in ret.iter(f"{{{NS_NFE}}}docZip"):
                raiz = etree.fromstring(gzip.decompress(base64.b64decode(dz.text)))
                if raiz.tag.endswith("nfeProc"):
                    cache.salvar_nfe(chave, etree.tostring(raiz))
                    return extrair_nfe(raiz)
                resumo = True      # resNFe: o interessado ainda não tem o XML completo
        elif c_stat == "656":
            erros.append("limite de consultas da SEFAZ atingido — aguarde 1 hora")
            break
        elif c_stat != "640":
            erros.append(f"{cert['nome']}: {c_stat} — {motivo}")
    if resumo:
        raise NfeNaoEncontrada(
            "A SEFAZ só liberou o resumo desta NF-e (sem o campo de transporte). O XML completo "
            "exige a manifestação de “Ciência da Operação” do destinatário.")
    msg = ("Nenhum dos certificados cadastrados é parte desta NF-e (emitente, destinatário ou "
           "transportador), ou ela não existe.")
    if erros:
        msg += "\n\n" + "\n".join(erros[:4])
    raise NfeNaoEncontrada(msg)


# ── Chave e extração dos dados ───────────────────────────────────────────────

def normalizar_chave(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def dv_chave_ok(chave: str) -> bool:
    """Confere o dígito verificador (módulo 11) da chave de 44 dígitos."""
    if len(chave) != 44 or not chave.isdigit():
        return False
    soma = sum(int(d) * (2 + i % 8) for i, d in enumerate(reversed(chave[:43])))
    resto = soma % 11
    return int(chave[43]) == (0 if resto in (0, 1) else 11 - resto)


def decodificar_chave(chave: str) -> dict:
    """Dados que já estão na própria chave de 44 dígitos."""
    return {
        "uf": UF_SIGLA.get(chave[0:2], chave[0:2]),
        "emissao_aamm": f"{chave[4:6]}/20{chave[2:4]}",
        "cnpj_emitente": chave[6:20],
        "modelo": chave[20:22],
        "serie": str(int(chave[22:25])),
        "numero": str(int(chave[25:34])),
    }


def _txt(no, caminho: str) -> str:
    if no is None:
        return ""
    el = no.find("/".join(f"{{{NS}}}{p}" for p in caminho.split("/")))
    return (el.text or "").strip() if el is not None else ""


def _parte(no) -> dict:
    if no is None:
        return {"nome": "", "doc": ""}
    return {"nome": _txt(no, "xNome"), "doc": _txt(no, "CNPJ") or _txt(no, "CPF")}


_TOMA3 = {"0": "rem", "1": "exped", "2": "receb", "3": "dest"}


def extrair_dados(cte_xml: etree._Element) -> dict:
    """Lê os dados do XML do CT-e (cteProc/CTe)."""
    inf = cte_xml.find(f".//{{{NS}}}infCte")
    if inf is None:
        raise ValueError("XML sem infCte")

    ide = inf.find(f"{{{NS}}}ide")
    tomador_tipo = ""
    toma3 = ide.find(f"{{{NS}}}toma3")
    toma4 = ide.find(f"{{{NS}}}toma4")
    if toma3 is not None:
        tomador_tipo = _TOMA3.get(_txt(toma3, "toma"), "")
        tomador = (_parte(inf.find(f"{{{NS}}}{tomador_tipo}")) if tomador_tipo
                   else {"nome": "", "doc": ""})
    elif toma4 is not None:
        tomador = _parte(toma4)
    else:
        tomador = {"nome": "", "doc": ""}

    # Frete CIF/FOB não existe no CT-e: deduz pelo tomador (remetente paga = CIF).
    frete = {"rem": "CIF", "dest": "FOB"}.get(tomador_tipo, "")

    notas = []
    for ch in inf.findall(f".//{{{NS}}}infNFe/{{{NS}}}chave"):
        chave_nf = (ch.text or "").strip()
        if len(chave_nf) == 44:
            notas.append({"nota": str(int(chave_nf[25:34])),
                          "serie": str(int(chave_nf[22:25])), "chave": chave_nf})

    prot = cte_xml.find(f".//{{{NS}}}infProt")
    return {
        "chave": (inf.get("Id") or "").replace("CTe", ""),
        "cfop": _txt(ide, "CFOP"),
        "natureza": _txt(ide, "natOp"),
        "frete": frete,
        "notas": notas,
        "numero": _txt(ide, "nCT"),
        "serie": _txt(ide, "serie"),
        "data": _txt(ide, "dhEmi"),
        "valor": _txt(inf, "vPrest/vTPrest"),
        "valor_receber": _txt(inf, "vPrest/vRec"),
        "tomador": tomador,
        "emitente": _parte(inf.find(f"{{{NS}}}emit")),
        "remetente": _parte(inf.find(f"{{{NS}}}rem")),
        "destinatario": _parte(inf.find(f"{{{NS}}}dest")),
        "origem": f'{_txt(ide, "xMunIni")}/{_txt(ide, "UFIni")}',
        "destino": f'{_txt(ide, "xMunFim")}/{_txt(ide, "UFFim")}',
        "protocolo": _txt(prot, "nProt"),
        "status": _txt(prot, "cStat"),
        "motivo": _txt(prot, "xMotivo"),
        "tomador_tipo": tomador_tipo,
    }


def mesma_empresa(doc_a: str, doc_b: str) -> bool:
    """Mesma empresa = mesma raiz de CNPJ (matriz e filiais contam como iguais); CPF compara inteiro."""
    a, b = re.sub(r"\D", "", doc_a or ""), re.sub(r"\D", "", doc_b or "")
    if len(a) == 14 and len(b) == 14:
        return a[:8] == b[:8]
    return a == b


def formatar_cnpj(doc: str) -> str:
    if len(doc) == 14:
        return f"{doc[:2]}.{doc[2:5]}.{doc[5:8]}/{doc[8:12]}-{doc[12:]}"
    if len(doc) == 11:
        return f"{doc[:3]}.{doc[3:6]}.{doc[6:9]}-{doc[9:]}"
    return doc


def formatar_valor(v: str) -> str:
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        return v


def formatar_data(v: str) -> str:
    try:
        return datetime.fromisoformat(v).strftime("%d/%m/%Y")
    except ValueError:
        return v
