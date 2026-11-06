"""Conferência Legado x Protheus por filial e CFOP.

Protheus = sistema novo (relatório Validação); Legado = sistema anterior (abas por filial).
"""
import argparse
import json
import os
import threading
import tkinter as tk
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from tkinter import filedialog, messagebox

from core import theme
from fiscal.leitura import abas, cfop, destino_seguro, filial, texto, valor

SEM_CFOP = 'SEM CFOP'


def ler(path):
    totais, nomes, auditoria, pendencias = defaultdict(Decimal), {}, [], []
    aliases = {'f': {'filial'}, 'c': {'codfiscal', 'cfop'}, 'v': {'vlrtotal', 'total'}}
    for aba, iterator in abas(path):
        cab, fixo, qtd = None, None, 0
        for n, row in enumerate(iterator, 1):
            row = list(row)
            if cab is None:
                norm = [texto(x) for x in row]
                for i, x in enumerate(norm[:-1]):
                    if x == 'filial' and row[i+1] not in (None, ''):
                        try: fixo = filial(row[i+1])
                        except ValueError: pass
                found = {k: next((i for i,x in enumerate(norm) if x in a), None) for k,a in aliases.items()}
                if found['c'] is not None and found['v'] is not None and (found['f'] is not None or fixo):
                    cab = found
                elif n > 80: break
                continue
            if not any(x not in (None, '') for x in row): continue
            c = cfop(row[cab['c']])
            if c is None:
                # Ignora apenas cabeçalhos repetidos e totais claramente identificados.
                if texto(row[cab['c']]) in aliases['c'] or any(texto(x).startswith(('total', 'subtotal')) for x in row): continue
                if str(row[cab['c']] or '').strip() == '':
                    c = SEM_CFOP   # mantém o valor no total e registra a linha na aba Pendências
                else:
                    raise ValueError(f'{Path(path).name} / {aba} / linha {n}: CFOP inválido {row[cab["c"]]!r}')
            try:
                original = str(row[cab['f']]).strip() if cab['f'] is not None else fixo
                f = filial(original)
                if f in nomes and nomes[f] != original:
                    raise ValueError(f'Filiais distintas com o mesmo sufixo {f}: {nomes[f]} e {original}')
                nomes[f] = original
                montante = valor(row[cab['v']])
                totais[f,c] += montante
                if c == SEM_CFOP:
                    pendencias.append([Path(path).name, aba, n, original, float(montante), 'CFOP em branco',
                                       json.dumps(row, ensure_ascii=False, default=str)])
                qtd += 1
            except Exception as e: raise ValueError(f'{Path(path).name} / {aba} / linha {n}: {e}') from e
        if cab is not None: auditoria.append([Path(path).name, aba, qtd])
    if not totais: raise ValueError(f'Não encontrei dados com Filial, CFOP/Cod. Fiscal e Total/Vlr.Total em {Path(path).name}')
    # Soma os valores originais e arredonda somente o total do grupo.
    return {k: v.quantize(Decimal('.01'), rounding=ROUND_HALF_UP) for k,v in totais.items()}, nomes, auditoria, pendencias


def comparar(novo, antigo):
    a, nomes, loga, pa = ler(novo)
    b, _, logb, pb = ler(antigo)
    rows = []
    chave = lambda k: (int(k[0]), int(k[1]) if k[1].isdigit() else 9999)
    for f,c in sorted(a.keys() | b.keys(), key=chave):
        x,y = a.get((f,c), Decimal(0)), b.get((f,c), Decimal(0))
        status = 'Ausente no Protheus' if (f,c) not in a else 'Ausente no Legado' if (f,c) not in b else 'OK' if x == y else 'Divergente'
        if c == SEM_CFOP: status = 'Pendente: sem CFOP'
        rows.append([f,c,float(y),float(x),float(y-x),status,nomes.get(f,f)])
    return {'rows':rows,'log':loga+logb,'pendencias':pa+pb,'novo':str(novo),'anterior':str(antigo)}


def resumo(dados):
    """Contagens reais para o histórico do módulo fiscal."""
    rows = dados['rows']
    ok = sum(r[5] == 'OK' for r in rows)
    pend = sum(r[5].startswith('Pendente') for r in rows)
    return {'total': len(rows), 'conformes': ok, 'divergencias': len(rows) - ok - pend, 'revisao': pend,
            'pendencias': [f'Filial {r[0]} / CFOP {r[1]}: {r[5]} (Legado {r[2]:,.2f} x Protheus {r[3]:,.2f})' for r in rows if r[5] != 'OK']}


def excel(dados, destino):
    import xlsxwriter
    with xlsxwriter.Workbook(destino) as w:
        s=w.add_worksheet('Resultado'); s.hide_gridlines(2)
        h=w.add_format({'bold':True,'bg_color':'#002060','font_color':'white','border':1,'align':'center'})
        normal=w.add_format({'border':1,'align':'center'})
        money=w.add_format({'border':1,'num_format':'#,##0.00;[Red]-#,##0.00'})
        red=w.add_format({'border':1,'num_format':'#,##0.00;[Red]-#,##0.00','bg_color':'#FCE4D6'})
        s.set_column('A:A',12);s.set_column('B:B',20);s.set_column('C:C',13);s.set_column('D:E',20);s.set_column('F:F',24)
        s.write_row(0,0,['LEGADO','Soma de TOTAL','PROTHEUS','Soma de Vlr.Total','Diferença','Situação'],h)
        r=1;atual=None
        for f,c,b,a,d,status,nome in dados['rows']:
            if f!=atual:
                s.write_row(r,0,[f,'Soma de TOTAL',nome,'Soma de Vlr.Total','Diferença','Situação'],h);r+=1;atual=f
            s.write(r,0,c if status!='Ausente no Legado' else '',normal)
            s.write_number(r,1,b,money)
            s.write(r,2,c if status!='Ausente no Protheus' else '',normal)
            s.write_number(r,3,a,money)
            s.write_formula(r,4,f'=ROUND(B{r+1}-D{r+1},2)',red if d else money,d)
            s.write(r,5,status,normal);r+=1
        s.write_row(r,0,['TOTAL',sum(x[2] for x in dados['rows']),'TOTAL',sum(x[3] for x in dados['rows'])],h)
        s.write_formula(r,4,f'=ROUND(B{r+1}-D{r+1},2)',money)
        s.freeze_panes(1,0);s.set_landscape();s.fit_to_pages(1,0);s.repeat_rows(0)
        t=w.add_worksheet('Dados agrupados')
        t.write_row(0,0,['Filial','CFOP','Legado','Protheus','Diferença','Situação','Filial Protheus'],h)
        for i,row in enumerate(dados['rows'],1): t.write_row(i,0,row)
        t.autofilter(0,0,len(dados['rows']),6);t.freeze_panes(1,0);t.set_column('A:B',12);t.set_column('C:E',20,money);t.set_column('F:G',24)
        l=w.add_worksheet('Leitura');l.write_row(0,0,['Arquivo','Aba utilizada','Registros somados'],h)
        for i,row in enumerate(dados['log'],1): l.write_row(i,0,row)
        l.set_column('A:A',65);l.set_column('B:C',25)
        l.write(len(dados['log'])+3,0,'Diferença = Legado - Protheus. Ausência é sinalizada mesmo quando o total é zero.')
        l.write(len(dados['log'])+4,0,'Compara totais por filial/CFOP; não valida cada NF-e. Não deduplica nem filtra situação de notas.')
        l.write(len(dados['log'])+5,0,'Linhas com CFOP em branco ficam no grupo SEM CFOP e são listadas na aba Pendências.')
        p=w.add_worksheet('Pendências')
        p.write_row(0,0,['Arquivo','Aba','Linha original','Filial original','Valor','Motivo','Linha original completa (JSON)'],h)
        for i,row in enumerate(dados.get('pendencias',[]),1): p.write_row(i,0,row)
        p.set_column('A:A',60);p.set_column('B:D',18);p.set_column('E:E',20,money);p.set_column('F:F',24);p.set_column('G:G',80)
        p.freeze_panes(1,0)

class ConferenciaFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=theme.BG_APP)
        self.paths = [tk.StringVar(value="Nenhum arquivo selecionado"),
                      tk.StringVar(value="Nenhum arquivo selecionado")]
        self._sel = [None, None]
        self.msg = tk.StringVar(value="Selecione os dois relatórios originais de Excel.")
        self._build()

    def _build(self):
        bar = tk.Frame(self, bg=theme.PRIMARY, height=44)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="⚖️  Conferência Legado x Protheus",
                 bg=theme.PRIMARY, fg="white",
                 font=("Segoe UI Semibold", 12)).pack(side="left", padx=16, pady=10)

        body = tk.Frame(self, bg=theme.BG_APP, padx=24, pady=16)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        tk.Label(body, text="Compara os totais por filial e CFOP entre os dois sistemas.",
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")

        self._row(body, "🅿️  Protheus — Validação (Filial / Cod. Fiscal / Vlr.Total)", 0, 1)
        self._row(body, "📜  Legado — abas por filial (CFOP / TOTAL)", 1, 3)

        rodape = tk.Frame(self, bg=theme.BG_APP, padx=24)
        rodape.pack(fill="x", pady=(0, 12))
        self.btn = tk.Button(rodape, text="📊  Gerar planilha de resultado",
                             bg=theme.PRIMARY, fg="white",
                             activebackground=theme.PRIMARY_HOVER, activeforeground="white",
                             font=("Segoe UI Semibold", 11), relief="flat", bd=0,
                             cursor="hand2", padx=18, pady=7, command=self._gerar)
        self.btn.pack(side="right")
        tk.Label(rodape, textvariable=self.msg, bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9), anchor="w", justify="left",
                 wraplength=420).pack(side="left", fill="x", expand=True)

    def _row(self, parent, label, i, row):
        tk.Label(parent, text=label, bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 10)).grid(row=row, column=0, sticky="w", pady=(14, 2))
        frm = tk.Frame(parent, bg=theme.BG_APP)
        frm.grid(row=row + 1, column=0, sticky="ew")
        tk.Label(frm, textvariable=self.paths[i], bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9), anchor="w", padx=8, pady=5, wraplength=420,
                 justify="left").pack(side="left", fill="x", expand=True)
        tk.Button(frm, text="Procurar", bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 10), relief="flat", bd=0, cursor="hand2",
                  padx=10, pady=5,
                  command=lambda: self._escolher(i)).pack(side="left", padx=(6, 0))

    def _escolher(self, i):
        p = filedialog.askopenfilename(parent=self, filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self._sel[i] = p
            self.paths[i].set(os.path.basename(p))

    def _gerar(self):
        if not all(self._sel):
            messagebox.showerror("Arquivos", "Selecione os dois arquivos.", parent=self); return
        dest = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx",
                                            initialfile="Resultado_conferencia.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not dest:
            return
        src = list(self._sel)   # [Protheus, Legado]
        try:
            destino_seguro(dest, src)
        except ValueError as e:
            messagebox.showerror("Destino", str(e), parent=self); return
        self.btn.config(state="disabled")
        self.msg.set("Processando. Aguarde a confirmação…")

        def trabalho():
            try:
                dados = comparar(*src); excel(dados, dest)
                qtd = sum(r[5] != "OK" for r in dados["rows"])
                try:
                    from fiscal import historico
                    historico.registrar("Legado x Protheus", dest, f"{qtd} grupos com divergência ou pendência", **resumo(dados))
                except Exception:
                    pass   # histórico é auxiliar: nunca impede a conferência
                self.after(0, lambda: self._concluir(
                    f"Arquivo salvo: {dest}\n{qtd} grupos com divergência ou ausência.",
                    False, dest))
            except Exception as e:
                mensagem = str(e)
                self.after(0, lambda: self._concluir(mensagem, True))
        threading.Thread(target=trabalho, daemon=True).start()

    def _concluir(self, m, erro, dest=None):
        self.btn.config(state="normal")
        self.msg.set("Erro na leitura." if erro else "Conferência concluída.")
        if erro:
            messagebox.showerror("Conferência", m, parent=self)
        elif messagebox.askyesno("Conferência", m + "\n\nDeseja abrir o arquivo?", parent=self):
            try:
                os.startfile(dest)
            except Exception:
                pass


def _build_content(win):
    win.title("Conferência Legado x Protheus")
    win.geometry("660x330")
    win.minsize(560, 320)
    win.configure(bg=theme.BG_APP)
    theme.apply_theme(win)
    ConferenciaFrame(win).pack(fill="both", expand=True)


def janela():
    root = tk.Tk()
    _build_content(root)
    root.mainloop()


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--novo', help='Protheus');p.add_argument('--anterior', help='Legado');p.add_argument('--json');p.add_argument('--saida');args=p.parse_args()
    if args.novo and args.anterior:
        d=comparar(args.novo,args.anterior)
        if args.json:Path(args.json).write_text(json.dumps(d,ensure_ascii=False),encoding='utf-8')
        if args.saida:excel(d,args.saida)
        print(f'{len(d["rows"])} grupos; {sum(x[5]!="OK" for x in d["rows"])} divergências/ausências')
    else:janela()
