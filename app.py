"""
Anonimizador Juridico LGPD — App local (v4)

Interface grafica local (Tkinter, sem dependencias novas, 100% offline) para o
motor do anonimizador.py, em fluxo de duas fases:

  1. ANALISAR   — o app detecta todas as entidades do arquivo original e lista
                  categoria, dado original, nº de ocorrencias e a substituicao
                  default (gerada pelos modelos por categoria, customizaveis).
  2. REVISAR    — o usuario pode editar o texto de substituicao de cada entidade
                  (duplo clique) e desmarcar falsos positivos (clique na 1ª coluna).
  3. ANONIMIZAR — grava saida/, cofre/ (mapa de reversao) e relatorios/ (LGPD).

  RESTAURAR     — processo de regressao: devolve o documento original a partir
                  do arquivo anonimizado + mapa do cofre. Funciona tanto com os
                  tokens default quanto com substituicoes customizadas, porque o
                  cofre guarda o texto efetivamente gravado no documento.

Preferencias (modelos de token por categoria) ficam em app_prefs.json.

Uso: duplo clique em "Anonimizador.bat" ou `pythonw app.py`.
"""
import json
import os
import re
import sys
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

import anonimizador as motor

BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_SAIDA = os.path.join(BASE, 'saida')
PASTA_COFRE = os.path.join(BASE, 'cofre')
PASTA_RELAT = os.path.join(BASE, 'relatorios')
ARQ_PREFS = os.path.join(BASE, 'app_prefs.json')

CATEGORIAS = ['NOME', 'DOCUMENTO', 'PROCESSO', 'CONTATO', 'ENDERECO',
              'DADOS_BANCARIOS', 'DADO_VEICULO', 'DADO_SENSIVEL', 'VALOR', 'DATA']

# Identidade visual — paleta neutra
COR_PRIMARIA = '#2F3B4C'   # grafite azulado (cabecalho, botoes)
COR_DESTAQUE = '#7D9CB8'   # azul aco (titulo, estados ativos)
COR_PAINEL = '#DDE3E8'     # cinza claro (barra de status)
COR_FUNDO = '#F4F6F8'      # fundo geral
COR_TEXTO = '#22262A'


def _fonte(preferidas, tamanho, peso='normal'):
    disponiveis = set(tkfont.families())
    for nome in preferidas:
        if nome in disponiveis:
            return (nome, tamanho, peso)
    return ('Segoe UI', tamanho, peso)


def carregar_prefs():
    try:
        with open(ARQ_PREFS, encoding='utf-8') as f:
            prefs = json.load(f)
    except (OSError, ValueError):
        prefs = {}
    prefs.setdefault('modelos_token', {})
    return prefs


def salvar_prefs(prefs):
    with open(ARQ_PREFS, 'w', encoding='utf-8') as f:
        json.dump(prefs, f, ensure_ascii=False, indent=2)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Anonimizador Jurídico LGPD (100% local)')
        self.geometry('1180x680')
        self.minsize(940, 540)
        self.configure(bg=COR_FUNDO)

        self.prefs = carregar_prefs()
        self.arquivo = None
        self.entidades = []          # resultado de analisar_arquivo
        self._editor = None          # Entry flutuante de edicao de celula

        self.fonte_titulo = _fonte(['Segoe UI Semibold', 'Segoe UI'], 15, 'bold')
        self.fonte_sub = _fonte(['Segoe UI'], 9)
        self.fonte_corpo = _fonte(['Segoe UI'], 10)

        self._montar_estilo()
        self._montar_layout()

    # ------------------------------------------------------------------ UI
    def _montar_estilo(self):
        st = ttk.Style(self)
        st.theme_use('clam')
        st.configure('TFrame', background=COR_FUNDO)
        st.configure('Header.TFrame', background=COR_PRIMARIA)
        st.configure('Header.TLabel', background=COR_PRIMARIA, foreground=COR_DESTAQUE,
                     font=self.fonte_titulo)
        st.configure('HeaderSub.TLabel', background=COR_PRIMARIA, foreground=COR_FUNDO,
                     font=self.fonte_sub)
        st.configure('TLabel', background=COR_FUNDO, foreground=COR_TEXTO,
                     font=self.fonte_corpo)
        st.configure('Status.TLabel', background=COR_PAINEL, foreground=COR_TEXTO,
                     font=self.fonte_corpo)
        st.configure('TButton', font=self.fonte_corpo, padding=(10, 5),
                     background=COR_PRIMARIA, foreground=COR_FUNDO, borderwidth=0)
        st.map('TButton',
               background=[('active', COR_DESTAQUE), ('disabled', COR_PAINEL)],
               foreground=[('active', COR_PRIMARIA), ('disabled', COR_FUNDO)])
        st.configure('Acao.TButton', font=(self.fonte_corpo[0], 10, 'bold'))
        st.configure('Treeview', font=self.fonte_corpo, rowheight=26,
                     background='white', fieldbackground='white', foreground=COR_TEXTO)
        st.configure('Treeview.Heading', font=(self.fonte_corpo[0], 9, 'bold'),
                     background=COR_PRIMARIA, foreground=COR_FUNDO)
        st.map('Treeview.Heading', background=[('active', COR_PRIMARIA)])
        st.map('Treeview', background=[('selected', COR_DESTAQUE)],
               foreground=[('selected', 'white')])

    def _montar_layout(self):
        # Cabecalho institucional
        header = ttk.Frame(self, style='Header.TFrame', padding=(18, 12))
        header.pack(fill='x')
        ttk.Label(header, text='Anonimizador Jurídico LGPD', style='Header.TLabel').pack(anchor='w')
        ttk.Label(header, text='Processamento 100% local · ferramenta de apoio — a revisão humana '
                               'das entidades detectadas é indispensável',
                  style='HeaderSub.TLabel').pack(anchor='w')

        # Barra de acoes
        barra = ttk.Frame(self, padding=(14, 10))
        barra.pack(fill='x')
        ttk.Button(barra, text='Selecionar arquivo…', command=self.selecionar_arquivo).pack(side='left')
        self.lbl_arquivo = ttk.Label(barra, text='Nenhum arquivo selecionado', width=52)
        self.lbl_arquivo.pack(side='left', padx=(10, 14))
        self.btn_analisar = ttk.Button(barra, text='1 · Analisar', style='Acao.TButton',
                                       command=self.analisar, state='disabled')
        self.btn_analisar.pack(side='left', padx=3)
        self.btn_aplicar = ttk.Button(barra, text='2 · Anonimizar e salvar', style='Acao.TButton',
                                      command=self.anonimizar, state='disabled')
        self.btn_aplicar.pack(side='left', padx=3)
        ttk.Label(barra, text='Saída do PDF:').pack(side='left', padx=(14, 4))
        self.var_formato_pdf = tk.StringVar(
            value=self.prefs.get('formato_saida_pdf', 'PDF'))
        combo_pdf = ttk.Combobox(barra, textvariable=self.var_formato_pdf,
                                 values=('PDF', 'TXT', 'DOCX'), width=6,
                                 state='readonly', font=self.fonte_corpo)
        combo_pdf.pack(side='left')
        combo_pdf.bind('<<ComboboxSelected>>', self._salvar_formato_pdf)
        ttk.Button(barra, text='Restaurar…', command=self.restaurar).pack(side='right', padx=3)
        ttk.Button(barra, text='Modelos de token…', command=self.editar_modelos).pack(side='right', padx=3)
        ttk.Button(barra, text='Config. LGPD…', command=self.editar_config).pack(side='right', padx=3)

        # Dica de uso
        ttk.Label(self, padding=(16, 0),
                  text='Revisão: clique na coluna “Usar” para incluir/excluir · duplo clique em '
                       '“Substituição” para editar o texto (a reversão funciona com qualquer texto).'
                  ).pack(fill='x')

        # Tabela de entidades
        quadro = ttk.Frame(self, padding=(14, 8))
        quadro.pack(fill='both', expand=True)
        cols = ('usar', 'categoria', 'ocorr', 'dado', 'subst')
        self.tree = ttk.Treeview(quadro, columns=cols, show='headings', selectmode='browse')
        self.tree.heading('usar', text='Usar')
        self.tree.heading('categoria', text='Categoria')
        self.tree.heading('ocorr', text='Ocorrências')
        self.tree.heading('dado', text='Dado original')
        self.tree.heading('subst', text='Substituição (editável)')
        self.tree.column('usar', width=52, anchor='center', stretch=False)
        self.tree.column('categoria', width=130, anchor='w', stretch=False)
        self.tree.column('ocorr', width=90, anchor='center', stretch=False)
        self.tree.column('dado', width=430, anchor='w')
        self.tree.column('subst', width=280, anchor='w')
        vsb = ttk.Scrollbar(quadro, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.tree.tag_configure('excluido', foreground='#9A9A90')
        self.tree.tag_configure('custom', foreground=COR_PRIMARIA)
        self.tree.bind('<Button-1>', self._clique_tree)
        self.tree.bind('<Double-1>', self._duplo_clique_tree)

        # Barra de status
        self.status = ttk.Label(self, text='Pronto.', style='Status.TLabel', padding=(14, 6))
        self.status.pack(fill='x', side='bottom')

    def _set_status(self, txt):
        self.status.config(text=txt)
        self.update_idletasks()

    def _salvar_formato_pdf(self, _=None):
        self.prefs['formato_saida_pdf'] = self.var_formato_pdf.get()
        salvar_prefs(self.prefs)

    # ------------------------------------------------------------ acoes
    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title='Selecione o documento original',
            initialdir=os.path.join(BASE, 'entrada'),
            filetypes=[('Documentos suportados', '*.docx *.txt *.pdf'),
                       ('Word', '*.docx'), ('Texto', '*.txt'), ('PDF', '*.pdf')])
        if not caminho:
            return
        self.arquivo = caminho
        self.lbl_arquivo.config(text=os.path.basename(caminho))
        self.entidades = []
        self.tree.delete(*self.tree.get_children())
        self.btn_analisar.config(state='normal')
        self.btn_aplicar.config(state='disabled')
        self._set_status('Arquivo selecionado. Clique em “1 · Analisar”.')

    def analisar(self):
        self._fechar_editor()
        try:
            self.config(cursor='watch')
            self._set_status('Analisando documento…')
            self.entidades = motor.analisar_arquivo(
                self.arquivo, modelos=self.prefs.get('modelos_token'))
        except Exception as e:
            messagebox.showerror('Erro na análise', str(e) + '\n\n' + traceback.format_exc(limit=3))
            self._set_status('Erro na análise.')
            return
        finally:
            self.config(cursor='')

        self.tree.delete(*self.tree.get_children())
        for ent in self.entidades:
            ent['usar'] = True
            ent['subst'] = ent['token']
            self.tree.insert('', 'end', iid=ent['token'], values=(
                '✔', ent['categoria'], ent['ocorrencias'], ent['dado'], ent['subst']))
        total = sum(e['ocorrencias'] for e in self.entidades)
        self.btn_aplicar.config(state='normal' if self.entidades else 'disabled')
        self._set_status('%d entidade(s) única(s), %d ocorrência(s). Revise e clique em '
                         '“2 · Anonimizar e salvar”.' % (len(self.entidades), total))

    def anonimizar(self):
        self._fechar_editor()
        if not self.entidades:
            return
        # Config LGPD precisa estar preenchida (log de compliance)
        campos, incompletos = motor.ler_config(BASE)
        if incompletos:
            messagebox.showwarning(
                'Configuração LGPD',
                'Preencha os metadados LGPD antes de anonimizar:\n· ' + '\n· '.join(incompletos))
            self.editar_config()
            campos, incompletos = motor.ler_config(BASE)
            if incompletos:
                return

        substituicoes = {e['token']: e['subst'] for e in self.entidades
                         if e['usar'] and e['subst'] != e['token']}
        excluidos = {e['token'] for e in self.entidades if not e['usar']}

        # Textos repetidos sao permitidos, mas tornam esses trechos IRREVERSIVEIS
        # (ficam fora do mapa de reversao). Pede confirmacao explicita.
        usados, repetidos = {}, {}
        for e in self.entidades:
            if not e['usar']:
                continue
            if e['subst'] in usados:
                repetidos.setdefault(e['subst'], 0)
                repetidos[e['subst']] += 1
            usados[e['subst']] = e['dado']
        permitir_ambiguas = False
        if repetidos:
            linhas = '\n'.join('· “%s” — usado para %d entidades'
                               % (t, n + 1) for t, n in sorted(repetidos.items()))
            if not messagebox.askyesno(
                    'Substituições repetidas — regressão parcial',
                    'Os textos abaixo foram usados para mais de um dado diferente:\n\n'
                    '%s\n\n'
                    'ATENÇÃO: esses trechos NÃO poderão ser restaurados — a função '
                    '“Restaurar” os manterá anonimizados (o restante do documento '
                    'continua reversível normalmente).\n\nContinuar mesmo assim?'
                    % linhas):
                return
            permitir_ambiguas = True

        os.makedirs(PASTA_SAIDA, exist_ok=True)
        os.makedirs(PASTA_COFRE, exist_ok=True)
        os.makedirs(PASTA_RELAT, exist_ok=True)
        nome = os.path.basename(self.arquivo)
        base_nome, ext = os.path.splitext(nome)
        ext_saida = ext
        if ext.lower() == '.pdf':
            ext_saida = '.' + self.var_formato_pdf.get().lower()
        caminho_saida = os.path.join(PASTA_SAIDA, base_nome + '_anonimizado' + ext_saida)
        caminho_mapa = os.path.join(PASTA_COFRE, base_nome + '_mapa.json')

        try:
            self.config(cursor='watch')
            self._set_status('Anonimizando…')
            rel = motor.aplicar_anonimizacao(
                self.arquivo, caminho_saida, caminho_mapa,
                substituicoes=substituicoes, excluidos=excluidos,
                modelos=self.prefs.get('modelos_token'), compliance=campos,
                permitir_ambiguas=permitir_ambiguas)
            motor.registrar_compliance(rel, PASTA_RELAT, base_nome)
        except Exception as e:
            messagebox.showerror('Erro ao anonimizar', str(e))
            self._set_status('Erro ao anonimizar.')
            return
        finally:
            self.config(cursor='')

        msg = ('Documento anonimizado com sucesso.\n\n'
               'Substituições: %d (%d entidades)\n'
               'Customizadas: %d · Excluídas: %d\n\n'
               'Saída:  %s\nCofre:  %s  [CONFIDENCIAL]\n\nAbrir a pasta de saída?'
               % (rel['total_substituicoes'], rel['entidades_unicas'],
                  rel['substituicoes_customizadas'], rel['entidades_excluidas'],
                  caminho_saida, caminho_mapa))
        if rel.get('avisos'):
            msg = 'ATENÇÃO:\n· ' + '\n· '.join(rel['avisos']) + '\n\n' + msg
        self._set_status('Concluído: %s' % os.path.basename(caminho_saida))
        if messagebox.askyesno('Concluído', msg):
            os.startfile(PASTA_SAIDA)

    def restaurar(self):
        """Processo de regressão: reverte um arquivo anonimizado usando o mapa
        do cofre — com tokens default ou customizados. Aceita também um
        documento DIFERENTE do original (ex.: parecer devolvido por IA) que
        contenha os mesmos tokens: se o mapa não for localizado pelo nome do
        arquivo, o usuário escolhe qual mapa do cofre aplicar."""
        self._fechar_editor()
        caminho = filedialog.askopenfilename(
            title='Selecione o arquivo ANONIMIZADO (ou derivado dele) a restaurar',
            initialdir=PASTA_SAIDA,
            filetypes=[('Documentos suportados', '*.docx *.txt *.pdf')])
        if not caminho:
            return

        # Localiza o mapa pelo nome; se não houver, pede para escolher no cofre.
        base_nome = os.path.basename(caminho)
        nome_base = os.path.splitext(re.sub(r'_anonimizado(\.[^.]+)$', r'\1', base_nome))[0]
        caminho_mapa = os.path.join(PASTA_COFRE, nome_base + '_mapa.json')
        if not os.path.exists(caminho_mapa):
            if not messagebox.askyesno(
                    'Mapa não localizado pelo nome',
                    'Não encontrei "%s" no cofre.\n\n'
                    'Se este documento foi gerado a partir de um arquivo anonimizado '
                    '(ex.: minuta revisada por IA) e mantém os mesmos tokens, você pode '
                    'escolher manualmente o mapa de reversão correspondente.\n\n'
                    'Escolher o mapa agora?' % (nome_base + '_mapa.json')):
                return
            caminho_mapa = filedialog.askopenfilename(
                title='Selecione o mapa de reversão (cofre)',
                initialdir=PASTA_COFRE,
                filetypes=[('Mapa de reversão', '*_mapa.json'), ('JSON', '*.json')])
            if not caminho_mapa:
                return

        try:
            self.config(cursor='watch')
            self._set_status('Restaurando…')
            saida, revertidos = motor.restaurar_arquivo(caminho, PASTA_COFRE,
                                                        caminho_mapa=caminho_mapa)
        except Exception as e:
            messagebox.showerror('Erro ao restaurar', str(e))
            self._set_status('Erro ao restaurar.')
            return
        finally:
            self.config(cursor='')
        self._set_status('Restaurado: %s (%d trechos revertidos)'
                         % (os.path.basename(saida), revertidos))
        msg = ('Arquivo restaurado:\n%s\n\nTrechos revertidos: %d\nMapa: %s\n\nAbrir a pasta?'
               % (saida, revertidos, os.path.basename(caminho_mapa)))
        if revertidos == 0:
            msg = ('ATENÇÃO: nenhum token do mapa foi encontrado no documento — '
                   'confira se escolheu o mapa correto.\n\n') + msg
        if messagebox.askyesno('Restaurado', msg):
            os.startfile(os.path.dirname(saida))

    # ------------------------------------------------- edicao na tabela
    def _ent_por_iid(self, iid):
        for e in self.entidades:
            if e['token'] == iid:
                return e
        return None

    def _clique_tree(self, evento):
        if self.tree.identify('region', evento.x, evento.y) != 'cell':
            return
        if self.tree.identify_column(evento.x) != '#1':
            return
        iid = self.tree.identify_row(evento.y)
        ent = self._ent_por_iid(iid)
        if not ent:
            return
        ent['usar'] = not ent['usar']
        self._atualizar_linha(ent)

    def _duplo_clique_tree(self, evento):
        if self.tree.identify('region', evento.x, evento.y) != 'cell':
            return
        if self.tree.identify_column(evento.x) != '#5':
            return
        iid = self.tree.identify_row(evento.y)
        ent = self._ent_por_iid(iid)
        if not ent:
            return
        self._fechar_editor()
        x, y, w, h = self.tree.bbox(iid, '#5')
        editor = tk.Entry(self.tree, font=self.fonte_corpo, relief='solid',
                          borderwidth=1, background='white')
        editor.insert(0, ent['subst'])
        editor.select_range(0, 'end')
        editor.place(x=x, y=y, width=w, height=h)
        editor.focus_set()

        def confirmar(_=None):
            novo = editor.get().strip()
            if novo:
                ent['subst'] = novo
            self._fechar_editor()
            self._atualizar_linha(ent)

        def cancelar(_=None):
            self._fechar_editor()

        editor.bind('<Return>', confirmar)
        editor.bind('<Escape>', cancelar)
        editor.bind('<FocusOut>', confirmar)
        self._editor = editor

    def _fechar_editor(self):
        if self._editor is not None:
            self._editor.destroy()
            self._editor = None

    def _atualizar_linha(self, ent):
        tags = ()
        if not ent['usar']:
            tags = ('excluido',)
        elif ent['subst'] != ent['token']:
            tags = ('custom',)
        self.tree.item(ent['token'], values=(
            '✔' if ent['usar'] else '—', ent['categoria'], ent['ocorrencias'],
            ent['dado'], ent['subst']), tags=tags)

    # --------------------------------------------------------- dialogos
    def editar_modelos(self):
        """Modelos default de substituicao por categoria (persistidos)."""
        dlg = tk.Toplevel(self)
        dlg.title('Modelos de token por categoria')
        dlg.configure(bg=COR_FUNDO)
        dlg.transient(self)
        dlg.grab_set()
        ttk.Label(dlg, padding=(12, 10),
                  text='Modelo default da substituição por categoria.\n'
                       'Placeholders: {cat} = categoria · {n} ou {n:02d} = número sequencial.\n'
                       'Exemplos: [{cat}_{n:02d}]  ·  PARTE {n}  ·  <<NOME {n}>>\n'
                       'Vale para a PRÓXIMA análise. A reversão funciona com qualquer modelo.'
                  ).grid(row=0, column=0, columnspan=2, sticky='w')
        entradas = {}
        modelos = self.prefs.get('modelos_token', {})
        for i, cat in enumerate(CATEGORIAS, start=1):
            ttk.Label(dlg, text=cat, padding=(12, 2)).grid(row=i, column=0, sticky='w')
            var = tk.StringVar(value=modelos.get(cat, motor.Registro.MODELO_PADRAO))
            ent = ttk.Entry(dlg, textvariable=var, width=36, font=self.fonte_corpo)
            ent.grid(row=i, column=1, sticky='we', padx=(0, 12), pady=2)
            entradas[cat] = var
        rodape = ttk.Frame(dlg, padding=12)
        rodape.grid(row=len(CATEGORIAS) + 1, column=0, columnspan=2, sticky='e')

        def salvar():
            novos = {}
            for cat, var in entradas.items():
                valor = var.get().strip()
                if not valor:
                    continue
                if '{n' not in valor:
                    messagebox.showerror(
                        'Modelo inválido',
                        'O modelo de %s precisa conter {n} (número sequencial), '
                        'senão entidades diferentes receberiam o mesmo texto.' % cat,
                        parent=dlg)
                    return
                if valor != motor.Registro.MODELO_PADRAO:
                    novos[cat] = valor
            self.prefs['modelos_token'] = novos
            salvar_prefs(self.prefs)
            dlg.destroy()
            self._set_status('Modelos salvos. Analise novamente para aplicá-los.')

        ttk.Button(rodape, text='Restaurar padrão', command=lambda: [
            v.set(motor.Registro.MODELO_PADRAO) for v in entradas.values()]).pack(side='left', padx=4)
        ttk.Button(rodape, text='Salvar', style='Acao.TButton', command=salvar).pack(side='left', padx=4)
        ttk.Button(rodape, text='Cancelar', command=dlg.destroy).pack(side='left', padx=4)

    def editar_config(self):
        """Metadados LGPD (config.json) — finalidade, destinatario etc."""
        campos, _ = motor.ler_config(BASE)
        dlg = tk.Toplevel(self)
        dlg.title('Configuração LGPD (art. 37 e art. 6º, VIII)')
        dlg.configure(bg=COR_FUNDO)
        dlg.transient(self)
        dlg.grab_set()
        rotulos = {
            'finalidade': 'Finalidade do tratamento',
            'destinatario': 'Destinatário do arquivo anonimizado',
            'base_legal': 'Base legal (LGPD)',
            'categoria_titular': 'Categoria dos titulares',
        }
        entradas = {}
        for i, (campo, rotulo) in enumerate(rotulos.items()):
            ttk.Label(dlg, text=rotulo, padding=(12, 4)).grid(row=i, column=0, sticky='w')
            valor = campos.get(campo, '')
            var = tk.StringVar(value='' if valor.startswith('PREENCHER') else valor)
            ttk.Entry(dlg, textvariable=var, width=64, font=self.fonte_corpo).grid(
                row=i, column=1, sticky='we', padx=(0, 12), pady=3)
            entradas[campo] = var
        rodape = ttk.Frame(dlg, padding=12)
        rodape.grid(row=len(rotulos), column=0, columnspan=2, sticky='e')

        def salvar():
            novos = {c: v.get().strip() for c, v in entradas.items()}
            vazios = [rotulos[c] for c, v in novos.items() if not v]
            if vazios:
                messagebox.showwarning('Campos vazios',
                                       'Preencha: ' + ', '.join(vazios), parent=dlg)
                return
            motor.salvar_config(BASE, novos)
            dlg.destroy()
            self._set_status('Configuração LGPD salva.')

        ttk.Button(rodape, text='Salvar', style='Acao.TButton', command=salvar).pack(side='left', padx=4)
        ttk.Button(rodape, text='Cancelar', command=dlg.destroy).pack(side='left', padx=4)


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
