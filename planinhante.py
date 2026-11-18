import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib

# FORÇA O MATPLOTLIB A USAR O TKINTER (Corrige travamentos ao abrir)
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd


class GeradorGraficosApp:

    def __init__(self, root):
        self.root = root
        self.root.title('Gerador de Gráficos de Planilhas')
        self.root.geometry('1100x760')
        self.root.configure(padx=15, pady=15)

        # Garante que a janela abra na frente de outros programas
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(self.root.attributes, '-topmost', False)

        self.df = None
        self.figura_atual = None
        self.canvas = None
        self.dados_agrupados = None
        self.janela_grafico = None
        self.frame_grafico = None

        self.criar_interface()

    def criar_interface(self):
        frame_arquivo = ttk.LabelFrame(
            self.root, text='1. Arquivo de Dados', padding=10
        )
        frame_arquivo.pack(fill='x', pady=(0, 10))

        btn_selecionar = ttk.Button(
            frame_arquivo,
            text='Selecionar Planilha (.xlsx, .csv)',
            command=self.selecionar_arquivo,
        )
        btn_selecionar.pack(side='left')

        self.lbl_arquivo = ttk.Label(
            frame_arquivo,
            text='Nenhum arquivo selecionado',
            foreground='gray',
        )
        self.lbl_arquivo.pack(side='left', padx=15)

        frame_controles = ttk.LabelFrame(
            self.root, text='2. Configurações do Gráfico', padding=10
        )
        frame_controles.pack(fill='x', pady=(0, 10))

        ttk.Label(frame_controles, text='Categoria (Texto):').grid(
            row=0, column=0, padx=5, pady=5, sticky='w'
        )
        self.combo_categoria = ttk.Combobox(
            frame_controles, state='disabled', width=22
        )
        self.combo_categoria.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_controles, text='Valor (Numérico):').grid(
            row=0, column=2, padx=5, pady=5, sticky='w'
        )
        self.combo_valor = ttk.Combobox(
            frame_controles, state='disabled', width=22
        )
        self.combo_valor.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_controles, text='Tipo de Gráfico:').grid(
            row=0, column=4, padx=5, pady=5, sticky='w'
        )
        self.combo_tipo = ttk.Combobox(
            frame_controles,
            values=[
                'Barras Verticais',
                'Barras Horizontais',
                'Pizza',
                'Linhas',
                'Dispersão',
            ],
            state='readonly',
            width=18,
        )
        self.combo_tipo.current(0)
        self.combo_tipo.grid(row=0, column=5, padx=5, pady=5)

        ttk.Label(frame_controles, text='Agregação:').grid(
            row=1, column=0, padx=5, pady=5, sticky='w'
        )
        self.combo_agregacao = ttk.Combobox(
            frame_controles,
            values=['Soma', 'Média', 'Contagem', 'Máximo', 'Mínimo'],
            state='readonly',
            width=22,
        )
        self.combo_agregacao.current(0)
        self.combo_agregacao.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_controles, text='Mostrar:').grid(
            row=1, column=2, padx=5, pady=5, sticky='w'
        )
        self.spin_top = tk.Spinbox(
            frame_controles, from_=0, to=1000, width=8
        )
        self.spin_top.delete(0, 'end')
        self.spin_top.insert(0, '20')
        self.spin_top.grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(frame_controles, text='Título:').grid(
            row=1, column=4, padx=5, pady=5, sticky='w'
        )
        self.entry_titulo = ttk.Entry(frame_controles, width=30)
        self.entry_titulo.grid(row=1, column=5, padx=5, pady=5)

        frame_filtros = ttk.LabelFrame(
            self.root, text='Filtros opcionais (aplicados ao gerar)', padding=8
        )
        frame_filtros.pack(fill='x', pady=(0, 10))
        ttk.Label(frame_filtros, text='Categoria contém:').pack(
            side='left', padx=(0, 5)
        )
        self.entry_filtro_texto = ttk.Entry(frame_filtros, width=20)
        self.entry_filtro_texto.pack(side='left', padx=(0, 12))
        ttk.Label(frame_filtros, text='Valor mínimo:').pack(
            side='left', padx=(0, 5)
        )
        self.entry_valor_min = ttk.Entry(frame_filtros, width=12)
        self.entry_valor_min.pack(side='left', padx=(0, 12))
        ttk.Label(frame_filtros, text='Valor máximo:').pack(
            side='left', padx=(0, 5)
        )
        self.entry_valor_max = ttk.Entry(frame_filtros, width=12)
        self.entry_valor_max.pack(side='left', padx=(0, 12))
        ttk.Button(
            frame_filtros, text='Limpar filtros', command=self.limpar_filtros
        ).pack(side='left')

        frame_preview = ttk.LabelFrame(
            self.root, text='Prévia dos Dados', padding=5
        )
        frame_preview.pack(fill='x', pady=(0, 10))

        self.preview = ttk.Treeview(frame_preview, show='headings', height=5)
        preview_y = ttk.Scrollbar(
            frame_preview, orient='vertical', command=self.preview.yview
        )
        preview_x = ttk.Scrollbar(
            frame_preview, orient='horizontal', command=self.preview.xview
        )
        self.preview.configure(
            yscrollcommand=preview_y.set, xscrollcommand=preview_x.set
        )
        self.preview.grid(row=0, column=0, sticky='nsew')
        preview_y.grid(row=0, column=1, sticky='ns')
        preview_x.grid(row=1, column=0, sticky='ew')
        frame_preview.rowconfigure(0, weight=1)
        frame_preview.columnconfigure(0, weight=1)

        frame_botoes = ttk.Frame(self.root)
        frame_botoes.pack(fill='x', pady=(0, 10))

        self.btn_gerar = ttk.Button(
            frame_botoes,
            text='Gerar Gráfico',
            command=self.gerar_grafico,
            state='disabled',
        )
        self.btn_gerar.pack(side='left', padx=(0, 10))

        self.btn_abrir_grafico = ttk.Button(
            frame_botoes,
            text='Abrir visualização',
            command=self.mostrar_grafico,
            state='disabled',
        )
        self.btn_abrir_grafico.pack(side='left', padx=(0, 10))

        self.btn_salvar = ttk.Button(
            frame_botoes,
            text='Baixar Imagem do Gráfico',
            command=self.salvar_imagem,
            state='disabled',
        )
        self.btn_salvar.pack(side='left')

        self.btn_exportar = ttk.Button(
            frame_botoes,
            text='Exportar Dados (.csv)',
            command=self.exportar_dados,
            state='disabled',
        )
        self.btn_exportar.pack(side='left', padx=(10, 0))

        self.lbl_status = ttk.Label(
            frame_botoes, text='', foreground='gray'
        )
        self.lbl_status.pack(side='right')

        frame_resumo = ttk.LabelFrame(
            self.root, text='Resumo dos dados filtrados', padding=8
        )
        frame_resumo.pack(fill='x', pady=(0, 10))
        self.lbl_resumo = ttk.Label(
            frame_resumo,
            text='Carregue uma planilha e gere um gráfico para ver o resumo.',
            justify='left',
        )
        self.lbl_resumo.pack(fill='x')

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title='Selecione uma planilha',
            filetypes=[
                ('Planilhas Excel e CSV', '*.xlsx *.xls *.csv'),
                ('Excel', '*.xlsx *.xls'),
                ('CSV', '*.csv'),
            ],
        )
        if not caminho:
            return

        try:
            extensao = os.path.splitext(caminho)[1].lower()
            if extensao in ['.xlsx', '.xls']:
                self.df = pd.read_excel(caminho)
            elif extensao == '.csv':
                self.df = self._ler_csv(caminho)
            else:
                raise ValueError('Formato de arquivo não suportado.')

            nome_arquivo = os.path.basename(caminho)
            self.lbl_arquivo.config(
                text=f'Carregado: {nome_arquivo}', foreground='green'
            )

            colunas = list(self.df.columns)
            self.combo_categoria.config(values=colunas, state='readonly')
            self.combo_valor.config(values=colunas, state='readonly')

            if len(colunas) >= 2:
                coluna_categoria = next(
                    (c for c in colunas if not self._eh_numerica(c)),
                    colunas[0],
                )
                coluna_valor = next(
                    (c for c in colunas if self._eh_numerica(c)),
                    colunas[min(1, len(colunas) - 1)],
                )
                self.combo_categoria.set(coluna_categoria)
                self.combo_valor.set(coluna_valor)

            self.btn_gerar.config(state='normal')
            self.btn_exportar.config(state='disabled')
            self._atualizar_preview()
            self.lbl_status.config(
                text=f'{len(self.df):,} linhas · {len(colunas)} colunas'
            )

        except Exception as e:
            messagebox.showerror(
                'Erro ao ler arquivo',
                f'Não foi possível abrir a planilha:\n{e}',
            )

    def _ler_csv(self, caminho):
        with open(caminho, 'rb') as arquivo:
            amostra_bytes = arquivo.read(8192)
        amostra = amostra_bytes.decode('utf-8-sig', errors='replace')
        try:
            dialecto = csv.Sniffer().sniff(amostra, delimiters=',;\t|')
            separador = dialecto.delimiter
        except csv.Error:
            separador = ';' if amostra.count(';') > amostra.count(',') else ','

        for encoding in ('utf-8-sig', 'cp1252', 'latin1'):
            try:
                return pd.read_csv(
                    caminho,
                    sep=separador,
                    encoding=encoding,
                    on_bad_lines='skip',
                )
            except UnicodeDecodeError:
                continue
        raise ValueError('Não foi possível decodificar o arquivo CSV.')

    def _eh_numerica(self, coluna):
        if self.df is None or coluna not in self.df:
            return False
        valores = self.df[coluna].dropna()
        if valores.empty:
            return False
        if pd.api.types.is_numeric_dtype(valores):
            return True
        convertidos = valores.astype(str).map(self._converter_numero)
        return convertidos.notna().sum() >= max(1, len(valores) * 0.6)

    def _converter_numero(self, valor):
        if pd.isna(valor):
            return None
        texto = str(valor).strip().replace('\u00a0', '').replace(' ', '')
        texto = texto.replace('R$', '').replace('€', '').replace('$', '')
        if ',' in texto and '.' in texto:
            if texto.rfind(',') > texto.rfind('.'):
                texto = texto.replace('.', '').replace(',', '.')
            else:
                texto = texto.replace(',', '')
        elif ',' in texto:
            texto = texto.replace('.', '').replace(',', '.')
        try:
            numero = float(texto)
            return numero if pd.notna(numero) else None
        except (TypeError, ValueError):
            return None

    def _atualizar_preview(self):
        for item in self.preview.get_children():
            self.preview.delete(item)
        if self.df is None:
            return
        colunas = list(self.df.columns)
        self.preview['columns'] = colunas
        for coluna in colunas:
            self.preview.heading(coluna, text=str(coluna))
            self.preview.column(coluna, width=120, anchor='w')
        for _, linha in self.df.head(100).iterrows():
            valores = [self._formatar_celula(v) for v in linha.tolist()]
            self.preview.insert('', 'end', values=valores)

    def _formatar_celula(self, valor):
        if pd.isna(valor):
            return ''
        if isinstance(valor, float):
            return f'{valor:,.2f}'
        return str(valor)

    def limpar_filtros(self):
        for campo in (
            self.entry_filtro_texto,
            self.entry_valor_min,
            self.entry_valor_max,
        ):
            campo.delete(0, 'end')

    def _obter_dados_filtrados(self, col_cat, col_val):
        df_temp = self.df.copy()
        df_temp[col_val] = df_temp[col_val].map(self._converter_numero)
        df_temp = df_temp.dropna(subset=[col_cat, col_val])
        df_temp[col_cat] = df_temp[col_cat].astype(str)

        texto_filtro = self.entry_filtro_texto.get().strip().casefold()
        if texto_filtro:
            corresponde = df_temp[col_cat].str.casefold().str.contains(
                texto_filtro, regex=False, na=False
            )
            df_temp = df_temp[corresponde]

        minimo_texto = self.entry_valor_min.get().strip()
        maximo_texto = self.entry_valor_max.get().strip()
        minimo = self._converter_numero(minimo_texto) if minimo_texto else None
        maximo = self._converter_numero(maximo_texto) if maximo_texto else None
        if minimo_texto and minimo is None:
            raise ValueError('O valor mínimo informado não é numérico.')
        if maximo_texto and maximo is None:
            raise ValueError('O valor máximo informado não é numérico.')
        if minimo is not None and maximo is not None and minimo > maximo:
            raise ValueError('O valor mínimo não pode ser maior que o máximo.')
        if minimo is not None:
            df_temp = df_temp[df_temp[col_val] >= minimo]
        if maximo is not None:
            df_temp = df_temp[df_temp[col_val] <= maximo]
        return df_temp

    def _obter_frame_grafico(self):
        if self.janela_grafico is None or not self.janela_grafico.winfo_exists():
            self.janela_grafico = tk.Toplevel(self.root)
            self.janela_grafico.title('Visualização do gráfico')
            self.janela_grafico.geometry('1000x650')
            self.janela_grafico.minsize(600, 400)
            self.janela_grafico.protocol(
                'WM_DELETE_WINDOW', self._fechar_janela_grafico
            )
            self.frame_grafico = ttk.Frame(
                self.janela_grafico, padding=10
            )
            self.frame_grafico.pack(fill='both', expand=True)
        return self.frame_grafico

    def _fechar_janela_grafico(self):
        if self.canvas is not None:
            self.canvas = None
        if self.janela_grafico is not None:
            self.janela_grafico.destroy()
        self.janela_grafico = None
        self.frame_grafico = None

    def mostrar_grafico(self):
        if self.figura_atual is None:
            return
        frame_grafico = self._obter_frame_grafico()
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
        self.canvas = FigureCanvasTkAgg(
            self.figura_atual, master=frame_grafico
        )
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
        self.janela_grafico.deiconify()
        self.janela_grafico.lift()

    def gerar_grafico(self):
        col_cat = self.combo_categoria.get()
        col_val = self.combo_valor.get()
        tipo = self.combo_tipo.get()

        if not col_cat or not col_val:
            messagebox.showwarning(
                'Aviso', 'Selecione as colunas de categoria e valor.'
            )
            return

        try:
            df_temp = self._obter_dados_filtrados(col_cat, col_val)
            if df_temp.empty:
                messagebox.showwarning(
                    'Sem resultados',
                    'Nenhuma linha corresponde aos filtros selecionados.',
                )
                return

            agregacao = self.combo_agregacao.get()
            agrupado = df_temp.groupby(col_cat)[col_val]
            if agregacao == 'Média':
                dados_agrupados = agrupado.mean()
            elif agregacao == 'Contagem':
                dados_agrupados = agrupado.count()
            elif agregacao == 'Máximo':
                dados_agrupados = agrupado.max()
            elif agregacao == 'Mínimo':
                dados_agrupados = agrupado.min()
            else:
                dados_agrupados = agrupado.sum()

            dados_agrupados = dados_agrupados.sort_values(ascending=False)

            if dados_agrupados.empty:
                messagebox.showwarning(
                    'Dados inválidos',
                    'A coluna de valores não contém números válidos.',
                )
                return

            try:
                top_n = max(0, int(self.spin_top.get()))
            except (TypeError, ValueError):
                top_n = 20
            if top_n > 0 and len(dados_agrupados) > top_n:
                principais = dados_agrupados.iloc[:top_n]
                if agregacao == 'Soma':
                    resto = dados_agrupados.iloc[top_n:].sum()
                    dados_agrupados = pd.concat(
                        [principais, pd.Series({'Outros': resto})]
                    )
                else:
                    dados_agrupados = principais
            self.dados_agrupados = dados_agrupados

            if self.figura_atual is not None:
                plt.close(self.figura_atual)
            fig, ax = plt.subplots(figsize=(8, 4.5), dpi=100)

            if tipo == 'Barras Verticais':
                barras = ax.bar(
                    dados_agrupados.index.astype(str),
                    dados_agrupados.values,
                    color='#3B82F6',
                )
                titulo = self.entry_titulo.get().strip()
                ax.set_title(titulo or f'{agregacao} de {col_val} por {col_cat}')
                ax.set_xlabel(col_cat)
                ax.set_ylabel(col_val)
                ax.tick_params(axis='x', rotation=35)
                ax.grid(axis='y', linestyle='--', alpha=0.6)

                for barra in barras:
                    h = barra.get_height()
                    ax.text(
                        barra.get_x() + barra.get_width() / 2,
                        h,
                        f'{h:,.2f}',
                        ha='center',
                        va='bottom',
                        fontsize=8,
                    )

            elif tipo == 'Barras Horizontais':
                dados_h = dados_agrupados.sort_values(ascending=True)
                ax.barh(
                    dados_h.index.astype(str),
                    dados_h.values,
                    color='#10B981',
                )
                titulo = self.entry_titulo.get().strip()
                ax.set_title(titulo or f'{agregacao} de {col_val} por {col_cat}')
                ax.set_xlabel(col_val)
                ax.set_ylabel(col_cat)
                ax.grid(axis='x', linestyle='--', alpha=0.6)

            elif tipo == 'Pizza':
                ax.pie(
                    dados_agrupados.values,
                    labels=dados_agrupados.index.astype(str),
                    autopct='%1.1f%%',
                    startangle=140,
                    colors=plt.cm.Set3.colors,
                )
                titulo = self.entry_titulo.get().strip()
                ax.set_title(titulo or f'Distribuição de {col_val} por {col_cat}')
                ax.axis('equal')

            elif tipo == 'Linhas':
                dados_linha = dados_agrupados.sort_index()
                ax.plot(
                    dados_linha.index.astype(str),
                    dados_linha.values,
                    marker='o',
                    color='#8B5CF6',
                    linewidth=2,
                )
                titulo = self.entry_titulo.get().strip()
                ax.set_title(titulo or f'Evolução de {col_val} por {col_cat}')
                ax.set_xlabel(col_cat)
                ax.set_ylabel(col_val)
                ax.tick_params(axis='x', rotation=35)
                ax.grid(True, linestyle='--', alpha=0.6)

            elif tipo == 'Dispersão':
                ax.scatter(
                    df_temp[col_cat].astype(str),
                    df_temp[col_val],
                    color='#F59E0B',
                    alpha=0.75,
                )
                titulo = self.entry_titulo.get().strip()
                ax.set_title(titulo or f'{col_val} em relação a {col_cat}')
                ax.set_xlabel(col_cat)
                ax.set_ylabel(col_val)
                ax.tick_params(axis='x', rotation=35)
                ax.grid(True, linestyle='--', alpha=0.6)

            fig.tight_layout()
            self.figura_atual = fig
            self.mostrar_grafico()

            self.btn_abrir_grafico.config(state='normal')
            self.btn_salvar.config(state='normal')
            self.btn_exportar.config(state='normal')
            self.lbl_status.config(
                text=f'{len(dados_agrupados):,} categorias exibidas'
            )
            valores = df_temp[col_val]
            self.lbl_resumo.config(
                text=(
                    f'Linhas usadas: {len(df_temp):,} de {len(self.df):,}  |  '
                    f'Categorias: {df_temp[col_cat].nunique():,}  |  '
                    f'Soma: {valores.sum():,.2f}  |  '
                    f'Média: {valores.mean():,.2f}  |  '
                    f'Mediana: {valores.median():,.2f}  |  '
                    f'Mínimo: {valores.min():,.2f}  |  '
                    f'Máximo: {valores.max():,.2f}'
                )
            )

        except Exception as e:
            messagebox.showerror(
                'Erro ao gerar gráfico', f'Ocorreu um erro:\n{e}'
            )

    def salvar_imagem(self):
        if not self.figura_atual:
            return

        caminho_salvar = filedialog.asksaveasfilename(
            title='Salvar Gráfico',
            defaultextension='.png',
            filetypes=[
                ('Imagem PNG', '*.png'),
                ('Imagem JPEG', '*.jpg'),
                ('Documento PDF', '*.pdf'),
            ],
        )

        if caminho_salvar:
            try:
                self.figura_atual.savefig(
                    caminho_salvar, dpi=300, bbox_inches='tight'
                )
                messagebox.showinfo(
                    'Sucesso', 'Imagem do gráfico salva com sucesso!'
                )
            except Exception as e:
                messagebox.showerror(
                    'Erro ao salvar', f'Não foi possível salvar a imagem:\n{e}'
                )

    def exportar_dados(self):
        if self.dados_agrupados is None:
            return
        caminho = filedialog.asksaveasfilename(
            title='Exportar dados agregados',
            defaultextension='.csv',
            filetypes=[('CSV', '*.csv')],
        )
        if not caminho:
            return
        try:
            nome_categoria = self.combo_categoria.get()
            nome_valor = self.combo_valor.get()
            dados = self.dados_agrupados.rename_axis(nome_categoria).reset_index(
                name=nome_valor
            )
            dados.to_csv(caminho, index=False, encoding='utf-8-sig')
            messagebox.showinfo(
                'Sucesso', 'Dados agregados exportados com sucesso!'
            )
        except Exception as e:
            messagebox.showerror(
                'Erro ao exportar',
                f'Não foi possível exportar os dados:\n{e}',
            )


if __name__ == '__main__':
    print('Iniciando janela do Tkinter...')
    root = tk.Tk()
    app = GeradorGraficosApp(root)
    root.mainloop()
