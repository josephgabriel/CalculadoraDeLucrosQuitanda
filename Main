from tkinter import *
from tkinter import ttk, messagebox

import numpy as np
import pandas as pd
import openpyxl as xl

class CalculadoraLucro():
    def __init__(self):
        self.root = Tk()
        self.root.title("Quitanda do José")
        self.root.geometry("1100x650")
        self.root.resizable(False, False)

        self.color_green = "#2ca063"
        self.color_yellow = "#e6a04b"
        self.color_white = "#f8f8f8"

        self.containers()
        self.itens_container01()
        self.itens_container02()
        self.itens_container03()
        self.root.mainloop()

    def containers(self):
        self.fr_container01 = Frame(
            self.root,
            width=1050,
            height=30,
            bg='white'
        )
        self.fr_container02 = Frame(
            self.root,
            width=1050,
            height=250,
            bg='white'
        )
        self.fr_container03 = Frame(
            self.root,
            width=1100,
            height=370,
            bg=self.color_green
        )

        self.fr_container01.propagate(0)
        self.fr_container02.propagate(0)
        self.fr_container03.propagate(0)
        self.fr_container01.pack()
        self.fr_container02.pack()
        self.fr_container03.pack()

    def itens_container01(self):
        self.lb_title = Label(
            self.fr_container01,
            text='Calculadora de Lucros',
            font='Verdana 20',
            bg=self.fr_container01.cget('bg')
        )
        self.lb_title.pack()

    def itens_container02(self):
        self.fr_subcontainer01 = Frame(
            self.fr_container02,
            bg='white',
            highlightthickness=1,
            highlightcolor='gray'
        )
        self.fr_subcontainer02 = Frame(
            self.fr_container02,
            bg='white',
            highlightthickness=1,
            highlightcolor='gray'
        )
        self.fr_subcontainer03 = Frame(
            self.fr_container02,
            bg='white',
            highlightthickness=1,
            highlightcolor='gray'
        )

        self.fr_container_btn = Frame(
            self.fr_subcontainer01,
            bg=self.fr_subcontainer01.cget('bg')
        )

        # secao 01
        self.lb_title_section01 = Label(
            self.fr_subcontainer01,
            text='Cadastro de produto',
            font='Verdana 15',
            bg=self.fr_subcontainer01.cget('bg')
        )

        # Nome do produto
        self.lb_nome_produto = Label(
            self.fr_subcontainer01,
            text='Nome do produto',
            font='Verdana',
            bg=self.fr_subcontainer01.cget('bg')
        )
        self.en_nome_produto = Entry(
            self.fr_subcontainer01,
            font='Verdana',
            highlightthickness=1,
            highlightbackground='gray',
            highlightcolor=self.color_green,
            bg='white'
        )

        # Quantidade do produto
        self.lb_qtd_produto = Label(
            self.fr_subcontainer01,
            text='Quantidade do produto',
            font='Verdana',
            bg=self.fr_subcontainer01.cget('bg')
        )
        self.en_qtd_produto = Entry(
            self.fr_subcontainer01,
            font='Verdana',
            highlightthickness=1,
            highlightbackground='gray',
            highlightcolor=self.color_green,
            bg='white'
        )

        # Preço de compra
        self.lb_preco_compra = Label(
            self.fr_subcontainer01,
            text='Preço de compra',
            font='Verdana',
            bg=self.fr_subcontainer01.cget('bg')
        )
        self.en_preco_compra = Entry(
            self.fr_subcontainer01,
            font='Verdana',
            highlightthickness=1,
            highlightbackground='gray',
            highlightcolor=self.color_green,
            bg='white'
        )

        # Preço de venda
        self.lb_preco_venda = Label(
            self.fr_subcontainer01,
            text='Preço de venda',
            font='Verdana',
            bg=self.fr_subcontainer01.cget('bg')
        )
        self.en_preco_venda = Entry(
            self.fr_subcontainer01,
            font='Verdana',
            highlightthickness=1,
            highlightbackground='gray',
            highlightcolor=self.color_green,
            bg='white'
        )

        # Custo frete
        self.lb_custo_frete = Label(
            self.fr_subcontainer01,
            text='Custo frete',
            font='Verdana',
            bg=self.fr_subcontainer01.cget('bg')
        )
        self.en_custo_frete = Entry(
            self.fr_subcontainer01,
            font='Verdana',
            highlightthickness=1,
            highlightbackground='gray',
            highlightcolor=self.color_green,
            bg='white'
        )

        # Custo adicional
        self.lb_custo_adicional = Label(
            self.fr_subcontainer01,
            text='Custo adicional',
            font='Verdana',
            bg=self.fr_subcontainer01.cget('bg')
        )
        self.en_custo_adicional = Entry(
            self.fr_subcontainer01,
            font='Verdana',
            highlightthickness=1,
            highlightbackground='gray',
            highlightcolor=self.color_green,
            bg='white'
        )

        # Botões
        self.btn_calcular = Button(
            self.fr_container_btn,
            text='Calcular',
            fg='white',
            bg='#0097b2',
            command=lambda: None
        )

        self.btn_salvar = Button(
            self.fr_container_btn,
            text='Salvar',
            fg='white',
            bg=self.color_green,
            font = 'Verdana',
            command=self.salvar_registro
        )

        self.btn_deletar = Button(
            self.fr_container_btn,
            text='Deletar',
            fg='white',
            bg='#ff3131',
            command=lambda: None
        )

        # secao 02
        self.lb_title_section02 = Label(
            self.fr_subcontainer02,
            text='Operações',
            font='Verdana 15',
            bg=self.fr_subcontainer02.cget('bg')
        )

        self.lb_texto_resultado_operacao = Label(
            self.fr_subcontainer02,
            text='O resultado aparecerá aqui assim que calcular um produto',
            font='Verdana',
            wraplength=250,
            justify=LEFT,
            bg=self.fr_subcontainer02.cget('bg')
        )

        self.lb_title_lucro_liquido = Label(
            self.fr_subcontainer02,
            text='Lucro liquido',
            font='Verdana',
            bg=self.fr_subcontainer02.cget('bg')
        )

        self.lb_resultado_lucro_liquido = Label(
            self.fr_subcontainer02,
            text='R$000,00',
            font='Verdana',
            width=20,
            padx=100,
            pady=5,
            highlightthickness=1,
            highlightbackground='gray',
            bg=self.fr_subcontainer02.cget('bg')
        )

        self.lb_title_margem_lucro = Label(
            self.fr_subcontainer02,
            text='Margem de Lucro',
            font='Verdana',
            bg=self.fr_subcontainer02.cget('bg')
        )

        self.lb_resultado_margem_lucro = Label(
            self.fr_subcontainer02,
            text='0,00%',
            font='Verdana',
            width=20,
            padx=100,
            pady=5,
            highlightthickness=1,
            highlightbackground='gray',
            bg=self.fr_subcontainer02.cget('bg')
        )

        # Alocando os containers
        self.fr_subcontainer01.grid(row=0, column=0, padx=10, pady=7, sticky=N)
        self.fr_subcontainer02.grid(row=0, column=1, padx=10, pady=7, sticky=N)
        self.fr_subcontainer03.grid(row=0, column=2, padx=10, pady=7, sticky=N)

        # Posicionando elementos subcontainer01
        self.lb_title_section01.grid(row=0, columnspan=2, pady=10)
        self.lb_nome_produto.grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.en_nome_produto.grid(row=1, column=1, sticky=W, padx=5)
        self.lb_qtd_produto.grid(row=2, column=0, sticky=W, padx=5, pady=5)
        self.en_qtd_produto.grid(row=2, column=1, sticky=W, padx=5)
        self.lb_preco_compra.grid(row=3, column=0, sticky=W, padx=5, pady=5)
        self.en_preco_compra.grid(row=3, column=1, sticky=W, padx=5)
        self.lb_preco_venda.grid(row=4, column=0, sticky=W, padx=5, pady=5)
        self.en_preco_venda.grid(row=4, column=1, sticky=W, padx=5)
        self.lb_custo_frete.grid(row=5, column=0, sticky=W, padx=5, pady=5)
        self.en_custo_frete.grid(row=5, column=1, sticky=W, padx=5)
        self.lb_custo_adicional.grid(row=6, column=0, sticky=W, padx=5, pady=5)
        self.en_custo_adicional.grid(row=6, column=1, sticky=W, padx=5)

        self.fr_container_btn.grid(row=7, columnspan=2, padx=10, sticky=W)
        self.btn_calcular.grid(row=0, column=0, sticky=W, padx=5)
        self.btn_salvar.grid(row=0, column=1, sticky=W, padx=5)
        self.btn_deletar.grid(row=0, column=2, sticky=W, padx=5)

        # Posicionando elementos do subcontainer02
        self.lb_title_section02.grid(row=0, column=0, padx=10)
        self.lb_texto_resultado_operacao.grid(row=1, column=0, padx=30)
        self.lb_title_lucro_liquido.grid(row=2, column=0)
        self.lb_resultado_lucro_liquido.grid(row=3, column=0, padx=5)
        self.lb_title_margem_lucro.grid(row=4, column=0)
        self.lb_resultado_margem_lucro.grid(row=5, column=0, padx=5, pady=(0, 60))

    def itens_container03(self):
        colunas_tabela = [
            'Nome do Produto', 'Preço de Compra(R$)', 'Preco de Venda(R$)', 'Qtd',
            'Custos adicionais(R$)', 'Custo médio do frete(R$)', 'Custo total(R$)',
            'Lucro Liquido(R$)', 'Margem de Lucro(%)'
        ]

        self.treeview = ttk.Treeview(
        self.fr_container03,
        columns=colunas_tabela,
        show='headings'
    )

        for col in colunas_tabela:
         self.treeview.heading(col, text=col)
         self.treeview.column(col, width=150, anchor=CENTER)

        # Scrollbar horizontal
        h_scroll = Scrollbar(self.fr_container03, orient=HORIZONTAL, command=self.treeview.xview)
        self.treeview.configure(xscrollcommand=h_scroll.set)
        
        # Scrollbar vertical
        v_scroll = Scrollbar(self.fr_container03, orient=VERTICAL, command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=v_scroll.set)

        # Posicionando a Treeview e as barras de rolagem
        self.treeview.grid(row=0, column=0, sticky='nsew')
        h_scroll.grid(row=1, column=0, sticky='ew')
        v_scroll.grid(row=0, column=1, sticky='ns')

        self.fr_container03.grid_rowconfigure(0, weight=1)
        self.fr_container03.grid_columnconfigure(0, weight=1)

        self.atualizar_tabela()

    # função para atualizar a tabela quando houver alterações
    def atualizar_tabela(self):
        df_list = self.obter_dados_tabela('planilha.xlsx')

        for i in self.treeview.get_children():
            self.treeview.delete(i)

        for item in df_list:
            self.treeview.insert('', 'end', values=item)

    def obter_dados_tabela(self, nome_planilha):
        wb = xl.load_workbook(nome_planilha)

        sheet = wb.active
        dados = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            dados.append(row)

        return dados
    #Função para salvar dados na planilha
    def salvar_registro(self):
        if self.valiar_entrys() == True:
            nome_produto = self.en_nome_produto.get()
            qtd_produto = float(self.en_qtd_produto.get())
            preco_venda = float(self.en_preco_venda.get())
            preco_compra = float(self.en_preco_compra.get())
            custo_adicional = float(self.en_custo_adicional.get())
            custo_frete = float(self.en_custo_frete.get())

            lucro_liquido = preco_venda - preco_compra - custo_adicional - custo_frete

            custo_geral = (preco_compra+ custo_adicional + custo_frete) * qtd_produto

            margem_lucro = (lucro_liquido/preco_venda) * 100
           #Carregar planilha existente ou criar uma nova
            try:
                wb = xl.load_workbook('planilha.xlsx')
                sheet = wb.active
            except:
                wb = xl.Workbook()
                sheet = wb.active
                sheet.title('Plan1')

            #Adicionando cabeçalhos
                sheet['A1'] = 'Nome do produto'
                sheet['B1'] = 'Preco da compra'
                sheet['C1'] = 'Preco da venda'
                sheet['D1'] = 'Quantidade'
                sheet['E1'] = 'Custos Adicionais'
                sheet['F1'] = 'Custo do Frete'
                sheet['G1'] = 'Custo Total'
                sheet['H1'] = 'Lucro Liquido'
                sheet['I1'] = 'Margem de Lucro (%)'

            #Adicionando valores na minha planilha
            row = sheet.max_row + 1
            sheet['A{}', format(row)] = nome_produto
            sheet['B{}', format(row)] = preco_compra
            sheet['C{}', format(row)] = preco_venda
            sheet['D{}', format(row)] = qtd_produto
            sheet['E{}', format(row)] = custo_adicional
            sheet['F{}', format(row)] = custo_frete
            sheet['G{}', format(row)] = custo_geral
            sheet['H{}', format(row)] = lucro_liquido
            sheet['I{}', format(row)] = margem_lucro

            wb.save('planilha.xlsx')
            self.atualizar_tabela()

        else:
            messagebox.showinfo('Campo vazio, existe algum campo obrigatorio vazio!')

    def valiar_entrys(self):
        campos = [
            self.en_nome_produto.get(),
            self.en_qtd_produto.get(),
            self.en_preco_compra.get(),
            self.en_nome_produto.get(),
            self.en_custo_adicional.get(),
            self.en_custo_frete.get()
        ]
        if all(campo.strip() for campo in campos):
            return True
        else: 
            return False
        #quero que todos os itens atribuidos dentro de campo tiver algum vazio me retorne false, caso ao contario retorne true.
CalculadoraLucro()
