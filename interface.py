import os
from tkinter import *
from tkinter import filedialog

from pdf2image import convert_from_path

class Conversor:
    opções = ['Converter pdf em jpg', 'Mesclar arquivos']

    def __init__(self, tela):
        self.frame_mestre = LabelFrame(tela, padx=0, pady=0)
        self.frame_mestre.pack(padx=1, pady=1)

        self.frame_de_botões = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.titulo = Label(
            self.frame_mestre, text='Conversor',
            pady=0, padx=200, bg='red', fg='white', bd=2, relief=SUNKEN, font=('Helvetica', 10, 'bold'))

        self.botão_de_seleção_de_aquivos = Button(
            self.frame_de_botões, text='Incluir arquivos na lista', command=self.seleciona_arquivos,
            padx=20, pady=0, bg='white', fg='red', font=('Helvetica', 9, 'bold'), bd=1)

        self.botão_excluir_arquivos_da_lista = Button(
            self.frame_de_botões, text='Retirar arquivos da lista', command=self.exclui_arquivos_da_lista,
            padx=20, pady=0, bg='white', fg='red', font=('Helvetica', 9, 'bold'), bd=1)

        self.listbox_de_arquivos = Listbox(self.frame_mestre, width=70, selectmode=MULTIPLE)

        self.variavel_de_opções = StringVar()
        self.variavel_de_opções.set("Selecione uma opção")
        self.validacao = OptionMenu(self.frame_mestre, self.variavel_de_opções, *Conversor.opções)

        self.roda_pe = Label(
            self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='red',
            fg='white', font=('Helvetica', 8, 'italic'), anchor=E
        )

        self.botão_converter = Button(
            self.frame_mestre, text='Apagar', command=self.converte_pdf_para_jpg,
            padx=0, pady=0, bg='red', fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo.grid(row=0, column=1, pady=10, padx=0, sticky=W+E)
        self.frame_de_botões.grid(row=1, column=1, sticky=W+E)
        self.botão_de_seleção_de_aquivos.grid(row=1, column=1, padx=30, pady=5)
        self.botão_excluir_arquivos_da_lista.grid(row=1, column=2, padx=10, pady=5)
        self.listbox_de_arquivos.grid(row=2, column=1, pady=10)
        self.validacao.grid(row=3, column=1, pady=10, ipadx=0, ipady=0)
        self.botão_converter.grid(row=4, column=1, pady=10)
        self.roda_pe.grid(row=5, column=1, pady=10, sticky=W+E)

        self.lista_de_arquivos = {}

    def atualiza_listbox(self):
        self.listbox_de_arquivos.delete(0, END)
        for item in self.lista_de_arquivos:
            self.listbox_de_arquivos.insert(END, f'{item}')

    def seleciona_arquivos(self):
        arquivos = filedialog.askopenfilenames(filetypes=(('Arquivos', '*.pdf'),("Todos os arquivos", '*.*')))
        for arquivo in arquivos:
            self.lista_de_arquivos[f'{os.path.basename(arquivo)}'] = arquivo
        self.atualiza_listbox()

    def exclui_arquivos_da_lista(self):
        for arquivo in self.listbox_de_arquivos.curselection():
            print(f'O arquivo {self.listbox_de_arquivos.get(arquivo)} foi excluído da lista para conversão')
            del self.lista_de_arquivos[self.listbox_de_arquivos.get(arquivo)]
        self.atualiza_listbox()

    def converte_pdf_para_jpg(self):
        if not self.lista_de_arquivos.keys():
            print('Nenhum arquivo selecionado.')
        else:
            caminho_para_salvar_arquivos = filedialog.askdirectory()
            print(caminho_para_salvar_arquivos)
            for arquivo in self.lista_de_arquivos:
                páginas = convert_from_path(self.lista_de_arquivos[arquivo], 300)
                print(f'o arquivo contém {len(páginas)} páginas')
                arquivo_sem_extensão = self.lista_de_arquivos[arquivo][:-4]
                contador = 0
                for página in páginas:
                    if contador < 9:
                        páginas[contador].save(f"{arquivo_sem_extensão} - {0}{contador + 1}.jpg", "JPEG")
                        contador += 1
                    else:
                        páginas[contador].save(f"{arquivo_sem_extensão} - {contador + 1}.jpg", "JPEG")
                        contador += 1
