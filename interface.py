import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

from pdf2image import convert_from_path
import PyPDF2

class Conversor:
    opções = ['Selecione uma opção' ,'Converter pdf em jpg', 'Mesclar arquivos']

    def __init__(self, tela):
        self.frame_mestre = LabelFrame(tela, padx=0, pady=0)
        self.frame_mestre.pack(padx=1, pady=1)

        self.frame_de_botões = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.titulo = Label(
            self.frame_mestre, text='Conversor',
            pady=0, padx=200, bg='red', fg='white', bd=2, relief=SUNKEN, font=('Helvetica', 12, 'bold'))

        self.botão_de_seleção_de_aquivos = Button(
            self.frame_de_botões, text='Incluir arquivos na lista', command=self.seleciona_arquivos,
            padx=20, pady=0, bg='white', fg='red', font=('Helvetica', 9, 'bold'), bd=1)

        self.botão_excluir_arquivos_da_lista = Button(
            self.frame_de_botões, text='Retirar arquivos da lista', command=self.exclui_arquivos_da_lista,
            padx=20, pady=0, bg='white', fg='red', font=('Helvetica', 9, 'bold'), bd=1)

        self.listbox_de_arquivos = Listbox(self.frame_mestre, width=70, selectmode=EXTENDED)

        self.variavel_de_opções = StringVar()
        self.variavel_de_opções.set("Selecione uma opção")
        self.validacao = OptionMenu(self.frame_mestre, self.variavel_de_opções, *Conversor.opções)

        self.roda_pe = Label(
            self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='red',
            fg='white', font=('Helvetica', 8, 'italic'), anchor=E
        )

        self.botão_converter = Button(
            self.frame_mestre, text='Executar', command=self.selecionador_de_opções,
            padx=0, pady=0, bg='red', fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo.grid(row=0, column=1, pady=10, padx=0, sticky=W+E)
        self.frame_de_botões.grid(row=1, column=1, sticky=W+E)
        self.botão_de_seleção_de_aquivos.grid(row=1, column=1, padx=35, pady=5)
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
        arquivos = filedialog.askopenfilenames(filetypes=(('Arquivos PDF', '*.pdf'),("Todos os arquivos", '*.*')))
        for arquivo in arquivos:
            self.lista_de_arquivos[f'{os.path.basename(arquivo)}'] = arquivo
        self.atualiza_listbox()

    def exclui_arquivos_da_lista(self):
        if not self.listbox_de_arquivos.curselection():
            messagebox.showerror('Não dá pra apagar o que não existe!', 'Não há arquivos listados e selecionados.')
        else:
            for arquivo in self.listbox_de_arquivos.curselection():
                messagebox.showinfo('Foi-se!', f'O arquivo ({self.listbox_de_arquivos.get(arquivo)}) foi excluído da lista para conversão')
                del self.lista_de_arquivos[self.listbox_de_arquivos.get(arquivo)]
            self.atualiza_listbox()

    def selecionador_de_opções(self):
        if self.variavel_de_opções.get() == 'Converter pdf em jpg':
            self.converte_pdf_para_jpg()
        elif self.variavel_de_opções.get() == 'Mesclar arquivos':
            self.mesclar_arquivos_pdf()
        else:
            messagebox.showwarning('Tem que escolher, fi!', 'Nenhuma opção selecionada!')

    def converte_pdf_para_jpg(self):
        if not self.lista_de_arquivos.keys():
            messagebox.showwarning('Tem nada aqui!', 'Nenhum arquivo incluído na lista de conversão.')
        else:
            caminho_para_salvar_arquivos = filedialog.askdirectory()
            número_de_páginas_dos_arquivos = {}
            for arquivo in self.lista_de_arquivos:
                páginas = convert_from_path(self.lista_de_arquivos[arquivo], 300)
                número_de_páginas_dos_arquivos[arquivo] = str(len(páginas))
                arquivo_sem_extensão = self.lista_de_arquivos[arquivo][:-4]
                arquivo_base = os.path.basename(arquivo_sem_extensão)
                contador = 0
                for página in páginas:
                    if contador < 9:
                        páginas[contador].save(f"{caminho_para_salvar_arquivos}/{arquivo_base} - {0}{contador + 1}.jpg", "JPEG")
                        contador += 1
                    else:
                        páginas[contador].save(f"{caminho_para_salvar_arquivos}/{arquivo_base} - {contador + 1}.jpg", "JPEG")
                        contador += 1
            páginas_por_arquivo = ''
            for arquivo_convertido in número_de_páginas_dos_arquivos:
                páginas_por_arquivo += f'{arquivo_convertido}; Número de imagens: {número_de_páginas_dos_arquivos[arquivo_convertido]}.\n'
            messagebox.showinfo('Rolou tranquilo!', f'Imagens criadas com sucesso!\n\n{páginas_por_arquivo}')



    def mesclar_arquivos_pdf(self):
        if not self.lista_de_arquivos.keys():
            messagebox.showwarning('Tem nada aqui!', 'Nenhum arquivo incluído na lista de conversão.')
        else:
            caminho_para_salvar_arquivos = filedialog.askdirectory()
            if os.path.exists(f'{caminho_para_salvar_arquivos}/Arquivos_mesclados.pdf'):
                print('Já existe um arquivo com o nome de "Arquivos_mesclados" na pasta que você está tentando criar.\n'
                      'Renomeie o arquivo existente antes de prosseguir com a ação.')
            else:
                with open(f'{caminho_para_salvar_arquivos}/Arquivos_mesclados.pdf', 'wb') as arquivo_final:
                    criador_de_pdf = PyPDF2.PdfFileWriter()
                    for arquivo in self.lista_de_arquivos:
                        with open(self.lista_de_arquivos[arquivo], 'rb') as arquivo_aberto:
                            arquivo_lido = PyPDF2.PdfFileReader(arquivo_aberto)
                            for página in range(arquivo_lido.numPages):
                                página_do_pdf = arquivo_lido.getPage(página)
                                criador_de_pdf.addPage(página_do_pdf)
                            criador_de_pdf.write(arquivo_final)