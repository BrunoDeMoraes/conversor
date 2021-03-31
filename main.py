from tkinter import *
from interface import Conversor

tela = Tk()

conversor = Conversor(tela)
tela.resizable(False, False)
tela.title('GEOF - Conversor de arquivos')

tela.mainloop()
