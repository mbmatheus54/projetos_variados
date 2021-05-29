"""
Desenvolvido por Matheus Monteiro
"""

import pandas
from tkinter import *
from tkinter.filedialog import askopenfilename

if __name__ == '__main__' :
    # Abertura de uma janela de seleção de arquivo
    arquivo_ = Tk ().withdraw ()
    arquivo = askopenfilename ()

    # Interpretando os dados do .csv por meio da biblioteca ``pandas``
    dados_csv = pandas.read_csv ( arquivo, sep=',', encoding='utf-8' )
    dataframe = pandas.DataFrame ( dados_csv )

    # Restirando o limite de linhas e colunas que são mostradas no display
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    print ( dataframe )

    mainloop()
