import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

global ID
global nomes

Tk ().withdraw ()
arquivo = askopenfilename ()

ERROR = True
while ERROR:
    Intro = print("Seu arquivo apresenta as seguintes abas: ")
    sheets = pd.ExcelFile ( arquivo )
    sheets = sheets.sheet_names

    tamanho_sheets = len ( sheets )
    for i in range(tamanho_sheets):
        print(f'[{i}] ' + sheets[i] )
    sn = int ( input ( 'informe qual aba deseja coletar os dados: '))

    arquivo_pd = pd.read_excel ( arquivo, sheet_name= sn )
    arquivo_df = pd.DataFrame ( arquivo_pd )
    try:
        nomes = arquivo_df.columns.values
        print(nomes)
        ID = input ( 'Qual o nome do cabeçário que estão os nomes? ' ).strip ()
        nomes = arquivo_df[ID].values

    except:
        ERROR = False
        exit(f'O valor fornecido no cabeçário é inválido {34}')


    curso = str ( input ( 'informe o curso: ' ) ).strip ()
    promovido = str ( input ( "forneça o nome de quem esta promovendo o evento: " ) ).strip ()
    dia = str ( input ( 'Informe da data do curso (dd/mm/aaaa): ' ) )
    horas = str ( input ( 'Informe as horas do curso: ' ) ).strip ()

    try :
        nome_arquivo = input("informe o nome que deseja para o seu arquivo: ")
        with open ( f'{nome_arquivo}.doc', 'w' ) as aq :
            for nome in nomes :
                texto = f"Certificamos que {nome} participou do {curso}, promovido {promovido} da Universidade Federal do Rio Grande – FURG no dia {dia}, completando {horas} horas de atividades."
                aq.write ( texto )
                aq.write ( '\n' * 20 )
    except PermissionError :
        f"Por favor feche o arquivo antes de realizar a operação"
    finally:
        break
input("Aperte 'ENTER' para sair")