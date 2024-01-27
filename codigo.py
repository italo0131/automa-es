import os
import pandas as pd 
import win32com.client as win32
import datetime
caminho = "automaçao_planilha/bases"
arquivos = os.listdir(caminho)

lista_tabelas = []


for nome_arquivo in arquivos:
    try:
        tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    
        if "Data de Venda" in tabela_vendas.columns:
            tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"], unit="d")
            lista_tabelas.append(tabela_vendas)
    
        else:
            print(f"A coluna 'Data de Venda' não está presente no arquivo {nome_arquivo}")

    except Exception as e :
        print(f"erro ao processar o arquivo {nome_arquivo}: {e}")


if lista_tabelas:
    tabela_consolidada = pd.concat(lista_tabelas)
    tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
    tabela_consolidada = tabela_consolidada.sort_index(drop=True)
    tabela_consolidada.to_excel("Vendas.xlsx", index=False)
    print(tabela_consolidada)


else:
    
    print("Nenhum DataFrame foi encontrado para processar." )

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateIntem(0)
email.to ="email aqi"
data_hoje = datetime.datetime.today().strftime("%d/%m/%Y")
email.Subject = f"relatorio de vendas {data_hoje}"
email.Body ="""
Prezados,


segue em anexo o Relatorio

"""

caminho = os.getcwd()

anexo = os.path.join(caminho, "#anexo")
email.Attachments.Add(anexo)

email.send()
