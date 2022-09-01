import time
from turtle import right
import pyautogui as pgui
import pyperclip as pclip
import pandas as pd

def abreSistema():
    #abrir o menu iniciar e localizar o navegador, no meu caso, uso o zorin os e por isso a função não é exibida usando a tecla win, se for outro s.o que abra o menu com o win, seria mais simples o processo 
    pgui.click(x=830, y=881)    
    pgui.write("chrome")
    pgui.press("enter")
    time.sleep(5)
    pgui.hotkey("ctrl","t")#rodou ok
    pclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?uschromep=sharing")
    time.sleep(2)
    pgui.hotkey("ctrl","v") 
    pgui.press("enter")
    time.sleep(5)
    pgui.click(x=1183, y=288, clicks=2, interval=0.2)#primeiro clique para expandir a pasta
    time.sleep(5)
    pgui.rightClick(x=1183, y=288) #abre o menu lateral
    pgui.click(x=1247, y=761) #clica na opção download do menu
    pgui.hotkey("ctrl","w")#fecha a guia aberta
    
    
def trazDados():        
    #faz a importação dos dados do sistema, analiza eles e atribui os calculos necessarios de acordo com o solicitado
    tabela = pd.read_excel("Vendas - Dez.xlsx") 
    quantidade = tabela["Quantidade"].sum()
    faturamento = tabela["Valor Final"].sum()
    gerarCorpoEmail(quantidade, faturamento) #chama a função de geração do email
  
def gerarCorpoEmail(quantidade, faturamento):
    #recebe o texto que será inserido no corpo do email encaminhado a diretoria
    texto = f"""
    Bom dia a todos!
    O faturamento do dia de ontem foi de R$ {faturamento:.2f} , tendo sido vendidos {quantidade:} itens do nosso estoque.
    
    Abraços e um bom dia de trabalho a todos!
    """
    enviarEmail(texto) #chama a função de envio dos emails

def enviarEmail(texto):
    #inicialmente abre o navegador e localiza o email utilizado
    pgui.click(x=830, y=881)
    pgui.write("chrome")
    pgui.press("enter")
    time.sleep(5)
    pgui.hotkey("ctrl","t")
    pgui.write("mail.google.com/mail/u/0/#inbox")
    pgui.press("enter")
    time.sleep(5)
    #caso precise digitar os dados de login e senha do email:
    #pgui.write("seuemail@email.com")
    #pgui.press("enter")
    #time.sleep(5)
    #pgui.write("senhaAqui")
    #pgui.press("enter")
    pgui.click(x=918, y=214) #abre um novo email
    time.sleep(2)
    pgui.write("seuemail@email.com")
    pgui.press("tab", presses=2, interval=0.2)
    time.sleep(2)
    pclip.copy("Relatório de Vendas")
    pgui.hotkey("Ctrl", "v")
    pgui.press("tab")
    time.sleep(2)
    pclip.copy(texto)
    pgui.hotkey("Ctrl", "v")
    pgui.click(x=1638, y=822) #finaliza o envio 
    pgui.hotkey("ctrl","w") #fecha a tela

abreSistema()
trazDados()
corpoGerado= trazDados()







