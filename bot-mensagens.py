import openpyxl #importando a biblioteca para ler a planilha
from urllib.parse import quote #faz com que a mensagem seja codificada de forma que o programa consiga enviá-la
import webbrowser #para conseguir abrir o navegador
from time import sleep #para dar o tempo de abertura da página do whatsapp
import pyautogui #para enviar a mensagem (clicar na seta de enviar no whatsapp)

webbrowser.open("https://web.whatsapp.com/")
sleep(30) #para o cliente fazer o login no whatsapp web

workbook = openpyxl.load_workbook("Pasta1.xlsx") #carregando a planilha (workbook -> variavel)
pagina_mensagens = workbook["Planilha1"] #chamar função

for linha in pagina_mensagens.iter_rows(min_row=2): #faz o programa ler a partir da 2 linha da planilha (onde está nome/telefone)
    
    nome = linha[0].value
    telefone = linha[1].value
    print(nome)
    print(telefone)
    #até aqui -> planilhas automatizadas   
    
    #mensagem formatada
    mensagem = f"Olá {nome}! Tenho um segredo para te contar... clique no link https://www.youtube.com/watch?v=dQw4w9WgXcQ"
    
    link_mensagem = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
    webbrowser.open(link_mensagem)
    sleep(10)
    try:
        sleep(5)  # Espera 5 segundos antes de pressionar Enter
        pyautogui.press('enter')  # Aciona a tecla Enter
        sleep(5)
        pyautogui.hotkey("ctrl", "w")  # Fecha a aba atual
        sleep(5)
    except Exception as e:
        print(f"Não foi possível enviar mensagem para {nome}: {e}")
    # Para registrar erros no arquivo csv
    with open("erros.csv", "a", newline="", encoding="utf-8") as arquivo:
        arquivo.write(f"{nome},{telefone}\n")

print("Processo concluído.")