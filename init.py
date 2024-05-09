import time
import gc
import os
import pyautogui
import sys

from threading import Thread
from functions.linx import open, auth, findStock, close
from functions.logger import logger
from functions.spreadsheet import selectProducts, checkPurchase, getHistory, getSeller, alterLastContact, verifyLastContact
from functions.whatsapp import noahSendMsgFile, noahSendMsgText
from functions.config2rfp import initSplash

from config import LINX_PATH, PRODUCTS_IMG_PATH, MESSAGEM

##--- INIT PROCESS
if __name__ == "__main__":
    app, message_window = initSplash()

    def run_main_program():
        gc.collect()

        logger(f"PROCESSO INICIADO", message_window)  

        ##--- SELECIONA NOVOS PRODUTOS (DENTRO DO PRAZO ESTIPULADO)
        newProducts = selectProducts()
        if not newProducts:
            logger(f"Nenhum produto encontrado. O processo foi finalizado.")
            return

        print("NOVOS PRODUTOS: " + str(newProducts) +"\n")

        ##--- VERIFICA QUAIS OS CLIENTES QUE COMPRARAM O PRODUTO REFERENCIA DENTRO DO PRAZO
        clients = checkPurchase(newProducts)
        if not clients:
            logger(f"Nenhum cliente encontrado. O processo foi finalizado.")
            print("Nenhum cliente encontrado. O processo foi finalizado.")
            return

        print("CLIENTES ENCONTRADOS NO FILTRO: " + str(clients) +"\n")


        ##--- VERIFICA QUAL O ULTIMO CONTATO COM O CLIENTE
        clients = verifyLastContact(clients)
        print("CLIENTES APOS VERIFICAÇÃO DO ULTIMO CONTATO: " + str(clients) +"\n")
        if not clients:
            logger(f"Nenhum cliente encontrado dentro do prazo estabelecido de contato minimo. O processo foi finalizado.")
            return


        ##--- VERIFICAR ESTOQUE NO LINX
        process = open(LINX_PATH)
        if not process:
            logger(f"ERROR: NÃO FOI POSSIVEL ABRIR O LINX")
            return

        #-- LOGIN
        coordinates = auth()
        if auth == False:
            logger(f"ERROR: NÃO FOI POSSIVEL FAZER LOGIN")
            return
        print("AQUI > "+str(coordinates))
        #-- NAVEGA ATÉ A TELA DE ESTOQUE
        x, y = coordinates[0], coordinates[1]
        pyautogui.click(x, y)
            
        for i in range(3):
            pyautogui.press("down")

        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(1)

        ##--- PESQUISA ESTOQUE DOS PRODUTOS SELECIONADOS
        estoque = findStock(clients)
        if not estoque:
            logger(f"*REGRA RETIRADA* NENHUM PRODUTO COM ESTOQUE FOI ENCONTRADO: O PROCESSO FOI FINALIZADO")
            #-- RETIRANDO A VERIFICAÇÃO DE ESTOQUE
            # close()
            # return

        print("PRODUTOS EM ESTOQUE: " + str(estoque) +"\n")

        close()

        ##--- PESQUISA HISTORICO DE COMPRA DOS CLIENTES SELECIONADOS
        finalResult = getHistory(estoque)
        if not finalResult:
            logger(f"ERROR: NÃO FOI POSSIVEL ACESSAR O HISTORICO DE CLIENTES")
            return

        print("HISTORICO DE CLIENTES RESGATADO : " + str(finalResult) +"\n")


        ##--- DISPARA MENSAGEM PARA VENDEDOR
        try:
            for item in finalResult:
                #-- PESQUISA VENDEDOR
                number = getSeller(item['seller'])

                if not number:
                    logger(f"VENDEDOR NAO ENCONTRADO NA PLANILHA: "+ str(item['seller']))
                else:
                    #-- FORMATA MENSAGEM
                    message = MESSAGEM.replace("#NAMECLIENTE#", str(item['nameClient']))
                    message = message.replace("#produto#", str(item['newProduct']))

                    #-- FORMATA HISTORICO
                    if item['history']:
                        msgHistory = "HISTORICO DE COMPRAS DO CLIENTE: " + str(item['nameClient']) + ".\n"
                        for values in item['history']:
                            msgHistory += "PRODUTO: " + str(values['cod']) +"\n"
                            msgHistory += "NUMERO: " + str(values['grid']) +"\n"
                            msgHistory += "VENDEDOR: " + str(values['seller']) +"\n"
                            msgHistory += "DATA: " + str(values['data']) +"\n"
                            msgHistory += "------------------------------ \n\n"

                    #-- BUSCAR IMAGEM DO PRODUTO
                    image = os.path.join(PRODUCTS_IMG_PATH, str(item['newProduct']+".jpg"))
                    print(image)

                    if not os.path.exists(image):
                        #-- DISPARO SEM IMAGEM
                        response = noahSendMsgText(number, message)
                    else:
                        #-- DISPARO COM IMAGEM
                        response = noahSendMsgFile(number, message, image, item['newProduct'])
                    time.sleep(5)
                    if item['history']:
                        response = noahSendMsgText(number, msgHistory)
                        
                    print("DISPARADO: " + str(number) +" - "+ str(response))
                    
                    alterLastContact(item)



        except Exception as e:
                logger(f"!! O script encontrou um erro no disparo de mensagens: {e}")

        print("PROCESSO FINALIZADO COM SUCESSO -----------")

main_thread = Thread(target=run_main_program)
main_thread.start()

sys.exit(app.exec_())        