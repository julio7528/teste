import re, os, traceback, inspect
import time
import pyautogui
from pywinauto import Application
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from src.config.config import get_caminho_rede, get_url_hml, get_user_SeSuite, get_password_SeSuite, get_contra_senha
from src.navigantion.base_page import BasePage
from src.navigantion.login_page import LoginPage
from src.services.db_service import  update_log_data



def upload_file_with_pywinauto(file_path):
    """
    Insere o caminho do arquivo no diálogo do Explorador de Arquivos do Windows usando PyWinAuto.

    :param file_path: Caminho completo do arquivo no sistema.
    """
    time.sleep(2)  # Aguarda o diálogo abrir

    try:
        # Conecta ao diálogo "Abrir"
        app = Application(backend="win32").connect(title_re="Abrir|Open")
        dialog = app.top_window()

        # Localiza o campo de entrada do ComboBox e insere o texto diretamente
        edit_field = dialog.child_window(class_name="Edit")
        edit_field.set_text(file_path)  # Define o caminho do arquivo diretamente

        # Localiza o botão "Abrir" e clica
        confirmar_button = dialog.child_window(title_re="&Abrir|Open", class_name="Button")
        confirmar_button.click()

    except Exception as e:
        print(f"Erro: {os.path.basename(__file__)} - Message: {e} - Line: {traceback.extract_tb(e.__traceback__)[0][1]}")
   

def criar_driver():
    """Configura e retorna uma instância do WebDriver para o Chrome."""
    print("Criando o WebDriver...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    print("WebDriver criado com sucesso.")
    return driver


def NavegationSeSuite(df_SeSuite):
    print("Iniciando navegação no SeSuite...")
    driver = criar_driver()

    try:
        # URL do sistema
        url = get_url_hml()
        user = get_user_SeSuite()
        password = get_password_SeSuite()
        print(f"Acessando URL: {url}")

        # Variáveis de Tempo
        Tempo = 0
        TempoCurto = 10
        TempoMedio = 30
        TempoLongo = 60
        TempoMuitoLogon = 120

        # Inicialização de páginas
        login_page = LoginPage(driver)
        base_page = BasePage(driver)

        # Acessar o sistema
        login_page.open_url(url)
        print("Sistema acessado. Realizando login...")

        time.sleep(2)

        element_xpath = "//*[@id='container']/div/div[1]/div/div/button"
        base_page.click_element(By.XPATH, element_xpath)
        print("Botão inicial clicado.")

        login_page.login(user, password)
        print("Login realizado com sucesso.")

        # Validação e clique no alerta
        print("Validando e clicando no alerta de confirmação...")
        timeout = TempoMedio
        element_xpath_clicar = "//*[@id='alertConfirm']"
        element_xapth_validar = "//*[@id='components']/a"
        base_page.validar_e_clicar(By.XPATH, timeout, element_xpath_clicar, element_xapth_validar)
        print("Alerta confirmado ou validado com sucesso.")



        #INICIO DO LOOP PARA NAVEGAÇÃO DOS ARQUIVOS

        for index, row in df_SeSuite.iterrows():
            nomearquivo = row['nomearquivo']
            # Expressão regular para capturar "MAME.###" ou "EME.###"
            # match = re.search(r'\b(EME|EMP|EPA|EPE|EPI|MAME|MAMP|MAPA|MAPE|MAPI)[.-](\d{3})\b(?:[^\w](.*?))?', nomearquivo)
            regex = r'(EME|EMP|EPA|EPE|EPI|MAME|MAMP|MAPA|MAPE|MAPI)[.-](\d{3})\b(?:[^\w](.*?))?' # Removed the first \b
            match = re.search(regex, nomearquivo)
            
            if match:
                resultado = match.group(0).replace('-','.').strip()  # Captura o tipo e o código completo (e.g., "EMP.079")
                print(f"Linha {index}: Resultado = {resultado}")
            else:
                print(f"Linha {index}: Nenhuma correspondência encontrada no nomearquivo '{nomearquivo}'")
        
            
            # Clicando no ícone
            element_xpath = "//*[@id='components']/a"
            base_page.click_element(By.XPATH, element_xpath)
            print("Ícone clicado.")

        
            # Clicando em documentos
            print("Clicando em documentos")
            timeout = TempoMedio
            element_xpath_clicar = "//*[contains(text(),'Documento')]"
            element_xapth_validar =  "//*[contains(text(),'Documento (DC010)')]"
            base_page.validar_e_clicar(By.XPATH, timeout, element_xpath_clicar, element_xapth_validar)
            print("Seção 'Documento' acessada.")
            
            # Na aba geral, clicar em documento (DC021)
            element_xpath = "//*[contains(text(),'Documento (DC010)')]"
            base_page.click_element(By.XPATH, element_xpath)
            print("Aba 'Documento (DC021)' clicada.")


            # Selecionando o tipo
            print("Click em categorias")
            timeout = TempoMedio
            element_xpath_clicar = "/html/body/div[2]/div/div/div/div[4]/div/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]"
            element_xapth_validar = "/html/body/div[2]/div/div/div/div[4]/div/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]"
            base_page.validar_e_clicar(By.XPATH, timeout, element_xpath_clicar, element_xapth_validar)
            
            element_xpath = "/html/body/div[5]/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div[1]/div/div/div/input"
            base_page.enter_text(By.XPATH, element_xpath, "CQ - Controle de Qualidade")
            

            # Clicar em CQ-Controle de Qualidade
            element_xpath = "//*[contains(text(), 'CQ - Controle de Qualidade')]"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicado em 'CQ-Controle de Qualidade'.")

            # Clicar em Aplicar
            element_xpath = "//*[contains(text(), 'Aplicar')]"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicado em 'Aplicar'.")

            # Clicando em filtro
            element_xpath = "/html/body/div[2]/div/div/div/div[4]/div/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/div/div/div[3]/button"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicando em filtro")

            # Clicando em Identificador
            element_xpath = "//*[contains(text(), 'Identificador')]"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicando em Identificador")


            element_xpath = "/html/body/div[5]/div[2]/div/div/div/div/div[2]/div/div/div[2]/div/div[2]/div/div/input"
            base_page.enter_text(By.XPATH, element_xpath, f"{resultado}")
            print("Pesquisando o Identificador")

            # Clicar em Aplicar
            element_xpath = "//*[contains(text(), 'Aplicar')]"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicado em 'Aplicar'.")


            # Clicar em Pesquisar
            element_xpath = "//*[contains(text(), 'Pesquisar')]"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicado em 'Pesquisar'.")


            # Clicando no primeiro item da tabela
            element_xpath = "/html/body/div[2]/div/div/div/div[4]/div/div/div[1]/div[1]/div/div[2]/div[3]/div/div[1]/div/div[1]/div[1]/table/tbody/tr/td[6]"
            #base_page.click_element(By.XPATH, element_xpath)

            time.sleep(3)
            print(f"Clicando no arquivo correspondente: {nomearquivo}")
            
            base_page.find_and_click_row(nomearquivo)

            time.sleep(1)

            #Clicando em revisão
            element_xpath = "/html/body/div[2]/div/div/div/div[4]/div/div/div[2]/div/div[1]/div[4]/i"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicando em revisão")

            time.sleep(1)
            #Clicando em Criar
            element_xpath = "/html/body/div[2]/div/div/div/div[4]/div/div/div[2]/div/div[5]/div[2]/div/div[2]/div/div/div/div[1]/button"
            base_page.click_element(By.XPATH, element_xpath)
            print("Clicando em Criar")

            timeout = TempoMedio
            element_xpath_clicar = "/html/body/div[5]/div/div/div[2]/div[2]/div[2]/button/span/div/span"
            element_xapth_validar = "/html/body/div[5]/div/div/div[2]/div[2]/div[2]/button/span/div/span"
            base_page.validar_e_clicar(By.XPATH, timeout, element_xpath_clicar, element_xapth_validar)

            time.sleep(3)

            base_page.switch_to_window_by_title("Revisão da estrutura")

            base_page.trocar_para_frame("iframeComposed")
            

            element_xpath = "/html/body/table/tbody/tr/td[2]/button[1]/img"
            base_page.click_element(By.XPATH, element_xpath)
        
            base_page.wait_and_accept_alert()

            base_page.close_current_window_and_switch()
            
            
            base_page.switch_to_window_by_title("Dados do documento")    

            #Clicando em Revisão
            element_xpath = "/html/body/span/div/div[1]/div[1]/div[2]/div/div[1]/div/div/div[3]/div[2]/table/tbody/tr/td[2]/a/span/span/span[2]"
            base_page.click_element(By.XPATH, element_xpath)

            base_page.trocar_para_frame("ribbonFrame")
            base_page.trocar_para_frame("iframeRevision")

            time.sleep(2)

            #Digita justificativa
            element_xpath = "/html/body/form/div[1]/div[2]/div/div/div/div[2]/div/div[2]/div[1]/div/div/div/div/div[4]/div/table/tbody/tr/td/table/tbody/tr/td/div/div/textarea"
            base_page.enter_text(By.XPATH, element_xpath, "Revisão dos documentos mediante ao CM-TBS-00728")

            #Clicar em Participante
            element_xpath = "/html/body/form/div[1]/div[2]/div/div/div/div[1]/div/div[2]/ul/li[2]/a"
            base_page.click_element(By.XPATH, element_xpath)

            base_page.trocar_para_frame("stepFrame")
            #Clicar em Importar Roteiro
            element_xpath = "/html/body/table/tbody/tr/td[2]/button[6]/img"
            base_page.click_element(By.XPATH, element_xpath)


            base_page.switch_to_window_by_title("Seleção de informação")

            #Digitando o identificador 
            element_xpath = "/html/body/table/tbody/tr[2]/td/div/div/form/div/table/tbody/tr/td/fieldset/table/tbody/tr/td[1]/input"
            base_page.enter_text(By.XPATH, element_xpath, "san")


            #Clicar em Pesquisar
            element_xpath = "/html/body/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/button/img"
            base_page.click_element(By.XPATH, element_xpath)

            time.sleep(2)
            #Clica em Salvar e Sair
            element_xpath = "/html/body/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr/td[4]/button/img"
            base_page.click_element(By.XPATH, element_xpath)

            time.sleep(2)
            base_page.switch_to_window_by_title("Dados do documento")

            base_page.voltar_para_conteudo_principal()

            #Clicar em Arquivo eletronico
            
            element_xpath = "//*[@id='btnEletronicfile-btnInnerEl']"
            base_page.execute_js_with_xpath("arguments[0].click();", element_xpath)

            time.sleep(3)


            base_page.trocar_para_frame("ribbonFrame")
            base_page.trocar_para_frame("iframeEletricFile")
            

            #Clicar em Substituir o arquivo
            element_xpath = "/html/body/div[4]/table/tbody/tr/td[2]/button[4]/img"
            base_page.click_element(By.XPATH, element_xpath)


            base_page.voltar_para_conteudo_principal()

            time.sleep(3)


            #Clica em Selecionar arquivo
            element_xpath = "//input[contains(@title, 'Selecionar arquivo')]"
            base_page.execute_js_with_xpath("arguments[0].click();", element_xpath)

            time.sleep(3)

            #Digitando o caminho do arquivo
            upload_file_with_pywinauto(rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS\{nomearquivo}")


            element_xpath = "#FileItem0-innerCt"
            validate_element = False
            while validate_element == False:
                validate_element = base_page.is_file_uploaded(element_xpath)


            time.sleep(3)
            

            #Clicando em Finalizar
            element_xpath = "//*[@id='dragbtnfinalise-btnIconEl']"
            base_page.execute_js_with_xpath("arguments[0].click();", element_xpath)

            time.sleep(3)


            #Clicando em Salvar e Sair
            element_xpath = "//*[@id='btnSaveExit-btnInnerEl']"
            base_page.execute_js_with_xpath("arguments[0].click();", element_xpath)

            time.sleep(3)

            base_page.switch_to_window_by_title("Documento (DC010)") 


            #=============================================Inicio Etapa Minhas Tarefas=========================================

            #Clicar em Minhas tarefas
            element_xpath = "/html/body/div[2]/div/div/div/div[1]/ul[1]/li[5]/a/div/b"
            base_page.click_element(By.XPATH, element_xpath)

            time.sleep(1)
            element_xpath = "/html/body/div[2]/div/div/div/div[1]/ul[1]/li[5]/div[1]/div[1]/div/div[1]/ul[3]/li[2]/a"
            base_page.click_element(By.XPATH, element_xpath)
            
            #Clicar em Revisão de documento
            element_xpath = "//*[contains(text(),'Revisão de documento')]"
            base_page.click_element(By.XPATH, element_xpath)

            base_page.trocar_para_frame("iframe")

            #Clico o filtro
            element_xpath = "/html/body/div[1]/form/div[2]/div/div[2]/div[2]/div[1]/div/div/div/div[1]/div/div/a[2]/span/span/span[1]"
            base_page.execute_js_with_xpath("arguments[0].click();", element_xpath)

            #Clico em documento
            element_xpath = "/html/body/div[5]/div/div[2]/div[2]/div[2]/a/span"
            base_page.execute_js_with_xpath("arguments[0].click();", element_xpath)


            #Digitando o identificador 
            element_xpath = "//*[@id='iddocument']"
            base_page.enter_text(By.XPATH, element_xpath, f"{resultado}")


            #Clicando em pesquisar
            element_xpath = "/html/body/div[1]/form/div[2]/div/div[2]/div[2]/div[1]/div/div/div/div[2]/div/div/a/span/span/span[1]"
            base_page.click_element(By.XPATH, element_xpath)

            time.sleep(2)

            #Clica em confirmar
            element_xpath = "/html/body/div[1]/form/div[2]/div/div[1]/div/span/div/div/div/table/tbody/tr/td[2]/table/tbody/tr[1]/td/div/div/div/div/div/div/a[1]/span/span/span[2]"
            base_page.click_element(By.XPATH, element_xpath)


            base_page.voltar_para_conteudo_principal()

            #Ditando contra senha
            element_xpath = "/html/body/div[5]/div/div/div[1]/div/form/div/input"
            base_page.enter_text(By.XPATH, element_xpath, f"{get_contra_senha()}")


            #Clicando em confirmar
            element_xpath = "//*[@id='confirm']"
            base_page.click_element(By.XPATH, element_xpath)

            time.sleep(3)

            base_page.switch_to_window_by_title("Dados do documento")

            #Aceitar revisão
            element_xpath = "//*[@id='btnAcceptRevision-btnIconEl']"
            base_page.click_element(By.XPATH, element_xpath)


            #confirmar a revisão
            element_xpath = "/html/body/div[6]/div/div[2]/button[2]"
            base_page.click_element(By.XPATH, element_xpath)

            print("Atualizando informação no banco")
            update_log_data(nomearquivo, statusenviadosesuite='OK')
            

            base_page.switch_to_window_by_title("Documento (DC010)") 

            print("Processo finalizado com sucesso!")

    except Exception as e:
        print(f"Erro: {os.path.basename(__file__)} - Message: {e} - Line: {traceback.extract_tb(e.__traceback__)[0][1]}")
        messagebox.showerror("Erro", f"Erro durante a execução: {e}")
    finally:
        print("Encerrando o WebDriver...")
        driver.quit()