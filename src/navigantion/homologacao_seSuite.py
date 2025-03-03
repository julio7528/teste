import re
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
from selenium.webdriver.common.action_chains import ActionChains


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
        print(f"Erro: {e}")


    

def criar_driver():
    """Configura e retorna uma instância do WebDriver para o Chrome."""
    print("Criando o WebDriver...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    print("WebDriver criado com sucesso.")
    return driver



def HomologacaoSeSuite(df_Homologacao):
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



        #=============================================Inicio Etapa Minhas Tarefas=========================================

        for index, row in df_Homologacao.iterrows():
            nomearquivo = row['nomearquivo']
            # Expressão regular para capturar "MAME.###" ou "EME.###"
            match = re.search(r'\b(EME|EMP|EPA|EPE|EPI|MAME|MAMP|MAPA|MAPE|MAPI)[.-](\d{3})\b(?:[^\w](.*?))?', nomearquivo)
            if match:
                resultado = match.group(0).replace('-','.').strip()  # Captura o tipo e o código completo (e.g., "EMP.079")
                print(f"Linha {index}: Resultado = {resultado}")
            else:
                print(f"Linha {index}: Nenhuma correspondência encontrada no nomearquivo '{nomearquivo}'")

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

            time.sleep(4)

            base_page.trocar_para_frame("isoright")

            result = base_page.find_and_click_row_homologacao(f"{resultado}")  # Substitua pelo nome do arquivo desejado


            # Validação com base no status
            if result["status"] == "success":   
                            
                base_page.click_checked_checkbox_column()


                base_page.voltar_para_conteudo_principal()
                base_page.trocar_para_frame('iframe')


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
                update_log_data(f"{nomearquivo}", statushomologado='OK')
                
                base_page.switch_to_window_by_title("Documento (DC010)") 

                print("Processo finalizado com sucesso!")

            elif result["status"] == "no_homologation":
                print("Linha encontrada, mas não possui title='Homologação'.")
                update_log_data(f"{nomearquivo}", statushomologado='ERRO')
            elif result["status"] == "not_found":
                print("Nenhuma linha correspondente encontrada.")
            else:
                print(f"Erro: {result['message']}")


    except Exception as e:
        print(f"Erro durante a execução: {e}")
        messagebox.showerror("Erro", f"Erro durante a execução: {e}")
    finally:
        print("Encerrando o WebDriver...")
        driver.quit()