# navigation/base_page.py

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchFrameException, WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
import time


class BasePage:

    def __init__(self, driver):
        self.driver = driver

    def open_url(self, url):
        """Abre uma URL no navegador."""
        self.driver.get(url)

    def find_element(self, by, value, timeout=10):
        """Encontra um elemento com espera explícita."""
        return WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    
    def close_current_window_and_switch(self):
        """
        Fecha a janela atual e alterna para a próxima janela disponível.
        """
        try:
            current_window = self.driver.current_window_handle  # Janela atual
            all_windows = self.driver.window_handles  # Todas as janelas abertas

            # Fecha a janela atual
            self.driver.close()

            # Alterna para outra janela (se existir)
            for window in all_windows:
                if window != current_window:
                    self.driver.switch_to.window(window)
                    return  # Sai da função após alternar

            print("Nenhuma outra janela disponível para alternar.")
        except Exception as e:
            print(f"Erro ao fechar e alternar a janela: {e}")


    def switch_to_window_by_title(self, title, timeout=10):
        """
        Muda para a janela com o título especificado.
        
        :param title: Título da janela desejada.
        :param timeout: Tempo máximo para tentar encontrar a janela, em segundos.
        """
        try:
            # Obtém os identificadores de todas as janelas abertas
            all_windows = self.driver.window_handles

            # Aguarda até que a janela com o título desejado seja encontrada
            end_time = time.time() + timeout
            while time.time() < end_time:
                for window_handle in all_windows:
                    try:
                        # Tenta alternar para a janela
                        self.driver.switch_to.window(window_handle)

                        # Verifica se o título da janela corresponde ao esperado
                        if title in self.driver.title:
                            print(f"Mudou para a janela com o título: {title}")
                            return  # Encontrou a janela e saiu da função
                    except Exception as e:
                        # Ignora janelas que não estão mais disponíveis
                        print(f"Erro ao tentar acessar a janela {window_handle}: {e}")
                time.sleep(0.5)  # Espera um pouco antes de tentar novamente

            # Se o título não for encontrado dentro do tempo
            print(f"Não foi possível encontrar a janela com o título '{title}' dentro do tempo.")
        except Exception as e:
            print(f"Erro ao tentar mudar para a janela com o título '{title}': {e}")
    
    def is_file_uploaded(self, elemento,timeout=10):
        """Valida se o arquivo foi carregado, verificando se o texto contém '100%'."""
        try:
            # Espera até que o elemento com o id 'FileItem0-innerCt' esteja presente
            file_item = self.find_element(By.CSS_SELECTOR, elemento, timeout)
            
            # Obtém o texto do elemento
            file_text = file_item.text.strip()

            # Verifica se o texto contém '100%' indicando que o upload foi concluído
            if '100%' in file_text:
                return True
            else:
                return False

        except Exception as e:
            print(f"Erro ao verificar upload do arquivo: {e}")
            return False

    def click_element(self, by, value, timeout=10):
        """Clica em um elemento após localizá-lo."""
        element = self.find_element(by, value, timeout)
        element.click()

    def enter_text(self, by, value, text, timeout=10):
        """
        Envia texto para um campo após garantir que ele esteja carregado e interativo.

        :param by: Método de localização do elemento (By.ID, By.XPATH, etc.).
        :param value: Valor usado para localizar o elemento.
        :param text: Texto a ser enviado ao campo.
        :param timeout: Tempo máximo para esperar pelo carregamento do elemento.
        """
        try:
            # Aguarda até que o elemento esteja clicável
            element = WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            element.clear()  # Limpa o campo antes de enviar texto
            element.send_keys(text)  # Envia o texto
            print(f"Texto '{text}' enviado para o elemento localizado por {by} = '{value}'")
        except TimeoutException:
            print(f"Elemento localizado por {by} = '{value}' não ficou clicável dentro do tempo limite de {timeout} segundos.")
        except Exception as e:
            print(f"Erro ao tentar enviar texto para o elemento: {e}")

    def wait_and_accept_alert(self, timeout=10):
        """Espera e aceita o alerta."""
        try:
            # Espera até que o alerta seja visível
            WebDriverWait(self.driver, timeout).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()  # Aceita o alerta
        except Exception as e:
            print(f"Erro ao aceitar o alerta: {e}")

    def switch_to_new_window(self):
        """Alterna para a nova janela ou aba."""
        # Pega o identificador da janela atual
        current_window = self.driver.current_window_handle

        # Pega o identificador de todas as janelas abertas
        all_windows = self.driver.window_handles

        # Alterna para a nova janela (que não é a janela atual)
        for window in all_windows:
            if window != current_window:
                self.driver.switch_to.window(window)
                break

    def switch_to_window_by_index(self, index):
            """Alterna para a janela pelo número (índice)."""
            all_windows = self.driver.window_handles

            # Verifica se o índice está dentro do intervalo válido
            if 0 <= index < len(all_windows):
                self.driver.switch_to.window(all_windows[index])
            else:
                raise IndexError("Índice fora do intervalo. Não existe uma janela com esse número.")


    def list_windows_and_urls(self):
        """Lista todas as janelas abertas e suas URLs."""
        windows = self.driver.window_handles  # Obtém todos os identificadores de janelas
        window_info = []  # Lista para armazenar informações das janelas

        for window in windows:
            self.driver.switch_to.window(window)  # Alterna para a janela
            url = self.driver.current_url  # Captura a URL da janela
            window_info.append({"window_handle": window, "url": url})
        
        return window_info
    
    def click_checked_checkbox_column(self):
            """
            Percorre as linhas da tabela, encontra o checkbox marcado e clica na sexta coluna da mesma linha.
            Se a ação abrir uma nova aba, a função a fecha e retorna para a aba principal.
            """
            try:
                # Verifica se o driver ainda está ativo
                if self.driver is None:
                    print("Erro: O driver do Selenium não está inicializado.")
                    return

                # Espera carregar a tabela
                self.driver.implicitly_wait(3)

                # Salva a referência da aba principal
                aba_principal = self.driver.current_window_handle

                # Encontra todas as linhas da tabela
                linhas = self.driver.find_elements(By.XPATH, "//table[@id='t_content_gridframe']/tbody/tr")

                for linha in linhas:
                    try:
                        # Encontra o checkbox da linha
                        checkbox = linha.find_element(By.XPATH, ".//td[@class='gridSelectTh']/input[@type='checkbox']")
                        
                        # Verifica se está marcado
                        if checkbox.is_selected():
                            # Localiza a sexta coluna (td[6])
                            sexta_coluna = linha.find_element(By.XPATH, ".//td[6]")

                            # Clica na sexta coluna
                            ActionChains(self.driver).move_to_element(sexta_coluna).click().perform()
                            print("Cliquei na sexta coluna da linha com checkbox marcado!")

                            # Aguarda a nova aba abrir
                            time.sleep(5)
                            
                            # Obtém todas as abas abertas
                            janelas = self.driver.window_handles

                            # Se houver uma nova aba, muda para ela
                            if len(janelas) > 1:
                                nova_aba = janelas[-1]  # Última aba aberta
                                self.driver.switch_to.window(nova_aba)
                                print("Alternado para a nova aba.")

                                # Realiza a interação necessária na nova tela
                                time.sleep(3)  # Ajuste se necessário
                                print("Interagindo com a nova tela...")

                                # Depois de interagir, fecha e retorna para a aba principal
                                self.close_current_window_and_switch()
                          

                    except NoSuchElementException:
                        print("Elemento não encontrado em uma das linhas, pulando...")
                    except WebDriverException as e:
                        print(f"Erro ao processar a linha: {e}")

            except WebDriverException as e:
                print(f"Erro crítico: {e}")
            
    def find_and_click_row(self, file_name):
        """
        Localiza a linha mais próxima com base no identificador do arquivo e clica no elemento correspondente.
        Se houver apenas uma linha na tabela, nenhuma ação é realizada.

        :param file_name: Nome do arquivo para buscar o identificador.
        """
        script = f"""
        // Função para encontrar o identificador mais próximo e relevante
        function findClosestIdentifier(fileName) {{
            const rows = document.querySelectorAll("tbody > tr");
            if (rows.length <= 1) {{
                // Retorna mensagem indicando que a tabela tem apenas uma linha
                return "A tabela contém apenas uma linha. Nenhuma ação será realizada.";
            }}

            let closestMatch = null;
            let highestScore = 0;

            rows.forEach(row => {{
                const identifierCell = row.querySelector('td:nth-child(6)');
                const identifier = identifierCell?.textContent.trim() || "";

                if (identifier && fileName.includes(identifier)) {{
                    const score = identifier.length;
                    if (score > highestScore) {{
                        highestScore = score;
                        closestMatch = {{ identifier, row }};
                    }}
                }}
            }});

            return closestMatch;
        }}

        // Obter o identificador mais próximo
        const closest = findClosestIdentifier("{file_name}");
        if (typeof closest === "string") {{
            // Retorna a mensagem se houver apenas uma linha
            return closest;
        }}

        console.log("Identificador encontrado:", closest.identifier);
        console.log("Linha correspondente:", closest.row);

        // Obter o número da linha
        const rowNumber = Array.from(document.querySelectorAll("tbody > tr")).indexOf(closest.row) + 1;

        // Gerar seletor dinâmico
        const rowSelector = `#st-container > div > div > div > div.center-container > div > div > div:nth-child(1) > div:nth-child(1) > div > div:nth-child(2) > div:nth-child(3) > div > div.rctContextMenuArea > div > div:nth-child(1) > div:nth-child(1) > table > tbody > tr:nth-child(${{rowNumber}}) > td:nth-child(4)`;
        const element = document.querySelector(rowSelector);
        if (element) {{
            element.click();
            return rowSelector; // Retorna o seletor clicado para debug
        }} else {{
            return "Elemento não encontrado.";
        }}
        """

        # Executa o JavaScript no contexto da página
        result = self.driver.execute_script(script)
        print("Resultado do script:", result)

    def find_and_click_row_homologacao(self, file_name):
        """
        Localiza a linha mais próxima com base no identificador do arquivo, marca o novo checkbox
        e desmarca qualquer checkbox previamente marcado.

        :param file_name: Nome do arquivo a ser buscado na tabela.
        :return: Dicionário com informações sobre o resultado da busca.
        """
        script = f"""
        function findClosestIdentifier(fileName) {{
            console.log("Iniciando busca pelo identificador...");

            const rows = document.querySelectorAll("#t_content_gridframe tbody > tr");
            if (rows.length === 0) {{
                console.log("Tabela não contém linhas.");
                return {{ status: "error", message: "Tabela não contém linhas." }};
            }}

            let closestMatch = null;
            let highestScore = 0;

            const fileNameWords = fileName.trim().toLowerCase().split(/\s+/);

            rows.forEach((row, index) => {{
                const identifierCell = row.querySelector("td:nth-child(11)"); 
                let identifier = identifierCell?.textContent.trim() || "";

                console.log(`Linha ${{index}}: Identificador encontrado -> '${{identifier}}'`);

                const identifierWords = identifier.toLowerCase().split(/\s+/);
                const allWordsMatch = fileNameWords.every(word => identifierWords.includes(word));

                if (allWordsMatch) {{
                    const score = identifierWords.length;
                    if (score > highestScore) {{
                        highestScore = score;
                        closestMatch = {{ identifier, row, index }};
                    }}
                }}
            }});

            if (!closestMatch) {{
                console.log("Nenhuma linha correspondente encontrada.");
                return {{ status: "not_found", message: "Nenhuma linha correspondente encontrada." }};
            }}

            console.log("Linha correspondente encontrada:", closestMatch);

            // Obtém o checkbox da linha encontrada
            const newCheckbox = closestMatch.row.querySelector("td.gridSelectTh input[type=checkbox]");

            if (!newCheckbox) {{
                console.log("Checkbox não encontrado na linha correspondente.");
                return {{ status: "error", message: "Checkbox não encontrado." }};
            }}

            // **Passo 1: Clicar no novo checkbox primeiro**
            if (!newCheckbox.checked) {{
                newCheckbox.click();
                console.log("Novo checkbox marcado.");
            }} else {{
                console.log("Checkbox já estava marcado.");
            }}

            // **Passo 2: Agora desmarca qualquer outro checkbox que estava marcado antes**
            rows.forEach((row) => {{
                const checkbox = row.querySelector("td.gridSelectTh input[type=checkbox]");
                if (checkbox && checkbox !== newCheckbox && checkbox.checked) {{
                    checkbox.click();
                    console.log("Checkbox anterior desmarcado.");
                }}
            }});

            return {{ 
                status: "success", 
                message: "Novo checkbox marcado e antigo desmarcado.",
                identifier: closestMatch.identifier
            }};
        }}

        // Executa a função e retorna o resultado
        return findClosestIdentifier("{file_name}");
        """

        # Executa o script no contexto da página com Selenium
        result = self.driver.execute_script(script)
        return result



    def execute_js_with_xpath(self, script, xpath, timeout=10):
        """
        Executa um script JavaScript em um elemento localizado pelo XPath, 
        aguardando o elemento estar presente no DOM.

        :param script: Código JavaScript que será executado.
        :param xpath: XPath do elemento no qual o script será executado.
        :param timeout: Tempo máximo (em segundos) para esperar o elemento.
        :return: Resultado da execução do script (se houver).
        """
        # Aguarda até que o elemento esteja presente e visível no DOM
        element = WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )

        # Executa o script no elemento localizado
        return self.driver.execute_script(script, element)


    def validar_e_clicar(self, by, timeout, xpath_alvo, xpath_alternativo):
        """
        Valida se um elemento aparece e clica nele. Se não aparecer, tenta o elemento alternativo.

        Args:
            by: Método de localização (e.g., By.XPATH).
            timeout: Tempo máximo de espera em segundos.
            xpath_alvo: XPath do elemento principal.
            xpath_alternativo: XPath do elemento alternativo.
        """
        poll_interval = 1  # Intervalo entre verificações, em segundos
        start_time = time.time()  # Marca o início do loop

        while True:
            # Verifica se o tempo limite foi atingido
            if time.time() - start_time > timeout:
                print("Tempo de espera excedido. Nenhum elemento encontrado.")
                break

            # Tenta encontrar e clicar no elemento principal
            try:
                element = self.find_element(by, xpath_alvo)
                if element:
                    element.click()
                    print("Elemento principal encontrado e clicado!")
            except Exception as e:
                print(f"Erro ao procurar o elemento principal: {e}")

            # Tenta encontrar e clicar no elemento alternativo
            try:
                element_alt = self.find_element(by, xpath_alternativo)

                if element_alt:
                    break
            except Exception as e:
                print(f"Erro ao procurar o elemento alternativo: {e}")

            # Aguarda antes de tentar novamente
            time.sleep(poll_interval)


    def trocar_para_frame(self, frame_identificador, timeout=10):
        """
        Troca para um iframe especificado por ID, nome ou índice.

        Args:
            frame_identificador: Identificador do frame (ID, nome ou índice).
            timeout: Tempo máximo de espera pelo frame (padrão: 10 segundos).

        Returns:
            True se a troca for bem-sucedida, False caso contrário.
        """
        try:
            if isinstance(frame_identificador, int):
                # Troca usando índice
                WebDriverWait(self.driver, timeout).until(
                    lambda d: len(d.find_elements(By.TAG_NAME, "iframe")) > frame_identificador
                )
                self.driver.switch_to.frame(frame_identificador)
            else:
                # Troca usando ID ou nome
                WebDriverWait(self.driver, timeout).until(
                    EC.frame_to_be_available_and_switch_to_it(frame_identificador)
                )
            print(f"Troca para o frame '{frame_identificador}' bem-sucedida.")
            return True
        except TimeoutException:
            print(f"Timeout: Frame '{frame_identificador}' não disponível.")
        except NoSuchFrameException:
            print(f"Erro: Frame '{frame_identificador}' não encontrado.")
        except Exception as e:
            print(f"Erro desconhecido ao trocar para o frame '{frame_identificador}': {e}")
        return False

    def voltar_para_conteudo_principal(self):
        """
        Retorna ao conteúdo principal fora de qualquer frame.
        """
        self.driver.switch_to.default_content()
        print("Voltou ao conteúdo principal.")
