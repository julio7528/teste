import os
import psutil
import time

def fechar_processos_office_e_limpar_memoria(pastas):
    """
    Fecha processos Office (Word, Excel, PowerPoint) e tenta liberar memória
    associada a esses processos.  Verifica arquivos abertos nas pastas.

    Args:
        pastas: Lista de caminhos de pastas a serem verificadas.
    """

    programas_alvo = {
        'WINWORD.EXE': ['.doc', '.docx'],
        'EXCEL.EXE': ['.xls', '.xlsx', '.xlsm'],
        'POWERPNT.EXE': ['.ppt', '.pptx'],
    }

    processos_fechados = set()  # Conjunto para rastrear PIDs dos processos já fechados

    for pasta in pastas:
        try:
            for nome_arquivo in os.listdir(pasta):
                if not nome_arquivo.startswith('~$'):
                    caminho_completo = os.path.join(pasta, nome_arquivo)
                    _, extensao = os.path.splitext(nome_arquivo)
                    extensao = extensao.lower()

                    for programa, extensoes_validas in programas_alvo.items():
                        if extensao in extensoes_validas:
                            for proc in psutil.process_iter(['pid', 'name', 'open_files', 'memory_info']):
                                try:
                                    if proc.info['name'] == programa and proc.pid not in processos_fechados:
                                        arquivo_aberto = False

                                        #Verifica se tem arquivos abertos
                                        if proc.info['open_files']:
                                            for file in proc.info['open_files']:
                                                if file.path == caminho_completo:
                                                    arquivo_aberto = True
                                                    break  # Sai do loop interno se encontrar o arquivo

                                        # Fecha *mesmo que não tenha o arquivo aberto*, mas seja um processo Office
                                        print(f"Processo Office encontrado: {proc.info['name']} (PID: {proc.info['pid']})")
                                        proc.kill()
                                        processos_fechados.add(proc.pid) # Adiciona ao conjunto de processos fechados
                                        print(f"Processo {proc.info['name']} (PID: {proc.info['pid']}) terminado.")

                                        # Tentativa de liberar memória (após fechar o processo)
                                        try:
                                            # Aguarda um curto período para o sistema liberar recursos.
                                            time.sleep(0.5)
                                            #Tenta coletar informações, mesmo depois de fechar
                                            if proc.is_running():
                                                mem_info = proc.memory_info()
                                                print(f"  Memória (antes da liberação): RSS={mem_info.rss / (1024 * 1024):.2f} MB, VMS={mem_info.vms / (1024 * 1024):.2f} MB")
                                        except (psutil.NoSuchProcess, psutil.ZombieProcess):
                                            print("Processo já não existe ou é zumbi (memória provavelmente já liberada).")
                                        except Exception as mem_err:
                                            print(f"  Erro ao obter informações de memória: {mem_err}")


                                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                                    pass
                                except Exception as e:
                                    print(f"Erro inesperado ao verificar processos: {e}")

        except FileNotFoundError:
            print(f"Erro: A pasta {pasta} não foi encontrada.")
        except Exception as e:
            print(f"Erro inesperado ao listar arquivos na pasta {pasta}: {e}")

    if not processos_fechados:
        print("Nenhum processo Office relevante encontrado nas pastas.")



def listar_arquivos(pastas):
    """
    Lista arquivos após fechar processos Office e tentar liberar memória.
    """

    fechar_processos_office_e_limpar_memoria(pastas)

    arquivos = []
    for pasta in pastas:
        try:
            for nome_arquivo in os.listdir(pasta):
                if not nome_arquivo.startswith('~$'):
                    caminho_completo = os.path.join(pasta, nome_arquivo)
                    arquivos.append(caminho_completo)
        except FileNotFoundError:
            print(f"Erro: A pasta {pasta} não foi encontrada.")
        except Exception as e:
            print(f"Erro inesperado ao listar arquivos na pasta {pasta}: {e}")
    return arquivos



# --- Exemplo de Uso ---
pastas_teste = ['C:\\RPA\\RPA001_Garantia_De_Qualidade\\data\\METODOS-ANEXOS']  # Sua lista de pastas
arquivos_listados = listar_arquivos(pastas_teste)
print("\nArquivos listados:")
for arq in arquivos_listados:
    print(arq)