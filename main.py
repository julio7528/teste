# Load config first to initialize logger
from src.config.config import load_config, get_database_url, get_caminho_rede, get_caminho_de_para, get_url_hml, get_user_SeSuite, get_password_SeSuite, generate_default_foldes, get_contra_senha
load_config()

# Now import other dependencies
import os, datetime
import re
import shutil
import psycopg2
from psycopg2 import sql
from tkinter import messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from src.utils.read_files_utils import capture_type_from_headers, capture_code_from_docx, capture_code_from_headers, convert_doc_to_docx, convert_docx_to_doc, read_excel_file, captura_tabela_embalagem, captura_codigo_nucleo, captura_tabela_fabricacao
from src.utils.edit_files_utiles import edit_file_ficha, edit_file_eme, edit_file_mame, edit_file_embalagem_fabricacao, adicionar_nova_linha_com_codigo_embalagem, substituir_codigo_nucleo,adicionar_nova_linha_fabricacao, editar_celula_codigo_embalagem
from src.utils.edit_files_utils_excel import edit_excel_codigo, adicionar_revisao, update_excel_footer
from src.utils.verify_excel_type import verify_excel_type
from src.navigantion.base_page import BasePage
from src.navigantion.login_page import LoginPage
from src.navigantion.upload_seSuite import NavegationSeSuite
from src.navigantion.homologacao_seSuite import HomologacaoSeSuite
from src.services.db_service import insert_log_data, query_to_dataframe, update_log_data
from src.utils.logger import ProcessType, LogStatus, get_logger
from sqlalchemy import create_engine
from sqlalchemy.sql import text

# Get logger instance after config is loaded
logger = get_logger()

os.system('cls' if os.name == 'nt' else 'clear')

# Função para carregar configurações
def carregar_configuracoes():
    try:
        logger.info("Carregando configurações do sistema", ProcessType.SYSTEM)
        print("\n" + "-"*50)
        print("CARREGANDO CONFIGURAÇÕES DO SISTEMA")
        print("-"*50)
        load_config()  # Assuming load_config correctly loads the configurations
        generate_default_foldes()
        logger.success("Configurações carregadas com sucesso", ProcessType.SYSTEM)
        print("✓ Configurações carregadas com sucesso.")
    except Exception as e:
        logger.critical(f"Erro ao carregar configurações: {e}", ProcessType.SYSTEM)
        print(f"❌ ERRO ao carregar configurações: {e}")

# Função para listar os arquivos nas pastas
def listar_arquivos(pastas):
    logger.info(f"Iniciando busca de arquivos em {len(pastas)} pastas", ProcessType.FILE)
    print("\n" + "-"*50)
    print("PROCURANDO ARQUIVOS PARA PROCESSAMENTO")
    print("-"*50)
    arquivos = []
    total_arquivos = 0
    
    for pasta in pastas:
        arquivos_pasta = 0
        logger.info(f"Verificando pasta: '{pasta}'", ProcessType.FILE)
        try:
            if os.path.exists(pasta) and os.path.isdir(pasta):
                for item in os.listdir(pasta):
                    caminho_completo = os.path.join(pasta, item)
                    if (os.path.isfile(caminho_completo) and
                        not item.startswith('.') and
                        not item.startswith('~$') and  # Adicionado o filtro para '~$'.
                        os.path.getsize(caminho_completo) > 0):
                        arquivos.append(caminho_completo)
                        arquivos_pasta += 1
                        total_arquivos += 1
                        logger.debug(f"Arquivo encontrado: {item}", ProcessType.FILE)
                logger.info(f"{arquivos_pasta} arquivo(s) encontrado(s) em '{pasta}'", ProcessType.FILE)
            else:
                logger.warning(f"'{pasta}' não é uma pasta válida ou não existe", ProcessType.FILE)

        except FileNotFoundError:
            logger.error(f"A pasta '{pasta}' não foi encontrada", ProcessType.FILE)
        except OSError as e:
            logger.error(f"Erro ao acessar a pasta '{pasta}': {e}", ProcessType.FILE)
        except Exception as e:
            logger.error(f"Erro inesperado na pasta '{pasta}': {e}", ProcessType.FILE)

    logger.info(f"Total de arquivos encontrados: {total_arquivos}", ProcessType.FILE)
    return arquivos

def gerar_relatorio_e_enviar_email(df):
    logger.info("Iniciando geração de relatório de execução", ProcessType.BUSINESS)
    print("\n" + "-"*50)
    print("GERANDO RELATÓRIO DE EXECUÇÃO")
    print("-"*50)
    print(f"Gerando relatório para {len(df)} registros...")

    # Salvar os dados no Excel
    relatorio_excel = "relatorio_rpa001.xlsx"
    df.to_excel(relatorio_excel, index=False, engine="openpyxl")
    logger.success(f"Relatório salvo em: {relatorio_excel}", ProcessType.FILE)
    print(f"✓ Relatório salvo em: {relatorio_excel}")

    # Configurações do e-mail
    remetente = "seu_email@example.com"
    senha = "sua_senha"
    destinatario = "destinatario@example.com"
    assunto = "Relatório RPA001 Controle Execução"

    logger.info("Preparando e-mail com relatório anexo", ProcessType.NETWORK)
    print("Preparando e-mail com relatório anexo...")
    # Criar o e-mail com anexo
    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = assunto

    # Anexar o arquivo Excel
    with open(relatorio_excel, "rb") as arquivo:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(arquivo.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(relatorio_excel)}")
        msg.attach(part)

    # Enviar o e-mail
    try:
        logger.info("Enviando e-mail com relatório", ProcessType.NETWORK)
        print("Enviando e-mail...")
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(remetente, senha)
            server.sendmail(remetente, destinatario, msg.as_string())
        logger.success("E-mail enviado com sucesso", ProcessType.NETWORK)
        print("✓ E-mail enviado com sucesso!")
        print(f"  - De: {remetente}")
        print(f"  - Para: {destinatario}")
        print(f"  - Assunto: {assunto}")
        print(f"  - Anexo: {relatorio_excel}")
    except Exception as e:
        logger.error(f"Erro ao enviar e-mail: {e}", ProcessType.NETWORK)
        print(f"❌ ERRO ao enviar e-mail: {e}")

# Função para editar o arquivo conforme o tipo
def editar_arquivo(arquivo, codigo_word, novo_codigo, tipoDocumento, df_para):
    try:
        logger.info(f"Identificando tipo do arquivo: {os.path.basename(arquivo)}", ProcessType.FILE)
        type_file = capture_type_from_headers(arquivo)
        logger.info(f"Tipo identificado: {type_file}", ProcessType.FILE)
        logger.info(f"Substituindo código: {codigo_word} → {novo_codigo}", ProcessType.BUSINESS)
        
        if type_file == "FICHA DE ANÁLISE":
            logger.info("Processando FICHA DE ANÁLISE...", ProcessType.BUSINESS)
            edit_file_ficha(arquivo, codigo_word, f"{codigo_word} ({novo_codigo})", arquivo)
            logger.success("FICHA DE ANÁLISE processada com sucesso", ProcessType.BUSINESS)
            return "Processado"
        elif type_file == "EME" or type_file == "EMP" or type_file == "EPA" or type_file == "EPE" or type_file == "EPI":
            logger.info(f"Processando {type_file}...", ProcessType.BUSINESS)
            edit_file_eme(arquivo, codigo_word, f"{codigo_word} ({novo_codigo})", arquivo)
            logger.success(f"{type_file} processado com sucesso", ProcessType.BUSINESS)
            return "Processado"
        elif type_file == "MAME" or type_file == "MAMP" or type_file == "MAPA" or type_file == "MAPE" or type_file == "MAPI":
            logger.info(f"Processando {type_file}...", ProcessType.BUSINESS)
            edit_file_mame(arquivo, codigo_word, f"{codigo_word} ({novo_codigo})", arquivo)
            logger.success(f"{type_file} processado com sucesso", ProcessType.BUSINESS)
            return "Processado"
        elif tipoDocumento=="INSTRUÇÃO DE EMBALAGEM":
            logger.info("Processando INSTRUÇÃO DE EMBALAGEM...", ProcessType.BUSINESS)
            df_embalagem = captura_tabela_embalagem(arquivo, "Componentes – Material de Embalagem")
            logger.info(f"Tabela de embalagem capturada: {len(df_embalagem)} linhas", ProcessType.BUSINESS)

            novas_linhas = []
            for item in df_embalagem:
                codigo_item = item['Código']
                logger.info(f"Processando código de embalagem: {codigo_item}", ProcessType.BUSINESS)
                codigo_para, descricao_item = buscar_codigo_para(codigo_item, df_para)

                if codigo_para != "Código não encontrado":
                    logger.info(f"Editando célula: {codigo_item} → {codigo_para}", ProcessType.BUSINESS)
                    editar_celula_codigo_embalagem(arquivo, str(codigo_item), str(codigo_para), arquivo)
            
            if codigo_word.startswith(('3', '5', '9')):
                logger.info(f"Editando cabeçalho com substituição direta: {codigo_word} → {novo_codigo}", ProcessType.BUSINESS)
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), str(novo_codigo), arquivo)
            else:
                logger.info(f"Editando cabeçalho com código duplo: {codigo_word} → {codigo_word} ({novo_codigo})", ProcessType.BUSINESS)
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), f"{str(codigo_word)} ({str(novo_codigo)})", arquivo)
            
            logger.success("INSTRUÇÃO DE EMBALAGEM processada com sucesso", ProcessType.BUSINESS)
            return "Processado"

        elif tipoDocumento=="INSTRUÇÃO DE FABRICAÇÃO":
            logger.info("Processando INSTRUÇÃO DE FABRICAÇÃO...", ProcessType.BUSINESS)
            if codigo_word.startswith(('3', '5', '9')):
                logger.info(f"Editando cabeçalho com substituição direta: {codigo_word} → {novo_codigo}", ProcessType.BUSINESS)
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), str(novo_codigo), arquivo)
            else:
                logger.info(f"Editando cabeçalho com código duplo: {codigo_word} → {codigo_word} ({novo_codigo})", ProcessType.BUSINESS)
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), f"{str(codigo_word)} ({str(novo_codigo)})", arquivo)

            codigos = captura_codigo_nucleo(arquivo, "Componentes – Núcleo")
            logger.info(f"Códigos núcleo capturados: {len(codigos)} códigos", ProcessType.BUSINESS)

            for cod in codigos:
                logger.info(f"Processando código núcleo: {cod}", ProcessType.BUSINESS)
                codigo_para, descricao_item = buscar_codigo_para(cod, df_para)

                if codigo_para != "Código não encontrado":
                    if codigo_word.startswith(('3')):
                        logger.info(f"Substituindo código núcleo diretamente: {cod} → {codigo_para}", ProcessType.BUSINESS)
                    substituir_codigo_nucleo(arquivo, str(cod), str(codigo_para), arquivo)
                else:
                    logger.info(f"Substituindo código núcleo com código duplo: {cod} → {cod} ({codigo_para})", ProcessType.BUSINESS)
                    substituir_codigo_nucleo(arquivo, cod, f"{str(cod)} ({str(codigo_para)})", arquivo)

            df_fabricacao = captura_tabela_fabricacao(arquivo, "Componentes – Núcleo")
            logger.info(f"Tabela de fabricação capturada: {len(df_fabricacao)} linhas", ProcessType.BUSINESS)
            
            novas_linhas = []
            for item in df_fabricacao:
                codigo_item = item['Código']
                logger.info(f"Processando código de fabricação: {codigo_item}", ProcessType.BUSINESS)
                codigo_para, descricao_item = buscar_codigo_para(codigo_item, df_para)

                if codigo_para != "Código não encontrado":
                    logger.info(f"Editando célula: {codigo_item} → {codigo_para}", ProcessType.BUSINESS)
                    editar_celula_codigo_embalagem(arquivo, str(codigo_item), str(codigo_para), arquivo)

                    logger.success("INSTRUÇÃO DE FABRICAÇÃO processada com sucesso", ProcessType.BUSINESS)
                    return "Processado"
                else:
                    logger.warning(f"Tipo de arquivo desconhecido: {arquivo}", ProcessType.FILE)
                    return "Tipo de arquivo desconhecido"
    except Exception as e:
        logger.exception(f"Erro ao editar o arquivo {os.path.basename(arquivo)}", ProcessType.FILE)
        return "Erro ao Processar"

# Função para buscar código na tabela
def buscar_codigo_para(codigo_de, df):
    logger.info(f"Buscando correspondência para código: {codigo_de}", ProcessType.DATABASE)
    try:
        # Filtrar o DataFrame com base no 'Codigo - DE'
        resultado = df.loc[df['Codigo - DE'].astype(str) == str(codigo_de)]
        
        # Verificar se a busca retornou resultados
        if not resultado.empty:
            codigo_para = resultado['Codigo -  PARA'].iloc[0]
            descricao_item = resultado.iloc[0, -1]  # Acessa a última coluna da linha
            logger.success(f"Correspondência encontrada: {codigo_de} → {codigo_para}", ProcessType.DATABASE)
            return codigo_para, descricao_item
        else:
            logger.warning(f"Código {codigo_de} não encontrado na tabela DE-PARA", ProcessType.DATABASE)
            return 'Código não encontrado', None
    except Exception as e:
        logger.error(f"Erro ao buscar código {codigo_de}: {e}", ProcessType.DATABASE)
        return f'Erro na busca de código {e}', None

# Função para processar arquivos .doc ou .docx
def processar_arquivo(arquivo, caminho_input, df):
    logger.info(f"Iniciando processamento do arquivo: {os.path.basename(arquivo)}", ProcessType.BUSINESS)
    print("\n" + "-"*50)
    print(f"PROCESSANDO: {os.path.basename(arquivo)}")
    print(f"Caminho: {caminho_input}")
    print(f"Tipo: {os.path.splitext(arquivo)[1]}")
    print("-"*50)
    
    try:
        extensao = os.path.splitext(arquivo)[1]
        arquivo_novo = os.path.join(caminho_input, arquivo.replace(".doc", ".docx") if extensao == ".doc" else arquivo)
        resultado = None
    except Exception as e:
        logger.error(f"Erro ao preparar o arquivo: {e}", ProcessType.FILE)
        print(f"❌ ERRO ao preparar o arquivo: {e}")
        resultado = "Erro ao Processar"
    
    match extensao:
        case ".doc" | ".docx":
            backup_path = rf"{get_caminho_rede()}\ARQUIVOS_BACKUP"
            logger.info(f"Criando backup em: {backup_path}", ProcessType.FILE)
            shutil.copy(os.path.join(caminho_input, arquivo), backup_path)
            converted = False
            if extensao == ".doc":
                logger.info(f"Convertendo .doc para .docx: {arquivo}", ProcessType.FILE)
                arquivo_doc = os.path.join(caminho_input, arquivo)
                convert_doc_to_docx(arquivo_doc, arquivo_novo)
                converted = True
                logger.success("Conversão concluída", ProcessType.FILE)
            
            logger.info("Capturando código do documento...", ProcessType.BUSINESS)
            codigo_word = capture_code_from_docx(arquivo_novo)
            tipoDocumento = ""
            
            if codigo_word == "Nenhum código encontrado.":
                logger.info("Código não encontrado no conteúdo, tentando capturar dos cabeçalhos...", ProcessType.BUSINESS)
                codigo_word, tipoDocumento = capture_code_from_headers(arquivo_novo)
                logger.info(f"Tipo de documento identificado: {tipoDocumento}", ProcessType.BUSINESS)

                if codigo_word == "Nenhum código encontrado.":
                    logger.error(f"Código não encontrado no arquivo: {arquivo_novo}", ProcessType.BUSINESS)
                    logger.info(f"Movendo arquivo para pasta ERRO...", ProcessType.FILE)
                    shutil.move(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\ERRO")
                    insert_log_data(arquivo, "ERRO")
                    logger.success("Arquivo movido para pasta ERRO e log atualizado", ProcessType.FILE)
                    return
                
            logger.info(f"Código encontrado: {codigo_word}", ProcessType.BUSINESS)
            ultimo_codigo = re.split(r'[\/,]', codigo_word)[-1].strip()
            logger.info(f"Último código a ser processado: {ultimo_codigo}", ProcessType.BUSINESS)
            
            logger.info("Buscando correspondência na tabela DE-PARA...", ProcessType.DATABASE)
            novo_codigo, descricao_item = buscar_codigo_para(ultimo_codigo, df)

            if novo_codigo == "Código não encontrado":
                logger.error(f"Correspondência não encontrada para: {ultimo_codigo}", ProcessType.DATABASE)
                logger.info(f"Movendo arquivo para pasta ERRO...", ProcessType.FILE)
                dir_destino = rf"{get_caminho_rede()}\ERRO"
                arq_destino = os.path.join(dir_destino, arquivo) # Caminho de destino inicial com nome original

                try: # Bloco try-except simplificado
                    if os.path.exists(arq_destino):
                        logger.warning(f"Arquivo '{arquivo}' já existe em '{dir_destino}'. Renomeando...", ProcessType.FILE)
                        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                        arquivo_novo = f"{timestamp}_{arquivo}"
                        arq_destino = os.path.join(dir_destino, arquivo_novo) # **Atualiza arq_destino com o novo nome**
                        logger.info(f"Novo nome para o arquivo será: '{arquivo_novo}'", ProcessType.FILE)

                    logger.info(f"Movendo arquivo '{arquivo}' para '{arq_destino}'...", ProcessType.FILE)
                    shutil.move(os.path.join(backup_path, arquivo), arq_destino) # **Usa arq_destino como destino**
                    logger.success(f"Arquivo '{arquivo}' movido com sucesso para '{arq_destino}'.", ProcessType.FILE)
                    insert_log_data(arquivo, "ERRO") # Log ainda usa o 'arquivo' original, se for desejado
                    logger.success("Log atualizado", ProcessType.DATABASE)

                except Exception as e: # Captura qualquer erro durante a movimentação
                    logger.error(f"ERRO ao mover arquivo '{arquivo}' para '{arq_destino}'. Erro: {e}", ProcessType.FILE)
                    insert_log_data(arquivo, f"ERRO - Movimentação: {e}") # Log de erro genérico
                    logger.error(f"Log de erro registrado: {e}", ProcessType.DATABASE)

                return

            logger.info(f"Iniciando edição do arquivo: {arquivo_novo}", ProcessType.BUSINESS)
            resultado = editar_arquivo(arquivo_novo, codigo_word, novo_codigo, tipoDocumento, df)
            logger.info(f"Resultado da edição: {resultado}", ProcessType.BUSINESS)
            
            if resultado == "Processado":
                logger.info("Movendo arquivos para as pastas correspondentes...", ProcessType.FILE)
                pasta_destino = rf"{get_caminho_rede()}\PROCESSADOS"
                logger.info(f"Movendo arquivo para: {pasta_destino}", ProcessType.FILE)
                
                if os.path.exists(os.path.join(pasta_destino, os.path.basename(arquivo))):
                    logger.warning(f"Arquivo já existe no destino, removendo arquivo existente...", ProcessType.FILE)
                    os.remove(os.path.join(pasta_destino, os.path.basename(arquivo)))
                shutil.move(os.path.join(backup_path, arquivo), pasta_destino)
                
                if converted:
                    logger.info("Convertendo .docx de volta para .doc...", ProcessType.FILE)
                    convert_docx_to_doc(arquivo_novo, os.path.join(caminho_input, arquivo))
                    os.remove(arquivo_novo)
                    logger.success("Conversão reversa concluída", ProcessType.FILE)
                # Check if file already exists in ARQUIVOS_REVISADOS and remove it
                arquivo_revisado = os.path.join(rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS", os.path.basename(arquivo))
                if os.path.exists(arquivo_revisado):
                    logger.warning(f"Arquivo já existe em ARQUIVOS_REVISADOS, removendo arquivo existente...", ProcessType.FILE)
                    os.remove(arquivo_revisado)
                
                shutil.move(os.path.join(caminho_input, arquivo), rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                insert_log_data(arquivo, "OK")
                logger.success("Arquivos movidos com sucesso e log atualizado com status OK", ProcessType.DATABASE)
            else:
                logger.error("Movendo arquivo para pasta ERRO devido a falha no processamento...", ProcessType.FILE)
                shutil.move(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\ERRO")
                os.remove(os.path.join(caminho_input, arquivo))
                insert_log_data(arquivo, "ERRO")
                logger.error("Arquivo movido para pasta ERRO e log atualizado", ProcessType.DATABASE)
        case ".xlsx":
            logger.info("Processando arquivo Excel...", ProcessType.EXCEL)
            backup_path = rf"{get_caminho_rede()}\ARQUIVOS_BACKUP"
            logger.info(f"Criando backup em: {backup_path}", ProcessType.FILE)
            shutil.copy(arquivo_novo, backup_path)
            
            logger.info("Verificando tipo do arquivo Excel...", ProcessType.EXCEL)
            excel_type_verification = verify_excel_type(arquivo_novo)
            logger.info(f"Tipo de Excel identificado: {excel_type_verification}", ProcessType.EXCEL)
            
            if excel_type_verification == "TYPE_A":
                logger.info("Processando Excel TYPE_A...", ProcessType.EXCEL)
                logger.info("Editando códigos conforme tabela DE-PARA...", ProcessType.EXCEL)
                edit_excel_codigo(arquivo_novo, get_caminho_de_para(), excel_type_verification)
                logger.info("Adicionando informação de revisão...", ProcessType.EXCEL)
                adicionar_revisao(arquivo_novo, "Revisão dos documentos mediante ao CM-TBS-00728")
                resultado = "Processado"
                logger.info("Movendo arquivos para as pastas correspondentes...", ProcessType.FILE)
                # DEFINE FOLDER PROCESSADOS IN  C:\RPA\RPA001_Garantia_De_Qualidade\data\PROCESSADOS
                pasta_destino = rf"{get_caminho_rede()}\PROCESSADOS"
                
                shutil.copy(os.path.join(backup_path, arquivo), pasta_destino)
                shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                
                if os.path.exists(arquivo_novo):
                    logger.warning(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...", ProcessType.FILE)
                    os.remove(arquivo_novo)
                
                insert_log_data(arquivo, "OK")
                logger.success("Excel TYPE_A processado com sucesso", ProcessType.EXCEL)
            elif excel_type_verification == "TYPE_B":
                logger.info("Processando Excel TYPE_B...", ProcessType.EXCEL)
                logger.info("Editando códigos conforme tabela DE-PARA...", ProcessType.EXCEL)
                edit_excel_codigo(arquivo_novo, get_caminho_de_para(), excel_type_verification)
                resultado = "Processado"
                logger.info("Movendo arquivos para as pastas correspondentes...", ProcessType.FILE)
                shutil.copy(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\PROCESSADOS")
                shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                
                if os.path.exists(arquivo_novo):
                    logger.warning(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...", ProcessType.FILE)
                    os.remove(arquivo_novo)
                
                insert_log_data(arquivo, "OK")
                logger.success("Excel TYPE_B processado com sucesso", ProcessType.EXCEL)
            elif excel_type_verification == "TYPE_C" or excel_type_verification == "TYPE_D" or excel_type_verification == "TYPE_REVISION":
                logger.info(f"Processando Excel {excel_type_verification}...", ProcessType.EXCEL)
                logger.info("Atualizando rodapé do Excel...", ProcessType.EXCEL)
                update_excel_footer(arquivo_novo)
                resultado = "Processado"
                logger.info("Movendo arquivos para as pastas correspondentes...", ProcessType.FILE)
                shutil.copy(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\PROCESSADOS")
                shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                
                if os.path.exists(arquivo_novo):
                    logger.warning(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...", ProcessType.FILE)
                    os.remove(arquivo_novo)
                
                insert_log_data(arquivo, "OK")
                logger.success(f"Excel {excel_type_verification} processado com sucesso", ProcessType.EXCEL)
                
            else:
                logger.error(f"Tipo Excel Desconhecido: {arquivo}", ProcessType.EXCEL)
                resultado = "Tipo de arquivo desconhecido"
                logger.warning("Movendo arquivo para pasta ERRO...", ProcessType.FILE)
                
                shutil.copy(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\ERRO")

                if os.path.exists(arquivo_novo):
                    logger.warning(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...", ProcessType.FILE)
                    os.remove(arquivo_novo)
                
                os.remove(arquivo_novo)
                insert_log_data(arquivo, "ERRO")
                logger.success("Arquivo movido para pasta ERRO e log atualizado", ProcessType.DATABASE)
        case _:
            logger.error(f"Tipo de arquivo desconhecido: {arquivo}", ProcessType.FILE)
            shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ERRO")
            insert_log_data(arquivo, "ERRO")
            logger.success("Arquivo copiado para pasta ERRO e log atualizado", ProcessType.DATABASE)

    if str(resultado) != "Tipo de arquivo desconhecido" and str(resultado) != "Erro ao Processar":
        resultado = "Processado"

        return resultado

# Função principal
def main():
    logger.info("Iniciando processo de automação RPA001", ProcessType.SYSTEM)
    print("\n" + "="*50)
    print("INICIANDO PROCESSO DE AUTOMAÇÃO RPA001")
    print("="*50)
    print(f"Data e hora de início: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    carregar_configuracoes()
    print("")
    
    # Definir os caminhos de entrada das pastas
    pastas = [
        rf"{get_caminho_rede()}\METODOS-ANEXOS",
        rf"{get_caminho_rede()}\IF-IE",
        rf"{get_caminho_rede()}\LISTA-DE-FORNECEDORES"
    ]
    logger.info("Procurando arquivos nas pastas definidas", ProcessType.FILE)
    
    # Iterar sobre as pastas e listar seus nomes
    for pasta in pastas:
        logger.info(f"Pasta: {pasta}", ProcessType.FILE)
    print("")
    
    # Listar os arquivos nas pastas especificadas
    arquivos = listar_arquivos(pastas)
    print("")
    
    if arquivos:
        logger.info("Carregando tabela DE-PARA...", ProcessType.DATABASE)
        caminho_arquivo_excel = get_caminho_de_para()
        df = read_excel_file(caminho_arquivo_excel)
        logger.success(f"Tabela DE-PARA carregada: {len(df)} registros", ProcessType.DATABASE)        
        
        print("\n" + "-"*50)
        logger.info(f"Iniciando processamento de {len(arquivos)} arquivos", ProcessType.BUSINESS)
        print("-"*50)
        
        for arquivo in arquivos:
            # Definir o caminho completo para o arquivo
            caminho_input = None
            
            for pasta in pastas:
                if os.path.basename(arquivo) in os.listdir(pasta):
                    caminho_input = pasta
                    break

            if caminho_input is None:
                logger.error(f"Caminho de entrada não encontrado para o arquivo {arquivo}", ProcessType.FILE)
                continue
            
            processar_arquivo(arquivo, caminho_input, df)
    else:
        logger.info("Nenhum arquivo encontrado nas pastas especificadas", ProcessType.FILE)
    print("\n" + "-"*50)
    logger.success("Processamento de arquivos concluído", ProcessType.BUSINESS)
    print("-"*50)


if __name__ == "__main__":

    main()

    # Lista de valores
    tipos = ["EME", "EMP", "EPA", "EPE", "EPI", "MAME", "MAMP", "MAPA", "MAPE", "MAPI"]

    # Gerar a parte do LIKE dinamicamente
    like_conditions = " OR ".join([f"nomearquivo LIKE '%{tipo}%'" for tipo in tipos])

    # Gerar o CASE dinamicamente
    case_conditions = "\n".join([f"WHEN nomearquivo LIKE '%{tipo}%' THEN {i + 1}" for i, tipo in enumerate(tipos)])

    # Query final
    query = f"""
    SELECT *
    FROM public.rpa001_controle_execucao
    WHERE statusrevisao = 'OK'
    AND statusenviadosesuite = ''
    AND ({like_conditions})
    ORDER BY 
    substring(nomearquivo FROM '[0-9]+')::int, -- Ordena pelo código numérico
    CASE 
        {case_conditions}
        ELSE {len(tipos) + 1}
    END,
    nomearquivo -- Ordem alfabética para desempate
    ;
    """


    print("\n" + "="*50)
    logger.info("Verificando arquivos para upload no SeSuite", ProcessType.DATABASE)
    print("="*50)
    print("Executando query:")
    print(query)
    
    df = query_to_dataframe(query)
    
    if not df.empty:
        logger.success(f"{len(df)} arquivos encontrados para upload no SeSuite", ProcessType.DATABASE)
        print("\n" + "="*50)
        logger.info("Iniciando upload para SeSuite", ProcessType.SELENIUM)
        print("="*50)
        NavegationSeSuite(df)        
        print("\n" + "="*50)
        logger.success("Upload para SeSuite finalizado", ProcessType.SELENIUM)
        print("="*50)
    else:
        logger.info("Nenhum arquivo encontrado para upload no SeSuite", ProcessType.DATABASE)

    query = f"""SELECT *
        FROM public.rpa001_controle_execucao
        WHERE statusrevisao = 'OK'
        AND statusenviadosesuite = 'OK'
        AND statushomologado = ''
        AND ({like_conditions})
        ORDER BY 
        substring(nomearquivo FROM '[0-9]+')::int, -- Ordena pelo código numérico
        CASE 
            {case_conditions}
            ELSE {len(tipos) + 1}
        END,
        nomearquivo -- Ordem alfabética para desempate
        ;
    """

    
    print("\n" + "="*50)
    logger.info("Verificando arquivos para homologação no SeSuite", ProcessType.DATABASE)
    print("="*50)
    print("Executando query:")
    print(query)
    
    df = query_to_dataframe(query)

    if not df.empty:
        logger.success(f"{len(df)} arquivos encontrados para homologação no SeSuite", ProcessType.DATABASE)        
        print("\n" + "="*50)
        logger.info("Iniciando homologação no SeSuite", ProcessType.SELENIUM)
        print("="*50)
        HomologacaoSeSuite(df)
        print("\n" + "="*50)
        logger.success("Homologação no SeSuite finalizada", ProcessType.SELENIUM)
        print("="*50)
    else:
        logger.info("Nenhum arquivo encontrado para homologação no SeSuite", ProcessType.DATABASE)

    query = """SELECT *
        FROM public.rpa001_controle_execucao
        WHERE COALESCE(statusrevisao, '') <> ''
        AND COALESCE(statusenviadosesuite, '') <> ''
        AND COALESCE(statushomologado, '') <> ''
    AND relatorio = '0'; 
    """

print("\n" + "="*50)
logger.info("Verificando registros para geração de relatório", ProcessType.DATABASE)
print("="*50)
print("Executando query:")
print(query)

df = query_to_dataframe(query)

if not df.empty:
    logger.success(f"{len(df)} registros encontrados para geração de relatório", ProcessType.DATABASE)
    gerar_relatorio_e_enviar_email(df)
    
    # Atualizar o status de relatório para evitar duplicação
    try:
        # Construir lista de IDs para atualização
        ids = ', '.join([str(row['id']) for _, row in df.iterrows()])
        update_query = f"""
        UPDATE public.rpa001_controle_execucao
        SET relatorio = '1'
        WHERE id IN ({ids});
        """
        
        # Executar a atualização
        with create_engine(get_database_url()).connect() as connection:
            connection.execute(text(update_query))
            connection.commit()
            
        logger.success("Status de relatório atualizado no banco de dados", ProcessType.DATABASE)
    except Exception as e:
        logger.error(f"Erro ao atualizar status de relatório: {e}", ProcessType.DATABASE)
else:
    logger.info("Nenhum registro encontrado para geração de relatório", ProcessType.DATABASE)

print("\n" + "="*50)
logger.success("Processo de automação RPA001 concluído", ProcessType.SYSTEM)
print(f"Data e hora de término: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print("="*50 + "\n")
