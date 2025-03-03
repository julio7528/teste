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
from src.config.config import load_config, get_database_url, get_caminho_rede, get_caminho_de_para, get_url_hml, get_user_SeSuite, get_password_SeSuite, generate_default_foldes, get_contra_senha
from src.utils.read_files_utils import capture_type_from_headers, capture_code_from_docx, capture_code_from_headers, convert_doc_to_docx, convert_docx_to_doc, read_excel_file, captura_tabela_embalagem, captura_codigo_nucleo, captura_tabela_fabricacao
from src.utils.edit_files_utiles import edit_file_ficha, edit_file_eme, edit_file_mame, edit_file_embalagem_fabricacao, adicionar_nova_linha_com_codigo_embalagem, substituir_codigo_nucleo,adicionar_nova_linha_fabricacao, editar_celula_codigo_embalagem
from src.utils.edit_files_utils_excel import edit_excel_codigo, adicionar_revisao, update_excel_footer
from src.utils.verify_excel_type import verify_excel_type
from src.navigantion.base_page import BasePage
from src.navigantion.login_page import LoginPage
from src.navigantion.upload_seSuite import NavegationSeSuite
from src.navigantion.homologacao_seSuite import HomologacaoSeSuite
from src.services.db_service import insert_log_data, query_to_dataframe, update_log_data

os.system('cls' if os.name == 'nt' else 'clear')

# Função para carregar configurações
def carregar_configuracoes():    
    try:
        print("\n" + "-"*50)
        print("CARREGANDO CONFIGURAÇÕES DO SISTEMA")
        print("-"*50)
        load_config()  # Assuming load_config correctly loads the configurations
        generate_default_foldes()
        print("✓ Configurações carregadas com sucesso.")
    except Exception as e:
        print(f"❌ ERRO ao carregar configurações: {e}")

# Função para listar os arquivos nas pastas
def listar_arquivos(pastas):
    print("\n" + "-"*50)
    print("PROCURANDO ARQUIVOS PARA PROCESSAMENTO")
    print("-"*50)
    arquivos = []
    total_arquivos = 0
    
    for pasta in pastas:
        arquivos_pasta = 0
        print(f"Verificando pasta: '{pasta}'...")
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
                        print(f"-- Arquivo encontrado: {item}")
                print(f"✓ {arquivos_pasta} arquivo(s) encontrado(s) em '{pasta}'")
            else:
                print(f"⚠️ AVISO: '{pasta}' não é uma pasta válida ou não existe.")

        except FileNotFoundError:
            print(f"❌ ERRO: A pasta '{pasta}' não foi encontrada.")
        except OSError as e:
            print(f"❌ ERRO ao acessar a pasta '{pasta}': {e}")
        except Exception as e:
            print(f"❌ ERRO inesperado na pasta '{pasta}': {e}")

    print(f"\n📁 Total de arquivos encontrados: {total_arquivos}")
    return arquivos

def gerar_relatorio_e_enviar_email(df):
    print("\n" + "-"*50)
    print("GERANDO RELATÓRIO DE EXECUÇÃO")
    print("-"*50)
    print(f"Gerando relatório para {len(df)} registros...")

    # Salvar os dados no Excel
    relatorio_excel = "relatorio_rpa001.xlsx"
    df.to_excel(relatorio_excel, index=False, engine="openpyxl")
    print(f"✓ Relatório salvo em: {relatorio_excel}")

    # Configurações do e-mail
    remetente = "seu_email@example.com"
    senha = "sua_senha"
    destinatario = "destinatario@example.com"
    assunto = "Relatório RPA001 Controle Execução"

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
        print("Enviando e-mail...")
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(remetente, senha)
            server.sendmail(remetente, destinatario, msg.as_string())
        print("✓ E-mail enviado com sucesso!")
        print(f"  - De: {remetente}")
        print(f"  - Para: {destinatario}")
        print(f"  - Assunto: {assunto}")
        print(f"  - Anexo: {relatorio_excel}")
    except Exception as e:
        print(f"❌ ERRO ao enviar e-mail: {e}")

# Função para editar o arquivo conforme o tipo
def editar_arquivo(arquivo, codigo_word, novo_codigo, tipoDocumento, df_para):
    try:
        print(f"Identificando tipo do arquivo: {os.path.basename(arquivo)}")
        type_file = capture_type_from_headers(arquivo)
        print(f"Tipo identificado: {type_file}")
        print(f"Substituindo código: {codigo_word} → {novo_codigo}")
        
        if type_file == "FICHA DE ANÁLISE":
            print("Processando FICHA DE ANÁLISE...")
            edit_file_ficha(arquivo, codigo_word, f"{codigo_word} ({novo_codigo})", arquivo)
            print("✓ FICHA DE ANÁLISE processada com sucesso")
            return "Processado"
        elif type_file == "EME" or type_file == "EMP" or type_file == "EPA" or type_file == "EPE" or type_file == "EPI":
            print(f"Processando {type_file}...")
            edit_file_eme(arquivo, codigo_word, f"{codigo_word} ({novo_codigo})", arquivo)
            print(f"✓ {type_file} processado com sucesso")
            return "Processado"
        elif type_file == "MAME" or type_file == "MAMP" or type_file == "MAPA" or type_file == "MAPE" or type_file == "MAPI":
            print(f"Processando {type_file}...")
            edit_file_mame(arquivo, codigo_word, f"{codigo_word} ({novo_codigo})", arquivo)
            print(f"✓ {type_file} processado com sucesso")
            return "Processado"
        elif tipoDocumento=="INSTRUÇÃO DE EMBALAGEM":
            print("Processando INSTRUÇÃO DE EMBALAGEM...")
            df_embalagem = captura_tabela_embalagem(arquivo, "Componentes – Material de Embalagem")
            print(f"✓ Tabela de embalagem capturada: {len(df_embalagem)} linhas")

            novas_linhas = []
            for item in df_embalagem:
                codigo_item = item['Código']
                print(f"Processando código de embalagem: {codigo_item}")
                codigo_para, descricao_item = buscar_codigo_para(codigo_item, df_para)

                if codigo_para != "Código não encontrado":
                    print(f"Editando célula: {codigo_item} → {codigo_para}")
                    editar_celula_codigo_embalagem(arquivo, str(codigo_item), str(codigo_para), arquivo)
            
            if codigo_word.startswith(('3', '5', '9')):
                print(f"Editando cabeçalho com substituição direta: {codigo_word} → {novo_codigo}")
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), str(novo_codigo), arquivo)
            else:
                print(f"Editando cabeçalho com código duplo: {codigo_word} → {codigo_word} ({novo_codigo})")
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), f"{str(codigo_word)} ({str(novo_codigo)})", arquivo)
            
            print("✓ INSTRUÇÃO DE EMBALAGEM processada com sucesso")
            return "Processado"

        elif tipoDocumento=="INSTRUÇÃO DE FABRICAÇÃO":
            print("Processando INSTRUÇÃO DE FABRICAÇÃO...")
            if codigo_word.startswith(('3', '5', '9')):
                print(f"Editando cabeçalho com substituição direta: {codigo_word} → {novo_codigo}")
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), str(novo_codigo), arquivo)
            else:
                print(f"Editando cabeçalho com código duplo: {codigo_word} → {codigo_word} ({novo_codigo})")
                edit_file_embalagem_fabricacao(arquivo, str(codigo_word), f"{str(codigo_word)} ({str(novo_codigo)})", arquivo)

            codigos = captura_codigo_nucleo(arquivo, "Componentes – Núcleo")
            print(f"✓ Códigos núcleo capturados: {len(codigos)} códigos")

            for cod in codigos:
                print(f"Processando código núcleo: {cod}")
                codigo_para, descricao_item = buscar_codigo_para(cod, df_para)

                if codigo_para != "Código não encontrado":
                    if codigo_word.startswith(('3')):
                        print(f"Substituindo código núcleo diretamente: {cod} → {codigo_para}")
                        substituir_codigo_nucleo(arquivo, str(cod), str(codigo_para), arquivo)                       
                    else:
                        print(f"Substituindo código núcleo com código duplo: {cod} → {cod} ({codigo_para})")
                        substituir_codigo_nucleo(arquivo, cod, f"{str(cod)} ({str(codigo_para)})", arquivo)

            df_fabricacao = captura_tabela_fabricacao(arquivo, "Componentes – Núcleo")
            print(f"✓ Tabela de fabricação capturada: {len(df_fabricacao)} linhas")
            
            novas_linhas = []
            for item in df_fabricacao:
                codigo_item = item['Código']
                print(f"Processando código de fabricação: {codigo_item}")
                codigo_para, descricao_item = buscar_codigo_para(codigo_item, df_para)

                if codigo_para != "Código não encontrado":
                    print(f"Editando célula: {codigo_item} → {codigo_para}")
                    editar_celula_codigo_embalagem(arquivo, str(codigo_item), str(codigo_para), arquivo)

            print("✓ INSTRUÇÃO DE FABRICAÇÃO processada com sucesso")
            return "Processado"
        else:
            print(f"❌ Tipo de arquivo desconhecido: {arquivo}")
            return "Tipo de arquivo desconhecido"
    except Exception as e:
        print(f"❌ ERRO ao editar o arquivo {os.path.basename(arquivo)}: {e}")
        return "Erro ao Processar"

# Função para buscar código na tabela
def buscar_codigo_para(codigo_de, df):
    print(f"Buscando correspondência para código: {codigo_de}")
    try:
        # Filtrar o DataFrame com base no 'Codigo - DE'
        resultado = df.loc[df['Codigo - DE'].astype(str) == str(codigo_de)]
        
        # Verificar se a busca retornou resultados
        if not resultado.empty:
            codigo_para = resultado['Codigo -  PARA'].iloc[0]
            descricao_item = resultado.iloc[0, -1]  # Acessa a última coluna da linha
            print(f"✓ Correspondência encontrada: {codigo_de} → {codigo_para}")
            return codigo_para, descricao_item
        else:
            print(f"⚠️ Código {codigo_de} não encontrado na tabela DE-PARA")
            return 'Código não encontrado', None
    except Exception as e:
        print(f"❌ ERRO ao buscar código {codigo_de}: {e}")
        return f'Erro na busca de código {e}', None

# Função para processar arquivos .doc ou .docx
def processar_arquivo(arquivo, caminho_input, df):
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
        print(f"❌ ERRO ao preparar o arquivo: {e}")
        resultado = "Erro ao Processar"
    
    match extensao:
        case ".doc" | ".docx":
            backup_path = rf"{get_caminho_rede()}\ARQUIVOS_BACKUP"
            print(f"Criando backup em: {backup_path}")
            shutil.copy(os.path.join(caminho_input, arquivo), backup_path)
            converted = False
            if extensao == ".doc":
                print(f"Convertendo .doc para .docx: {arquivo}")
                arquivo_doc = os.path.join(caminho_input, arquivo)
                convert_doc_to_docx(arquivo_doc, arquivo_novo)
                converted = True
                print("✓ Conversão concluída")
            
            print("Capturando código do documento...")
            codigo_word = capture_code_from_docx(arquivo_novo)
            tipoDocumento = ""
            
            if codigo_word == "Nenhum código encontrado.":
                print("Código não encontrado no conteúdo, tentando capturar dos cabeçalhos...")
                codigo_word, tipoDocumento = capture_code_from_headers(arquivo_novo)
                print(f"Tipo de documento identificado: {tipoDocumento}")

                if codigo_word == "Nenhum código encontrado.":
                    print(f"❌ Código não encontrado no arquivo: {arquivo_novo}")
                    print(f"Movendo arquivo para pasta ERRO...")
                    shutil.move(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\ERRO")
                    insert_log_data(arquivo, "ERRO")
                    print("✓ Arquivo movido para pasta ERRO e log atualizado")
                    return
                
            print(f"Código encontrado: {codigo_word}")
            ultimo_codigo = re.split(r'[\/,]', codigo_word)[-1].strip()
            print(f"Último código a ser processado: {ultimo_codigo}")
            
            print("Buscando correspondência na tabela DE-PARA...")
            novo_codigo, descricao_item = buscar_codigo_para(ultimo_codigo, df)

            if novo_codigo == "Código não encontrado":
                print(f"❌ Correspondência não encontrada para: {ultimo_codigo}")
                print(f"Movendo arquivo para pasta ERRO...")
                dir_destino = rf"{get_caminho_rede()}\ERRO"
                arq_destino = os.path.join(dir_destino, arquivo) # Caminho de destino inicial com nome original

                try: # Bloco try-except simplificado
                    if os.path.exists(arq_destino):
                        print(f"⚠️ Arquivo '{arquivo}' já existe em '{dir_destino}'. Renomeando...")
                        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                        arquivo_novo = f"{timestamp}_{arquivo}"
                        arq_destino = os.path.join(dir_destino, arquivo_novo) # **Atualiza arq_destino com o novo nome**
                        print(f"Novo nome para o arquivo será: '{arquivo_novo}'")

                    print(f"Movendo arquivo '{arquivo}' para '{arq_destino}'...") # Print antes de mover
                    shutil.move(os.path.join(backup_path, arquivo), arq_destino) # **Usa arq_destino como destino**
                    print(f"✓ Arquivo '{arquivo}' movido com sucesso para '{arq_destino}'.") # Print de sucesso
                    insert_log_data(arquivo, "ERRO") # Log ainda usa o 'arquivo' original, se for desejado
                    print("✓ Log atualizado")

                except Exception as e: # Captura qualquer erro durante a movimentação
                    print(f"❌ ERRO ao mover arquivo '{arquivo}' para '{arq_destino}'. Erro: {e}")
                    insert_log_data(arquivo, f"ERRO - Movimentação: {e}") # Log de erro genérico

                return

            print(f"Iniciando edição do arquivo: {arquivo_novo}")
            resultado = editar_arquivo(arquivo_novo, codigo_word, novo_codigo, tipoDocumento, df)
            print(f"Resultado da edição: {resultado}")
            
            if resultado == "Processado":
                print("Movendo arquivos para as pastas correspondentes...")
                pasta_destino = rf"{get_caminho_rede()}\PROCESSADOS"
                print(f"Movendo arquivo para: {pasta_destino}")
                
                if os.path.exists(os.path.join(pasta_destino, os.path.basename(arquivo))):
                    print(f"Arquivo já existe no destino, removendo arquivo existente...")
                    os.remove(os.path.join(pasta_destino, os.path.basename(arquivo)))
                shutil.move(os.path.join(backup_path, arquivo), pasta_destino)
                
                if converted:
                    print("Convertendo .docx de volta para .doc...")
                    convert_docx_to_doc(arquivo_novo, os.path.join(caminho_input, arquivo))
                    os.remove(arquivo_novo)
                    print("✓ Conversão reversa concluída")
                # Check if file already exists in ARQUIVOS_REVISADOS and remove it
                arquivo_revisado = os.path.join(rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS", os.path.basename(arquivo))
                if os.path.exists(arquivo_revisado):
                    print(f"Arquivo já existe em ARQUIVOS_REVISADOS, removendo arquivo existente...")
                    os.remove(arquivo_revisado)
                
                shutil.move(os.path.join(caminho_input, arquivo), rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                insert_log_data(arquivo, "OK")
                print("✓ Arquivos movidos com sucesso e log atualizado com status OK")
            else:
                print("Movendo arquivo para pasta ERRO devido a falha no processamento...")
                shutil.move(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\ERRO")
                os.remove(os.path.join(caminho_input, arquivo))
                insert_log_data(arquivo, "ERRO")
                print("✓ Arquivo movido para pasta ERRO e log atualizado")
        case ".xlsx":
            print("Processando arquivo Excel...")
            backup_path = rf"{get_caminho_rede()}\ARQUIVOS_BACKUP"
            print(f"Criando backup em: {backup_path}")
            shutil.copy(arquivo_novo, backup_path)
            
            print("Verificando tipo do arquivo Excel...")
            excel_type_verification = verify_excel_type(arquivo_novo)
            print(f"Tipo de Excel identificado: {excel_type_verification}")
            
            if excel_type_verification == "TYPE_A":
                print("Processando Excel TYPE_A...")
                print("Editando códigos conforme tabela DE-PARA...")
                edit_excel_codigo(arquivo_novo, get_caminho_de_para(), excel_type_verification)
                print("Adicionando informação de revisão...")
                adicionar_revisao(arquivo_novo, "Revisão dos documentos mediante ao CM-TBS-00728")
                resultado = "Processado"
                print("Movendo arquivos para as pastas correspondentes...")
                # DEFINE FOLDER PROCESSADOS IN  C:\RPA\RPA001_Garantia_De_Qualidade\data\PROCESSADOS
                pasta_destino = rf"{get_caminho_rede()}\PROCESSADOS"
                
                shutil.copy(os.path.join(backup_path, arquivo), pasta_destino)
                shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                
                if os.path.exists(arquivo_novo):
                    print(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...")
                    os.remove(arquivo_novo)
                
                insert_log_data(arquivo, "OK")
                print("✓ Excel TYPE_A processado com sucesso")
            elif excel_type_verification == "TYPE_B":
                print("Processando Excel TYPE_B...")
                print("Editando códigos conforme tabela DE-PARA...")
                edit_excel_codigo(arquivo_novo, get_caminho_de_para(), excel_type_verification)
                resultado = "Processado"
                print("Movendo arquivos para as pastas correspondentes...")
                shutil.copy(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\PROCESSADOS")
                shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                
                if os.path.exists(arquivo_novo):
                    print(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...")
                    os.remove(arquivo_novo)
                
                insert_log_data(arquivo, "OK")
                print("✓ Excel TYPE_B processado com sucesso")
            elif excel_type_verification == "TYPE_C" or excel_type_verification == "TYPE_D" or excel_type_verification == "TYPE_REVISION":
                print(f"Processando Excel {excel_type_verification}...")
                print("Atualizando rodapé do Excel...")
                update_excel_footer(arquivo_novo)
                resultado = "Processado"
                print("Movendo arquivos para as pastas correspondentes...")
                shutil.copy(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\PROCESSADOS")
                shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ARQUIVOS_REVISADOS")
                
                if os.path.exists(arquivo_novo):
                    print(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...")
                    os.remove(arquivo_novo)
                
                insert_log_data(arquivo, "OK")
                print(f"✓ Excel {excel_type_verification} processado com sucesso")
                        
            else:
                print(f"❌ Tipo Excel Desconhecido: {arquivo}")
                resultado = "Tipo de arquivo desconhecido"
                print("Movendo arquivo para pasta ERRO...")
                
                shutil.copy(os.path.join(backup_path, arquivo), rf"{get_caminho_rede()}\ERRO")

                if os.path.exists(arquivo_novo):
                    print(f"Arquivo {arquivo_novo} já existe em ARQUIVOS_REVISADOS e PROCESSADOS, removendo arquivo existente EM LISTA-DE-FORNECEDORES...")
                    os.remove(arquivo_novo)
                
                os.remove(arquivo_novo)
                insert_log_data(arquivo, "ERRO")
                print("✓ Arquivo movido para pasta ERRO e log atualizado")
        case _:
            print(f"❌ Tipo de arquivo desconhecido: {arquivo}")
            shutil.copy(arquivo_novo, rf"{get_caminho_rede()}\ERRO")
            insert_log_data(arquivo, "ERRO")
            print("✓ Arquivo copiado para pasta ERRO e log atualizado")

    if str(resultado) != "Tipo de arquivo desconhecido" and str(resultado) != "Erro ao Processar":
        resultado = "Processado"

    return resultado

# Função principal
def main():
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
    print("Procurando arquivos nas pastas:")   
    
    # Iterar sobre as pastas e listar seus nomes
    for pasta in pastas:
        print(f"- {pasta}")
    print("")
    
    # Listar os arquivos nas pastas especificadas
    arquivos = listar_arquivos(pastas)
    print("")
    
    if arquivos:
        print(f"Carregando tabela DE-PARA...")
        caminho_arquivo_excel = get_caminho_de_para()
        df = read_excel_file(caminho_arquivo_excel)
        print(f"✓ Tabela DE-PARA carregada: {len(df)} registros")
        
        print("\n" + "-"*50)
        print(f"INICIANDO PROCESSAMENTO DE {len(arquivos)} ARQUIVOS")
        print("-"*50)
        
        for arquivo in arquivos:
            # Definir o caminho completo para o arquivo
            caminho_input = None
            
            for pasta in pastas:
                if os.path.basename(arquivo) in os.listdir(pasta):
                    caminho_input = pasta
                    break

            if caminho_input is None:
                print(f"❌ ERRO: Caminho de entrada não encontrado para o arquivo {arquivo}")
                continue
            
            processar_arquivo(arquivo, caminho_input, df)
    else:
        print("ℹ️ Nenhum arquivo encontrado nas pastas especificadas.")

    print("\n" + "-"*50)
    print("PROCESSAMENTO DE ARQUIVOS CONCLUÍDO")
    print("-"*50)


if __name__ == "__main__":

    main() #TODO: STARTING SETUP FILE TASKKILL RELATED APPS AND CHECK IS IN MEMORY HAVE OPENED MS RELATED FILE

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
    print("VERIFICANDO ARQUIVOS PARA UPLOAD NO SESUITE")
    print("="*50)
    print("Executando query:")
    print(query)
    
    df = query_to_dataframe(query)
    
    if not df.empty:
        print(f"✓ {len(df)} arquivos encontrados para upload no SeSuite")
        print("\n" + "="*50)
        print("INICIANDO UPLOAD PARA SESUITE")
        print("="*50)
        NavegationSeSuite(df)
        print("\n" + "="*50)
        print("UPLOAD PARA SESUITE FINALIZADO")
        print("="*50)
    else:
        print("ℹ️ Nenhum arquivo encontrado para upload no SeSuite")

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
    print("VERIFICANDO ARQUIVOS PARA HOMOLOGAÇÃO NO SESUITE")
    print("="*50)
    print("Executando query:")
    print(query)
    
    df = query_to_dataframe(query)

    if not df.empty:
        print(f"✓ {len(df)} arquivos encontrados para homologação no SeSuite")
        print("\n" + "="*50)
        print("INICIANDO HOMOLOGAÇÃO NO SESUITE")
        print("="*50)
        HomologacaoSeSuite(df)
        print("\n" + "="*50)
        print("HOMOLOGAÇÃO NO SESUITE FINALIZADA")
        print("="*50)
    else:
        print("ℹ️ Nenhum arquivo encontrado para homologação no SeSuite")

    query = """SELECT *
        FROM public.rpa001_controle_execucao
        WHERE COALESCE(statusrevisao, '') <> ''
        AND COALESCE(statusenviadosesuite, '') <> ''
        AND COALESCE(statushomologado, '') <> ''
        AND relatorio = '0'; 
        """
    
    print("\n" + "="*50)
    print("VERIFICANDO REGISTROS PARA GERAÇÃO DE RELATÓRIO")
    print("="*50)
    print("Executando query:")
    print(query)
    
    df = query_to_dataframe(query)

    if not df.empty:
        print(f"✓ {len(df)} registros encontrados para geração de relatório")
        gerar_relatorio_e_enviar_email(df)
    else:
        print("ℹ️ Nenhum registro encontrado para geração de relatório")

    print("\n" + "="*50)
    print("PROCESSO DE AUTOMAÇÃO RPA001 CONCLUÍDO")
    print(f"Data e hora de término: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("="*50 + "\n")