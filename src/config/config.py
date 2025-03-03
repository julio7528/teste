
from dotenv import load_dotenv
import os
from urllib.parse import quote
from src.utils.logger import ProcessType, initialize_logger, get_logger

#declarar uma variavel looger global


# Função para carregar variáveis de ambiente
def load_config():
    load_dotenv()
    # Initialize logger with database URL
    initialize_logger(get_database_url())
    logger = get_logger()
    logger.info("Carregando configurações do env", ProcessType.SYSTEM)


def get_environment():
    """Obtém o ambiente de execução (local ou supabase)."""
    return os.getenv("ENVIRONMENT", "local")

def get_user_database():
    """Obtém a variável USER_DATABASE."""
    env = get_environment()
    if env == "supabase":
        return os.getenv("SUPABASE_USER")
    return os.getenv("USER_DATABASE")

def get_senha_database():
    """Obtém a variável SENHA_DATABASE."""
    env = get_environment()
    if env == "supabase":
        return quote(os.getenv("SUPABASE_PASSWORD"))
    return quote(os.getenv("SENHA_DATABASE"))

def get_server_database():
    """Obtém a variável SERVER_DATABASE."""
    env = get_environment()
    if env == "supabase":
        return os.getenv("SUPABASE_HOST")
    return os.getenv("SERVER_DATABASE")

def get_database():
    """Obtém a variável DATABASE."""
    env = get_environment()
    if env == "supabase":
        return os.getenv("SUPABASE_DATABASE")
    return os.getenv("DATABASE")

def get_database_url():
    """Constrói a URL de conexão ao banco de dados."""
    env = get_environment()
    
    # Se disponível, use diretamente a URL completa do Supabase
    if env == "supabase" and os.getenv("SUPABASE_URL"):
        # Verificar se a URL é completa
        supabase_url = os.getenv("SUPABASE_URL")
        if supabase_url.startswith("postgresql://"):
            return supabase_url
    
    # Construção padrão da URL
    user = get_user_database()
    senha = get_senha_database()
    server = get_server_database()
    database = get_database()
    
    # Adiciona porta se estiver usando Supabase
    if env == "supabase":
        port = os.getenv("SUPABASE_PORT", "5432")
        return f"postgresql://{user}:{senha}@{server}:{port}/{database}"
    
    return f"postgresql://{user}:{senha}@{server}/{database}"

def get_caminho_rede():
    return os.getenv("CAMINHO_REDE")

def get_caminho_de_para():
    return os.getenv("CAMINHO_DE_PARA")

def get_url_hml():
    return os.getenv("URL_SESUITE_HML")

def get_user_SeSuite():
    return os.getenv("USER_SESUITE")

def get_password_SeSuite():
    return os.getenv("SENHA_SESUITE")

def get_contra_senha():
    return os.getenv("CONTRA_SENHA")

def generate_default_foldes(): 
    default_path = get_caminho_rede()

    foldes_to_create = [
        "DE-PARA",
        "ERRO",
        "IF-IE",
        "LISTA-DE-FORNECEDORES",
        "METODOS-ANEXOS",
        "PROCESSADOS",
        "ARQUIVOS_BACKUP",
        "ARQUIVOS_REVISADOS"
    ]

    if not os.path.exists(default_path):
        raise FileNotFoundError(f"O caminho '{get_caminho_rede()}' nao foi encontrado.")
    
    for folder in foldes_to_create:
        if not os.path.exists(fr"{default_path}\{folder}"):
            os.makedirs(fr"{default_path}\{folder}")
            logger.info(f"Caminho {folder} criado com sucesso.")
