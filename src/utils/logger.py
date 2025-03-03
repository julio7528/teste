import os
import sys
import time
import logging
import inspect
import traceback
import psutil
import pandas as pd
from datetime import datetime
from enum import Enum
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError

# Adicionar o diretório src ao path para poder importar módulos
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, '..')
sys.path.append(src_dir)

# Importar a config para acessar variáveis de ambiente
from config.config import load_config, get_database_url

# Carregar as configurações
load_config()

# Definir constantes
LOG_DIRECTORY = os.path.join(src_dir, '..', 'logs')
LOG_FILE_NAME = f"rpa001_{datetime.now().strftime('%Y%m%d')}.log"
LOG_TABLE_NAME = "rpa001_logs"

# Garantir que o diretório de logs exista
os.makedirs(LOG_DIRECTORY, exist_ok=True)

# Enums para padronizar os valores
class ProcessType(Enum):
    SYSTEM = "system"
    BUSINESS = "business"
    DATABASE = "database"
    FILE = "file"
    NETWORK = "network"
    SELENIUM = "selenium"
    INTERFACE = "interface"
    EXCEL = "excel"
    WORD = "word"

class LogStatus(Enum):
    INFO = "information"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"
    DEBUG = "debug"
    SUCCESS = "success"

class RPALogger:
    """
    Classe de log para o RPA001, que salva logs em arquivo e banco de dados.
    """
    
    def __init__(self):
        """Inicializa o logger."""
        self.task_name = os.getenv("RPA_TASK_NAME", "RPA001")
        
        # Configurar logger de arquivo
        self.file_logger = self._setup_file_logger()
        
        # Preparar conexão com banco de dados
        self.db_url = get_database_url()
        self.engine = None
        try:
            self.engine = create_engine(self.db_url)
            self._create_log_table_if_not_exists()
        except Exception as e:
            self.file_logger.error(f"Erro ao configurar conexão com banco de dados: {e}")
    
    def _setup_file_logger(self):
        """Configura o logger para salvar em arquivo."""
        logger = logging.getLogger("rpa001_logger")
        
        # Evitar duplicação de handlers
        if logger.handlers:
            return logger  # Se já tiver handlers, retorna o logger como está

        logger.setLevel(logging.DEBUG)
        
        # Criar o arquivo de log
        log_file_path = os.path.join(LOG_DIRECTORY, LOG_FILE_NAME)
        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        
        # Definir o formato do log no arquivo
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        
        # Removemos o handler do console para evitar a duplicação
        # Este foi o principal problema - não adicionamos mais o console_handler
        
        # Impedir a propagação para o logger raiz (que pode ter um handler de console)
        logger.propagate = False
        
        return logger
    
    def _create_log_table_if_not_exists(self):
        """Cria a tabela de log se não existir."""
        if not self.engine:
            return
            
        query = """
        CREATE TABLE IF NOT EXISTS public.{table_name} (
            id SERIAL PRIMARY KEY,
            timestamp TIMESTAMP NOT NULL,
            task VARCHAR(100) NOT NULL,
            function VARCHAR(255) NOT NULL,
            file VARCHAR(255) NOT NULL,
            message TEXT NOT NULL,
            process_type VARCHAR(50) NOT NULL,
            status VARCHAR(50) NOT NULL,
            cpu_usage DECIMAL(5,2),
            memory_usage DECIMAL(5,2)
        );
        
        CREATE INDEX IF NOT EXISTS idx_{table_name}_timestamp 
        ON public.{table_name}(timestamp);
        
        CREATE INDEX IF NOT EXISTS idx_{table_name}_status 
        ON public.{table_name}(status);
        
        CREATE INDEX IF NOT EXISTS idx_{table_name}_process_type 
        ON public.{table_name}(process_type);
        """.format(table_name=LOG_TABLE_NAME)
        
        try:
            with self.engine.begin() as connection:
                connection.execute(text(query))
        except SQLAlchemyError as e:
            self.file_logger.error(f"Erro ao criar tabela de logs: {e}")
    
    def _get_system_info(self):
        """Obtém informações do sistema como uso de CPU e memória."""
        cpu_usage = psutil.cpu_percent(interval=0.1)
        memory_usage = psutil.virtual_memory().percent
        return cpu_usage, memory_usage
    
    def _get_caller_info(self):
        """Obtém informações sobre a função e arquivo que chamou o logger."""
        stack = inspect.stack()
        # Pular o frame do próprio logger e pegar o chamador
        frame = stack[2]
        file_path = os.path.basename(frame.filename)
        function_name = frame.function
        return file_path, function_name
    
    def _log_to_database(self, message, process_type, status, file=None, function=None):
        """Salva o log no banco de dados."""
        if not self.engine:
            return
            
        try:
            # Obter informações do sistema
            cpu_usage, memory_usage = self._get_system_info()
            
            # Se file e function não foram fornecidos, tentar detectar automaticamente
            if not file or not function:
                detected_file, detected_function = self._get_caller_info()
                file = file or detected_file
                function = function or detected_function
                
            # Preparar os dados do log
            log_data = {
                "timestamp": datetime.now(),
                "task": self.task_name,
                "function": function,
                "file": file,
                "message": message,
                "process_type": process_type.value,
                "status": status.value,
                "cpu_usage": cpu_usage,
                "memory_usage": memory_usage
            }
            
            # Inserir no banco de dados usando transação
            query = f"""
            INSERT INTO public.{LOG_TABLE_NAME} 
            (timestamp, task, function, file, message, process_type, status, cpu_usage, memory_usage)
            VALUES 
            (:timestamp, :task, :function, :file, :message, :process_type, :status, :cpu_usage, :memory_usage)
            """
            
            # Usar with begin() para gerenciar a transação automaticamente
            with self.engine.begin() as connection:
                connection.execute(text(query), log_data)
                
        except Exception as e:
            # Falhar silenciosamente e logar apenas no arquivo
            self.file_logger.error(f"Erro ao salvar log no banco de dados: {e}")
    
    def _format_log_message(self, message, file, function, process_type, status, cpu_usage, memory_usage):
        """Formata a mensagem de log para o arquivo."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        return (f"{timestamp} | {self.task_name} | {function} | {file} | {message} | "
                f"{process_type.value} | {status.value} | CPU: {cpu_usage:.2f}% | MEM: {memory_usage:.2f}%")
    
    def log(self, message, process_type, status, file=None, function=None):
        """
        Registra uma mensagem de log no arquivo e no banco de dados.
        
        Args:
            message (str): Mensagem de log
            process_type (ProcessType): Tipo de processo
            status (LogStatus): Status do log
            file (str, optional): Nome do arquivo. Se None, será detectado automaticamente.
            function (str, optional): Nome da função. Se None, será detectado automaticamente.
        """
        # Obter informações do sistema
        cpu_usage, memory_usage = self._get_system_info()
        
        # Se file e function não foram fornecidos, tentar detectar automaticamente
        if not file or not function:
            detected_file, detected_function = self._get_caller_info()
            file = file or detected_file
            function = function or detected_function
            
        # Formatar a mensagem para o arquivo de log
        formatted_message = self._format_log_message(
            message, file, function, process_type, status, cpu_usage, memory_usage
        )
        
        # Vamos imprimir a mensagem formatada no console nós mesmos, para termos controle do formato
        print(formatted_message)
        
        # Logar no arquivo conforme o status
        if status == LogStatus.INFO:
            self.file_logger.info(formatted_message)
        elif status == LogStatus.WARNING:
            self.file_logger.warning(formatted_message)
        elif status == LogStatus.ERROR:
            self.file_logger.error(formatted_message)
        elif status == LogStatus.CRITICAL:
            self.file_logger.critical(formatted_message)
        elif status == LogStatus.DEBUG:
            self.file_logger.debug(formatted_message)
        elif status == LogStatus.SUCCESS:
            self.file_logger.info(formatted_message)  # Usar info para success
            
        # Salvar no banco de dados
        self._log_to_database(message, process_type, status, file, function)
    
    def info(self, message, process_type, file=None, function=None):
        """Registra uma mensagem informativa."""
        self.log(message, process_type, LogStatus.INFO, file, function)
    
    def warning(self, message, process_type, file=None, function=None):
        """Registra uma mensagem de aviso."""
        self.log(message, process_type, LogStatus.WARNING, file, function)
    
    def error(self, message, process_type, file=None, function=None):
        """Registra uma mensagem de erro."""
        self.log(message, process_type, LogStatus.ERROR, file, function)
    
    def critical(self, message, process_type, file=None, function=None):
        """Registra uma mensagem crítica."""
        self.log(message, process_type, LogStatus.CRITICAL, file, function)
    
    def debug(self, message, process_type, file=None, function=None):
        """Registra uma mensagem de depuração."""
        self.log(message, process_type, LogStatus.DEBUG, file, function)
    
    def success(self, message, process_type, file=None, function=None):
        """Registra uma mensagem de sucesso."""
        self.log(message, process_type, LogStatus.SUCCESS, file, function)
    
    def exception(self, message, process_type, file=None, function=None):
        """
        Registra uma exceção com stack trace.
        Deve ser chamado dentro de um bloco except.
        """
        exc_type, exc_value, exc_traceback = sys.exc_info()
        stack_trace = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        full_message = f"{message}\n{stack_trace}"
        self.log(full_message, process_type, LogStatus.ERROR, file, function)
    
    def get_logs(self, limit=100, status=None, process_type=None, start_date=None, end_date=None):
        """
        Recupera logs do banco de dados com filtros opcionais.
        
        Args:
            limit (int): Número máximo de registros a retornar
            status (LogStatus, optional): Filtrar por status
            process_type (ProcessType, optional): Filtrar por tipo de processo
            start_date (datetime, optional): Data de início
            end_date (datetime, optional): Data de fim
            
        Returns:
            DataFrame: DataFrame pandas com os logs encontrados
        """
        if not self.engine:
            self.file_logger.error("Conexão com banco de dados não disponível")
            return pd.DataFrame()
            
        try:
            # Construir a query base
            query = f"SELECT * FROM public.{LOG_TABLE_NAME} WHERE 1=1"
            params = {}
            
            # Adicionar filtros se fornecidos
            if status:
                query += " AND status = :status"
                params["status"] = status.value
                
            if process_type:
                query += " AND process_type = :process_type"
                params["process_type"] = process_type.value
                
            if start_date:
                query += " AND timestamp >= :start_date"
                params["start_date"] = start_date
                
            if end_date:
                query += " AND timestamp <= :end_date"
                params["end_date"] = end_date
                
            # Ordenar e limitar
            query += " ORDER BY timestamp DESC LIMIT :limit"
            params["limit"] = limit
            
            # Executar a query
            with self.engine.connect() as connection:
                result = connection.execute(text(query), params)
                df = pd.DataFrame(result.fetchall(), columns=result.keys())
                
            return df
            
        except Exception as e:
            self.file_logger.error(f"Erro ao recuperar logs: {e}")
            return pd.DataFrame()


# Instância global do logger
logger = RPALogger()

# Função para obter a instância do logger
def get_logger():
    """Retorna a instância global do logger."""
    return logger