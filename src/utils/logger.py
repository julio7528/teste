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

# Definir constantes
LOG_DIRECTORY = os.path.join(src_dir, '..', 'logs')
# file with date and time 
LOG_FILE_NAME = f"rpa001_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
    
    def __init__(self, db_url=None):
        """Inicializa o logger.
        
        Args:
            db_url (str, optional): URL de conexão com o banco de dados. Se None, 
                                o log no banco será desabilitado.
        """
        self.task_name = os.getenv("RPA_TASK_NAME", "RPA001")
        
        # Configurar logger de arquivo
        self.file_logger = self._setup_file_logger()
        
        self.column_widths = {
            'timestamp': 19,
            'level': 12,
            'task': 8,
            'function': 25,
            'file': 15,
            'message': 60,
            'process_type': 12,
            'status': 12,
            'cpu_usage': 9,
            'mem_usage': 5
            }
            
        # Preparar conexão com banco de dados
        self.db_url = db_url
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
        
        # Formatador personalizado que não faz nada, apenas passa a mensagem
        class PassThroughFormatter(logging.Formatter):
            def format(self, record):
                return record.getMessage()
        
        formatter = PassThroughFormatter()
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        
        # Impedir a propagação para o logger raiz
        logger.propagate = False
        
        # Definir larguras exatas das colunas para garantir alinhamento
        self.column_widths = {
            'timestamp': 19,
            'level': 12,  # Aumentado para acomodar "INFORMATION" e "SUCCESS"
            'task': 8,
            'function': 25,
            'file': 15,
            'message': 60,
            'process_type': 12,
            'status': 12,
            'cpu_usage': 5,
            'mem_usage': 5
        }
        
        # Adicionar cabeçalho ao arquivo de log se for um arquivo novo
        if not os.path.exists(log_file_path) or os.path.getsize(log_file_path) == 0:
            with open(log_file_path, 'w', encoding='utf-8') as f:
                # Cabeçalho com as larguras definidas e alinhadas
                header_row = (
                    f"{'TIMESTAMP'.ljust(self.column_widths['timestamp'])} | "
                    f"{'LEVEL'.ljust(self.column_widths['level'])} | "
                    f"{'TASK'.ljust(self.column_widths['task'])} | "
                    f"{'FUNCTION'.ljust(self.column_widths['function'])} | "
                    f"{'FILE'.ljust(self.column_widths['file'])} | "
                    f"{'MESSAGE'.ljust(self.column_widths['message'])} | "
                    f"{'PROCESS_TYPE'.ljust(self.column_widths['process_type'])} | "
                    f"{'STATUS'.ljust(self.column_widths['status'])} | "
                    f"{'CPU_USAGE'.ljust(self.column_widths['cpu_usage'])} | "
                    f"{'MEM_USAGE'.ljust(self.column_widths['mem_usage'])}"
                )
                f.write(header_row + "\n")
                
                # Linha separadora com exatamente o mesmo comprimento
                separator_line = "-" * len(header_row)
                f.write(separator_line + "\n")
        
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
            """
            Obtém informações sobre a função e arquivo que chamou o logger.
            Percorre a pilha de chamadas para encontrar o primeiro frame fora do módulo de logging.
            """
            stack = inspect.stack()
            for frame_info in stack[2:]:  # Começa do terceiro frame para pular os frames internos
                file_path = frame_info.filename
                
                # Ignora frames de dentro do módulo de logging
                if 'logger.py' not in file_path and 'log_viewer.py' not in file_path:
                    file_path = os.path.basename(file_path)
                    function_name = frame_info.function
                    return file_path, function_name
            
            # Se não encontrar um frame válido, usa o padrão
            frame = stack[2]
            file_path = os.path.basename(frame.filename)
            function_name = frame.function
            return file_path, function_name
    
    def _format_log_message(self, message, file, function, process_type, status, cpu_usage, memory_usage):
        """
        Formata a mensagem de log para o arquivo com colunas de largura fixa.
        
        Args:
            message (str): Mensagem de log
            file (str): Nome do arquivo
            function (str): Nome da função
            process_type (ProcessType): Tipo de processo
            status (LogStatus): Status do log
            cpu_usage (float): Uso de CPU
            memory_usage (float): Uso de memória
            
        Returns:
            str: Mensagem formatada
        """
        # Gera o timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Formata os valores com base nas larguras de coluna
        timestamp_str = timestamp.ljust(self.column_widths['timestamp'])
        level_str = status.value.upper().ljust(self.column_widths['level'])
        task_str = self.task_name.ljust(self.column_widths['task'])
        
        # Truncar campos longos
        function_str = function[:self.column_widths['function']].ljust(self.column_widths['function'])
        file_str = file[:self.column_widths['file']].ljust(self.column_widths['file'])
        process_type_str = process_type.value.ljust(self.column_widths['process_type'])
        status_str = status.value.ljust(self.column_widths['status'])
        
        # Formatação dos valores de CPU e memória
        cpu_str = f"{int(cpu_usage)}".ljust(self.column_widths['cpu_usage'])
        mem_str = f"{int(memory_usage)}".ljust(self.column_widths['mem_usage'])
        
        # Lista para armazenar as linhas formatadas
        formatted_lines = []
        
        # Quebra a mensagem em partes que cabem na coluna mensagem
        remaining_message = message
        while remaining_message:
            if len(remaining_message) <= self.column_widths['message']:
                # Se a mensagem inteira cabe na largura da coluna
                message_part = remaining_message.ljust(self.column_widths['message'])
                remaining_message = ""
            else:
                # Encontra um ponto de quebra (espaço) próximo ao limite
                split_pos = remaining_message[:self.column_widths['message']].rfind(' ')
                if split_pos == -1 or split_pos < self.column_widths['message'] * 0.8:
                    # Se não encontrar um espaço ou se estiver muito no início, corta no tamanho exato
                    split_pos = self.column_widths['message']
                
                message_part = remaining_message[:split_pos].ljust(self.column_widths['message'])
                remaining_message = remaining_message[split_pos:].lstrip()
            
            if not formatted_lines:
                # Primeira linha: inclui todas as colunas
                formatted_line = (
                    f"{timestamp_str} | {level_str} | {task_str} | {function_str} | "
                    f"{file_str} | {message_part} | {process_type_str} | {status_str} | "
                    f"{cpu_str} | {mem_str}"
                )
            else:
                # Linhas continuação: só inclui a mensagem, mantendo o alinhamento
                formatted_line = (
                    f"{' '.ljust(self.column_widths['timestamp'])} | "
                    f"{' '.ljust(self.column_widths['level'])} | "
                    f"{' '.ljust(self.column_widths['task'])} | "
                    f"{' '.ljust(self.column_widths['function'])} | "
                    f"{' '.ljust(self.column_widths['file'])} | "
                    f"{message_part} | "
                    f"{' '.ljust(self.column_widths['process_type'])} | "
                    f"{' '.ljust(self.column_widths['status'])} | "
                    f"{' '.ljust(self.column_widths['cpu_usage'])} | "
                    f"{' '.ljust(self.column_widths['mem_usage'])}"
                )
            
            formatted_lines.append(formatted_line)
        
        return "\n".join(formatted_lines)

    def _log_to_database(self, message, process_type, status, file=None, function=None):
        """Salva o log no banco de dados."""
        if not self.db_url:
            print(f"\033[93m[AVISO] URL do banco de dados não configurada. Log não será salvo no banco.\033[0m")
            return
            
        try:
            # Obter informações do sistema
            cpu_usage, memory_usage = self._get_system_info()
            
            # Verificar se a engine é válida ou criar uma nova
            if not self.engine:
                try:
                    self.engine = create_engine(self.db_url)
                    self._create_log_table_if_not_exists()
                except Exception as e:
                    print(f"\033[91m[ERRO] Falha ao criar conexão com o banco: {e}\033[0m")
                    return
                
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
            
            # Inserir no banco de dados
            query = f"""
            INSERT INTO public.{LOG_TABLE_NAME} 
            (timestamp, task, function, file, message, process_type, status, cpu_usage, memory_usage)
            VALUES
            (:timestamp, :task, :function, :file, :message, :process_type, :status, :cpu_usage, :memory_usage)
            """
            
            with self.engine.begin() as connection:
                connection.execute(text(query), log_data)
                
        except Exception as e:
            print(f"\033[91m[ERRO DB] Falha ao salvar log no banco: {e}\033[0m")
            # Imprime stack trace para facilitar diagnóstico
            print(f"\033[91m{traceback.format_exc()}\033[0m")
    

        """
        Formata a mensagem de log para o arquivo com colunas de largura fixa.
        
        Args:
            message (str): Mensagem de log
            file (str): Nome do arquivo
            function (str): Nome da função
            process_type (ProcessType): Tipo de processo
            status (LogStatus): Status do log
            cpu_usage (float): Uso de CPU
            memory_usage (float): Uso de memória
            
        Returns:
            str: Mensagem formatada
        """
        # Define as larguras das colunas
        timestamp_width = 19
        level_width = 8
        task_width = 8
        function_width = 25
        file_width = 15
        message_width = 60
        process_type_width = 12
        status_width = 12
        cpu_usage_width = 10
        mem_usage_width = 10
        
        # Gera o timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Trunca ou preenche os campos para garantir a largura fixa
        timestamp_str = timestamp.ljust(timestamp_width)
        level_str = status.value.upper().ljust(level_width)
        task_str = self.task_name.ljust(task_width)
        
        # Trunca os campos que podem ser muito longos
        function_str = function[:function_width].ljust(function_width)
        file_str = file[:file_width].ljust(file_width)
        process_type_str = process_type.value.ljust(process_type_width)
        status_str = status.value.ljust(status_width)
        
        # Formata os valores de CPU e memória
        cpu_str = f"{cpu_usage:.0f}".ljust(cpu_usage_width)
        mem_str = f"{memory_usage:.0f}".ljust(mem_usage_width)
        
        # Se a mensagem for mais longa que a largura máxima, quebra em múltiplas linhas
        messages = []
        remaining = message
        
        # Primeira linha com todos os campos
        first_part = remaining[:message_width]
        messages.append(
            f"{timestamp_str} | {level_str} | {task_str} | {function_str} | {file_str} | "
            f"{first_part.ljust(message_width)} | {process_type_str} | {status_str} | "
            f"{cpu_str} | {mem_str}"
        )
        
        # Linhas adicionais apenas com a mensagem
        if len(remaining) > message_width:
            remaining = remaining[message_width:]
            while remaining:
                next_part = remaining[:message_width]
                remaining = remaining[message_width:]
                
                # Cria uma linha continuação apenas com a mensagem, mantendo os alinhamentos
                continuation = (
                    f"{' '.ljust(timestamp_width)} | {' '.ljust(level_width)} | {' '.ljust(task_width)} | "
                    f"{' '.ljust(function_width)} | {' '.ljust(file_width)} | {next_part.ljust(message_width)} | "
                    f"{' '.ljust(process_type_width)} | {' '.ljust(status_width)} | "
                    f"{' '.ljust(cpu_usage_width)} | {' '.ljust(mem_usage_width)}"
                )
                messages.append(continuation)
        
        return "\n".join(messages)
    
    def log(self, message, process_type, status, file=None, function=None):
        """
        Registra uma mensagem de log no arquivo, banco de dados e console.
        
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
        
        # Exibir no console com cores
        self._print_console_message(message, process_type, status, file, function, cpu_usage, memory_usage)
        
        # Logar no arquivo
        self.file_logger.info(formatted_message)
        
        # Salvar no banco de dados
        self._log_to_database(message, process_type, status, file, function)
    
    def _print_console_message(self, message, process_type, status, file, function, cpu_usage, memory_usage):
        """
        Exibe uma mensagem formatada no console com cores.
        """
        # Timestamp para uso no console
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Definir as cores para diferentes status
        color_code = ""
        if status == LogStatus.ERROR or status == LogStatus.CRITICAL:
            color_code = "\033[91m"  # Vermelho
        elif status == LogStatus.WARNING:
            color_code = "\033[93m"  # Amarelo
        elif status == LogStatus.SUCCESS:
            color_code = "\033[92m"  # Verde
        elif status == LogStatus.INFO:
            color_code = "\033[94m"  # Azul
        elif status == LogStatus.DEBUG:
            color_code = "\033[90m"  # Cinza
        
        reset_code = "\033[0m"  # Resetar cor
        
        # Formatar mensagem do console
        console_message = f"[{timestamp}] [{status.value.upper()}] [{process_type.value}] {message} (CPU: {int(cpu_usage)}%, MEM: {int(memory_usage)}%)"
        
        # Exibir mensagem no console
        print(f"{color_code}{console_message}{reset_code}")
            

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


# Instância global do logger (inicializada como None)
logger = None

def initialize_logger(db_url=None):
    """Inicializa o logger global.
    
    Args:
        db_url (str, optional): URL de conexão com o banco de dados.
    """
    global logger
    logger = RPALogger(db_url)

def get_logger():
    """Retorna a instância global do logger.
    
    Raises:
        RuntimeError: Se o logger não foi inicializado.
    """
    if logger is None:
        raise RuntimeError("Logger não foi inicializado. Chame initialize_logger() primeiro.")
    return logger
