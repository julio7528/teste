import pandas as pd
import os
import sys
from sqlalchemy import create_engine, text


current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, '..') 
sys.path.append(src_dir)

from config.config import load_config, get_database_url, get_caminho_de_para
from utils.read_files_utils import read_excel_file


load_config()

def insert_dataframe_to_postgres(df, table_name, db_url):
    try:
        # Cria uma engine do SQLAlchemy para se conectar ao banco
        engine = create_engine(db_url)

        # Insere o DataFrame na tabela do banco de dados
        df.to_sql(table_name, con=engine, if_exists='append', index=False)

        print(f"Dados inseridos com sucesso na tabela '{table_name}'.")
    except Exception as e:
        print(f"Erro ao tentar inserir dados no PostgreSQL: {e}")

# Exemplo de uso da função

def insert_with_query(query, db_url, values=None):
    """
    Insere dados no banco de dados PostgreSQL usando uma query SQL.

    Args:
        query (str): Query SQL para inserção.
        db_url (str): URL de conexão com o banco de dados.
        values (dict, optional): Valores a serem inseridos na query.

    Returns:
        None
    """
    try:
        # Cria a engine de conexão ao banco de dados
        engine = create_engine(db_url)

        # Conecta ao banco e executa a query
        with engine.connect() as connection:
            if values:
                connection.execute(text(query), values)
                connection.commit()
            else:
                connection.execute(text(query)) 
                connection.commit()

        
        
        print("Dados inseridos com sucesso!")
    except Exception as e:
        print(f"Erro ao inserir dados: {e}")



def update_log_data(nomearquivo, statusrevisao=None, statusenviadosesuite=None, statushomologado=None):
    """
    Atualiza os dados na tabela public.rpa001_controle_execucao com base no nome do arquivo.

    :param nomearquivo: Nome do arquivo para identificar o registro.
    :param statusrevisao: Novo valor para a coluna 'statusrevisao' (opcional).
    :param statusenviadosesuite: Novo valor para a coluna 'statusenviadosesuite' (opcional).
    :param statushomologado: Novo valor para a coluna 'statushomologado' (opcional).
    """
    db_url = get_database_url()

    # Monta os pares coluna = valor dinamicamente
    updates = []
    if statusrevisao is not None:
        updates.append(f"statusrevisao = '{statusrevisao}'")
    if statusenviadosesuite is not None:
        updates.append(f"statusenviadosesuite = '{statusenviadosesuite}'")
    if statushomologado is not None:
        updates.append(f"statushomologado = '{statushomologado}'")

    if not updates:
        print("Nenhuma coluna foi especificada para atualização.")
        return

    # Cria a query de update
    updates_query = ", ".join(updates)
    query = f"""
        UPDATE public.rpa001_controle_execucao
        SET {updates_query}
        WHERE nomearquivo = '{nomearquivo}';
    """

    # Executa a query
    insert_with_query(query, db_url)


def insert_log_data(nomearquivo, statusrevisao):
    db_url = get_database_url()
    query = f"""INSERT INTO public.rpa001_controle_execucao
        (nomearquivo, statusrevisao, statusenviadosesuite, statushomologado)
        VALUES('{nomearquivo}', '{statusrevisao}', '', '');
        """
    insert_with_query(query, db_url)


def query_to_dataframe(query, params=None):
    db_url = get_database_url()

    """
    Realiza uma consulta no banco de dados e retorna os resultados em um DataFrame.

    Args:
        query (str): Query SQL para consulta.
        db_url (str): URL de conexão com o banco de dados.
        params (dict, optional): Parâmetros para a query SQL.

    Returns:
        pd.DataFrame: DataFrame contendo os resultados da consulta.
    """
    try:
        # Cria a engine de conexão ao banco de dados
        engine = create_engine(db_url)

        # Conecta ao banco e realiza a consulta
        with engine.connect() as connection:
            if params:
                result = connection.execute(text(query), params)
                connection.commit()
            else:
                result = connection.execute(text(query))
                connection.commit()
            
            # Converte o resultado em DataFrame
            df = pd.DataFrame(result.fetchall(), columns=result.keys())
        
        return df

    except Exception as e:
        print(f"Erro ao realizar consulta: {e}")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro
