import os
import sys
import datetime
import pandas as pd
import argparse
from tabulate import tabulate

# Adicionar o diretório src ao path para poder importar módulos
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, '..')
sys.path.append(src_dir)

# Importar o logger
from utils.logger import get_logger, ProcessType, LogStatus

def parse_date(date_str):
    """Converte uma string de data para um objeto datetime."""
    try:
        return datetime.datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        print(f"Erro: formato de data inválido '{date_str}'. Use o formato YYYY-MM-DD.")
        sys.exit(1)

def main():
    """Função principal para visualizar os logs."""
    parser = argparse.ArgumentParser(description="Visualizador de logs do RPA001")
    
    # Adicionar argumentos
    parser.add_argument("-n", "--limit", type=int, default=50,
                        help="Número máximo de logs a mostrar (padrão: 50)")
    parser.add_argument("-s", "--status", type=str, choices=["information", "warning", "error", "critical", "debug", "success"],
                        help="Filtrar por status")
    parser.add_argument("-p", "--process-type", type=str, 
                        choices=["system", "business", "database", "file", "network", "selenium", "interface", "excel", "word"],
                        help="Filtrar por tipo de processo")
    parser.add_argument("-d", "--date", type=str,
                        help="Filtrar por data (formato YYYY-MM-DD)")
    parser.add_argument("-f", "--from-date", type=str,
                        help="Filtrar a partir da data (formato YYYY-MM-DD)")
    parser.add_argument("-t", "--to-date", type=str,
                        help="Filtrar até a data (formato YYYY-MM-DD)")
    parser.add_argument("-o", "--output", type=str,
                        help="Salvar a saída em um arquivo CSV")
    parser.add_argument("--html", action="store_true",
                        help="Gerar relatório em formato HTML")
    
    # Parsear argumentos
    args = parser.parse_args()
    
    # Obter o logger
    logger = get_logger()
    
    # Processar argumentos de data
    start_date = None
    end_date = None
    
    if args.date:
        date = parse_date(args.date)
        start_date = date
        end_date = date + datetime.timedelta(days=1)
    
    if args.from_date:
        start_date = parse_date(args.from_date)
    
    if args.to_date:
        end_date = parse_date(args.to_date)
        # Adiciona um dia para incluir logs do dia especificado
        end_date = end_date + datetime.timedelta(days=1)
    
    # Converter status e process_type para enums se fornecidos
    status = None
    if args.status:
        for s in LogStatus:
            if s.value == args.status:
                status = s
                break
    
    process_type = None
    if args.process_type:
        for p in ProcessType:
            if p.value == args.process_type:
                process_type = p
                break
    
    # Buscar logs
    print(f"Buscando logs (limite: {args.limit})...")
    logs_df = logger.get_logs(
        limit=args.limit,
        status=status,
        process_type=process_type,
        start_date=start_date,
        end_date=end_date
    )
    
    if logs_df.empty:
        print("Nenhum log encontrado com os filtros especificados.")
        return
    
    # Selecionar colunas relevantes e ordenar
    logs_df = logs_df[[
        'timestamp', 'task', 'function', 'file', 'message', 
        'process_type', 'status', 'cpu_usage', 'memory_usage'
    ]].sort_values('timestamp', ascending=False)
    
    # Formatar timestamp para melhor visualização
    logs_df['timestamp'] = logs_df['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # Truncar mensagens longas para melhor visualização
    logs_df['message_short'] = logs_df['message'].str.slice(0, 100).str.replace('\n', ' ')
    logs_df.loc[logs_df['message'].str.len() > 100, 'message_short'] += '...'
    
    # Selecionar colunas para exibição
    display_df = logs_df[[
        'timestamp', 'task', 'function', 'file', 'message_short', 
        'process_type', 'status', 'cpu_usage', 'memory_usage'
    ]]
    
    # Renomear colunas para exibição
    display_df.columns = [
        'Timestamp', 'Task', 'Function', 'File', 'Message', 
        'Process Type', 'Status', 'CPU %', 'Mem %'
    ]
    
    # Exibir resultado
    print(f"\nEncontrados {len(logs_df)} logs.")
    
    # Se o usuário pediu para salvar em CSV
    if args.output:
        logs_df.to_csv(args.output, index=False)
        print(f"Logs salvos em {args.output}")
    
    # Se o usuário pediu para gerar HTML
    if args.html:
        html_file = args.output.replace('.csv', '.html') if args.output else "rpa001_logs.html"
        
        # Estilizar o HTML para cores por status
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>RPA001 Logs</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                table { border-collapse: collapse; width: 100%; }
                th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }
                th { background-color: #f2f2f2; }
                tr:nth-child(even) { background-color: #f9f9f9; }
                .information { background-color: #e3f2fd; }
                .warning { background-color: #fff9c4; }
                .error { background-color: #ffcdd2; }
                .critical { background-color: #ffebee; }
                .debug { background-color: #f5f5f5; }
                .success { background-color: #e8f5e9; }
                .timestamp { white-space: nowrap; }
                .header { margin-bottom: 20px; }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>RPA001 Logs</h1>
                <p>Gerado em: {datetime_now}</p>
                <p>Total de logs: {total_logs}</p>
            </div>
            {table_html}
        </body>
        </html>
        """.format(
            datetime_now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            total_logs=len(logs_df),
            table_html=logs_df.to_html(
                classes='logs', 
                index=False,
                columns=['timestamp', 'task', 'function', 'file', 'message', 'process_type', 'status', 'cpu_usage', 'memory_usage'],
                table_id='logs-table'
            )
        )
        
        # Adicionar classes baseadas no status
        for status in ["information", "warning", "error", "critical", "debug", "success"]:
            html_content = html_content.replace(
                f'<td>{status}</td>', 
                f'<td class="{status}">{status}</td>'
            )
        
        # Adicionar classe para timestamp
        html_content = html_content.replace(
            '<th>timestamp</th>', 
            '<th class="timestamp">timestamp</th>'
        )
        
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Relatório HTML gerado em {html_file}")
    
    # Exibir logs na tela
    print("\n" + tabulate(display_df, headers='keys', tablefmt='psql'))

if __name__ == "__main__":
    main()