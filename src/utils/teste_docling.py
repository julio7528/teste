from docling.document_converter import DocumentConverter
import win32com.client

def convert_doc_to_docx(input_path, output_path):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(input_path)
    doc.SaveAs(output_path, FileFormat=16)  # 16 corresponde ao formato .docx
    doc.Close()
    word.Quit()

# Caminho do arquivo .doc de entrada
input_file = r"C:\RPA\MODELOS ESP-METODO\EME-MAME-ANEXO- ESPECIFICAÇÕES E METODOLOGIAS DE MATERIAL DE EMBALAGEM\EME.044 - 05 - EME.044.doc"
# Caminho do arquivo .docx de saída
output_file = input_file.replace(".doc", ".docx")

convert_doc_to_docx(input_file, output_file)
print(f"Arquivo convertido para: {output_file}")






# Caminho do documento de entrada
source = output_file

converter = DocumentConverter()

# Convertendo o documento
result = converter.convert(source)

# Exportando o conteúdo para markdown
markdown_content = result.document.export_to_markdown()

# Caminho do arquivo de saída
output_path = r"C:\RPA\RPA001_Garantia_De_Qualidade\data\saida_markdown.txt"

# Gravando o conteúdo em um arquivo de texto
with open(output_path, "a", encoding="utf-8") as file:
    file.write(markdown_content)

print(f"Conteúdo salvo em {output_path}")
