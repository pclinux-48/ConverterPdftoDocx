from pdf2docx import Converter

# Caminho do arquivo PDF de entrada
pdf_file = input("Digite o caminho do arquivo PDF: ")
# Caminho do arquivo Word de saída
word_file = pdf_file+"convertido.docx"

# Criar conversor
cv = Converter(pdf_file)

# Converter todo o conteúdo do PDF para Word
cv.convert(word_file, start=0, end=None)

# Fechar conversor
cv.close()

print("Conversão concluída com sucesso!")
