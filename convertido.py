from pdf2docx import Converter

def pdf_para_word(pdf_path, word_path):
    # Criar conversor
    cv = Converter(pdf_path)
    
    # Converter todas as páginas para Word editável
    cv.convert(word_path, start=0, end=None)
    
    # Fechar conversor
    cv.close()
    print(f"✅ Arquivo convertido: {word_path}")

# Exemplo de uso
pdf_para_word("Todos os documentos.pdf", "saida3.docx")
