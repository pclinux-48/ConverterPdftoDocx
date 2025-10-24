from converter_word_perfeito import PDFToWordPerfeito
from pathlib import Path

def testar_conversao_perfeita():
    """Testa a conversão perfeita para Word"""
    
    pdf_file = "tabela pontuação.pdf"
    
    if not Path(pdf_file).exists():
        print(f"❌ Arquivo {pdf_file} não encontrado!")
        return
    
    print("🎯 CONVERSOR PDF → WORD PERFEITO")
    print("=" * 50)
    print("🔧 Focado em máxima fidelidade ao original")
    print("📊 Preservação de tabelas e formatação")
    print("📝 Layout e espaçamento mantidos")
    print("=" * 50)
    
    try:
        converter = PDFToWordPerfeito(pdf_file)
        
        print(f"📄 Processando: {pdf_file}")
        
        # Converter
        word_file = converter.convert_to_word()
        
        print(f"\n🎉 SUCESSO!")
        print(f"📄 Arquivo Word criado: {word_file}")
        print(f"💡 Abra o arquivo para verificar a fidelidade da conversão")
        
        # Verificar se o arquivo foi criado
        if Path(word_file).exists():
            size = Path(word_file).stat().st_size
            print(f"📊 Tamanho do arquivo: {size:,} bytes")
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    testar_conversao_perfeita()