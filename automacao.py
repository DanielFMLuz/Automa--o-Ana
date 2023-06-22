import PyPDF2
import pandas as pd

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)
        text = []
        for page in reader.pages:
            text.append(page.extract_text())
        return text

def find_words_in_text(word_list, text):
    result = {}
    for word in word_list:
        word_pages = []
        for page, page_text in enumerate(text):
            if word.lower() in page_text.lower():
                word_pages.append(page + 1)  # Adding 1 to page index to match human-readable page numbers
        result[word] = word_pages
    return result

# Ler o documento PDF
pdf_path = r'C:\Users\User\Desktop\Automação Ana\Apostila_EAP_PRAÇA_VOLUME-III - PARTE I.pdf'
pdf_text = extract_text_from_pdf(pdf_path)

# Ler a planilha de palavras
planilha_path = r'C:\Users\User\Desktop\Automação Ana\Planilha sem título(1).xlsx'
df = pd.read_excel(planilha_path)
word_list = df.iloc[:, 0].dropna().tolist()

# Relacionar as palavras com as páginas do PDF
result = find_words_in_text(word_list, pdf_text)

# Imprimir o resultado
for word, pages in result.items():
    formatted_pages = [str(page + 2) for page in pages]
    pages_str = ', '.join(formatted_pages)
    print(f"'{word}': {pages_str}")


