import pdfplumber
import pandas as pd
import re

def extract_data_from_pdf(pdf_path):
    data = []
    cpf = ""
    nome = ""
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Processando página {page_num + 1}...")
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                
                for i, line in enumerate(lines):
                    cpf_nome_match = re.match(r"(\d{3}\.\d{3}\.\d{3}-\d{2})\s+(.+)", line)
                    if cpf_nome_match:
                        cpf = cpf_nome_match.group(1).strip()
                        nome = cpf_nome_match.group(2).strip()
                        print(f"Encontrado CPF: {cpf}, Nome: {nome}")
                    codigo_descricao = re.match(r"(\d+\.\d+) - [^-]+(?: - [^-]+)*\s+([\d\.\,]+)", line)
                    if codigo_descricao:
                        codigo = codigo_descricao.group(1).strip()
                        valor = codigo_descricao.group(2).strip()
                        data.append([cpf, nome, codigo, valor])
                        print(f"Adicionado: {cpf}, {nome}, {codigo}, {valor}")
                    plano_saude = re.match(r"(Plano Médico [^\-]+) - ([^\d]+) ([\d\.\,]+)", line)
                    plano_odontologico = re.match(r"(Plano Odontológico [^\-]+) - ([^\d]+) ([\d\.\,]+)", line)
                    
                    if plano_saude:
                        evento = "7.01"
                        pessoa = plano_saude.group(2).strip()
                        valor = plano_saude.group(3).strip()
                        data.append([cpf, nome, evento + " - " + pessoa, valor])
                        print(f"Adicionado: {cpf}, {nome}, {evento} - {pessoa}, {valor}")
                    
                    if plano_odontologico:
                        evento = "7.02"
                        pessoa = plano_odontologico.group(2).strip()
                        valor = plano_odontologico.group(3).strip()
                        data.append([cpf, nome, evento + " - " + pessoa, valor])
                        print(f"Adicionado: {cpf}, {nome}, {evento} - {pessoa}, {valor}")
    
    return data


pdf_path = r'C:\Users\pedro.magalhaes\Documents\DIRF Folha\INFORME WGNAII W152.pdf'


data = extract_data_from_pdf(pdf_path)


df = pd.DataFrame(data, columns=["CPF", "Nome", "Código", "Valor (R$)"])


csv_path = 'dados_extraidos.csv'
df.to_csv(csv_path, index=False)

print(f"Dados salvos em {csv_path}")

###SGL Niku###
