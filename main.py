import pandas as pd
from pathlib import Path
from datetime import datetime
from docx import Document
import re
import os

# 1. Caminhos
caminho_documents = Path.home() / "Documents"
caminho_excel = caminho_documents / "ciclo.XLS"
caminho_docx = caminho_documents / "CICLO.docx"

# 2. Função para limpar nomes de pacientes
def limpar_nome_paciente(nome):
    if isinstance(nome, str):
        partes = nome.split(',')
        if len(partes) > 1:
            return partes[1].strip()
        else:
            return nome.strip()
    return nome

# 3. Ler Excel
dados = pd.read_excel(caminho_excel, engine='xlrd')

# 4. Filtrar apenas consultas "CONSULTA NO CONSULTO"
dados_filtrado = dados[dados['psv_cid'] == "CONSULTA NO CONSULTO"]

# 5. Nome do médico
nome_medico_excel = dados_filtrado['psv_apel'].iloc[0].strip()
nome_medico_final = f"DR. {nome_medico_excel}"  # adiciona DR. antes do nome

# 6. CRM
crm_numero = str(dados_filtrado['fle_psv_cod'].iloc[0])  # só o número

# 7. Pacientes
pacientes = dados_filtrado['pac_nome'].dropna().apply(limpar_nome_paciente).tolist()

# 8. Data Atual
data_hoje = datetime.now().strftime('%d/%m/%Y')

# 9. Ler Word
doc = Document(caminho_docx)

# Atualizar Médico, Data e CRM nos parágrafos
for paragrafo in doc.paragraphs:
    if "DR. DENISE LUCAS VIANA" in paragrafo.text:
        paragrafo.clear()
        run = paragrafo.add_run(nome_medico_final)
        run.bold = True

    if "20/06/2023" in paragrafo.text:
        paragrafo.text = paragrafo.text.replace("20/06/2023", data_hoje)

    if "CRM:" in paragrafo.text:
        paragrafo.text = re.sub(r"(CRM:\s*)\d+", f"CRM: {crm_numero}", paragrafo.text)

# Atualizar Médico, Data e CRM nas tabelas também
for tabela in doc.tables:
    for linha in tabela.rows:
        for celula in linha.cells:
            texto = celula.text.strip()

            if "20/06/2023" in texto:
                celula.text = texto.replace("20/06/2023", data_hoje)

            if "DR. DENISE LUCAS VIANA" in texto:
                celula.clear()
                run = celula.paragraphs[0].add_run(nome_medico_final)
                run.bold = True

            if "CRM:" in texto:
                celula.text = re.sub(r"(CRM:\s*)\d+", f"CRM: {crm_numero}", texto)

# Atualizar nomes dos pacientes
if len(doc.tables) > 1:
    tabela_pacientes = doc.tables[1]
    total_celulas = len(tabela_pacientes._cells)

    for idx, nome in enumerate(pacientes):
        if idx < total_celulas:
            tabela_pacientes._cells[idx].text = nome
        else:
            print(f"⚠️ Mais pacientes do que células disponíveis! Paciente ignorado: {nome}")

# 10. Salvar novo Word
novo_docx = caminho_documents / "CICLO_ATUALIZADO.docx"
doc.save(novo_docx)
os.startfile(novo_docx) # Abre o arquivo após salvar
print(f"✅ Documento Word atualizado salvo em: {novo_docx}")
