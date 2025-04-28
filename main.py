import pandas as pd
from pathlib import Path
from datetime import datetime
from docx import Document
import re
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

# 1. Função para detectar caminho real no PyInstaller
def resource_path(relative_path):
    """Pega o caminho real do arquivo, dentro ou fora do PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).parent / relative_path

# 2. Função para abrir a janela de seleção do ciclo.XLS
def selecionar_arquivo_xls():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo ciclo.XLS",
        filetypes=[("Arquivos Excel", "*.xls")]
    )
    if not caminho_arquivo:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado. O programa será encerrado.")
        exit()
    return Path(caminho_arquivo)

# 3. Função para limpar nome dos pacientes
def limpar_nome_paciente(nome):
    if isinstance(nome, str):
        partes = nome.split(',')
        if len(partes) > 1:
            return partes[1].strip()
        else:
            return nome.strip()
    return nome

# 4. Definir caminhos
caminho_docx = resource_path("CICLO.docx")
caminho_excel = selecionar_arquivo_xls()

# 5. Ler Excel
dados = pd.read_excel(caminho_excel, engine='xlrd')

# 6. Filtrar apenas "CONSULTA NO CONSULTO"
dados_filtrado = dados[dados['psv_cid'] == "CONSULTA NO CONSULTO"]

# 7. Nome do médico
nome_medico_excel = dados_filtrado['psv_apel'].iloc[0].strip()
nome_medico_final = f"DR. {nome_medico_excel}"

# 8. CRM
crm_numero = str(dados_filtrado['fle_psv_cod'].iloc[0])

# 9. Lista de pacientes
pacientes = dados_filtrado['pac_nome'].dropna().apply(limpar_nome_paciente).tolist()

# 10. Data Atual
data_hoje = datetime.now().strftime('%d/%m/%Y')

# 11. Ler documento Word
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

# Atualizar Médico, Data e CRM nas tabelas
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

# 12. Salvar o novo documento no Desktop
desktop_path = Path.home() / "Desktop"
novo_docx = desktop_path / "CICLO_ATUALIZADO.docx"
doc.save(novo_docx)

print(f"✅ Documento Word atualizado salvo em: {novo_docx}")

# 13. Abrir o documento automaticamente
if novo_docx.exists():
    os.startfile(novo_docx)
else:
    print("❌ Erro: não foi possível encontrar o arquivo salvo.")
