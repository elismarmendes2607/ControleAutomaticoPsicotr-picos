# Controle de CICLO

Projeto para automação de atualização de documentos Word (`.docx`) e arquivos Excel (`.xls`) de controle de psicotrópicos.

---

## 📚 Funcionalidades

- 📝 Atualiza automaticamente o nome do médico no documento (`DR. Nome` em negrito)
- 📅 Atualiza a data da receita para o dia atual
- 🆔 Atualiza apenas o número do CRM no campo existente (`CRM: 11077`)
- 👤 Atualiza os nomes dos pacientes na tabela do Word
- 🚀 Abre o documento Word atualizado automaticamente após gerar

---

## ⚙️ Requisitos

- Python 3.10 ou superior
- Pacotes instalados:
  - `pandas`
  - `python-docx`
  - `xlrd`
  - `openpyxl`

> Instalar dependências:
> ```bash
> pip install -r requirements.txt
> ```

---

## 🛠 Como funciona

1. Ler o arquivo **`ciclo.XLS`** localizado em `C:\Users\SeuUsuario\Documents`.
2. Filtrar apenas as linhas onde:
   - `psv_cid` = `CONSULTA NO CONSULTO`
3. Atualizar no Word:
   - Nome do médico encontrado em `psv_apel` (adiciona "DR." e deixa em negrito)
   - Número do CRM encontrado em `fle_psv_cod`
   - Data para o dia atual
   - Nomes dos pacientes da coluna `pac_nome`
4. Salvar o arquivo Word atualizado como `CICLO_ATUALIZADO.docx`.
5. Abrir automaticamente o documento finalizado.

---

## 📂 Estrutura de Pastas

