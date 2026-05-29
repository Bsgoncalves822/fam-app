import pandas as pd
import re
import os
import sys

# --- PATH RESOLUTION ---
# Works on any machine regardless of username
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

# Try to find source files relative to common locations
def find_file(filename, search_dirs):
    for d in search_dirs:
        p = os.path.join(d, filename)
        if os.path.exists(p):
            return p
    return None

username = os.environ.get('USERNAME', os.environ.get('USER', ''))
downloads = os.path.join('C:\\Users', username, 'Downloads')

search_dirs = [
    os.path.join(downloads, 'FAM'),
    os.path.join(downloads, 'FAM', 'Duplicate handling'),
    downloads,
    BASE_DIR,
]

# --- FUNCIONARIOS: from FOLHA.xlsx ---
folha_path = find_file('FOLHA.xlsx', search_dirs)
if folha_path:
    print(f"Lendo FOLHA.xlsx: {folha_path}")
    folha_df = pd.read_excel(folha_path, sheet_name=0)
    name_col = next((c for c in folha_df.columns if any(k in c.lower() for k in ['nome', 'colaborador'])), folha_df.columns[0])
    folha_names = folha_df[name_col].dropna().unique()
    pd.DataFrame(folha_names, columns=['NOME']).to_csv(
        os.path.join(DATA_DIR, 'funcionarios.csv'), index=False, encoding='utf-8')
    print(f"[OK] Funcionarios: {len(folha_names)} registros")
else:
    print("[AVISO] FOLHA.xlsx nao encontrado - funcionarios.csv nao atualizado")

# --- FORNECEDORES: from DuplicateHandler.xlsx ---
def clean_name(name):
    if pd.isna(name): return ""
    return re.split(r'\s+\d+-', str(name))[0].strip()

dup_path = find_file('DuplicateHandler.xlsx', search_dirs)
if dup_path:
    print(f"Lendo DuplicateHandler.xlsx: {dup_path}")
    df = pd.read_excel(dup_path, sheet_name="FAM_ORIGIN")
    forn = df[df['Categoria'] == 'FORNECEDOR'][['CNPJ', 'Nome']].dropna()
    forn = forn.copy()
    forn['CNPJ'] = forn['CNPJ'].apply(lambda x: re.sub(r'[^\d]', '', str(x)))
    forn['NOME'] = forn['Nome'].apply(clean_name)
    forn.to_csv(os.path.join(DATA_DIR, 'fornecedores.csv'), index=False, columns=['CNPJ', 'NOME'], encoding='utf-8')
    print(f"[OK] Fornecedores: {len(forn)} registros")
else:
    print("[AVISO] DuplicateHandler.xlsx nao encontrado - fornecedores.csv nao atualizado")

print("\nDone. Execute python app.py para iniciar.")
