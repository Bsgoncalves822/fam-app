# -*- coding: utf-8 -*-
"""
FAM App - Patch / Auto-updater
Run this once on any machine to pull latest files from GitHub.
Also used by app.py on startup for auto-update.
Public repo — no PAT needed.
"""

import os, json, shutil, urllib.request, urllib.error
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')

FILES_TO_UPDATE = [
    'app.py',
    'generate_data.py',
    'requirements.txt',
    'launch.vbs',
    'templates/index.html',
]

def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_config(cfg):
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

def fetch_file(repo, filepath):
    """Fetch raw file content from public GitHub repo."""
    url = f'https://raw.githubusercontent.com/{repo}/main/{filepath}'
    req = urllib.request.Request(url, headers={'User-Agent': 'FAM-Updater/1.0'})
    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.read()

def backup_file(path):
    if os.path.exists(path):
        shutil.copy2(path, path + '.bak')

def write_file(path, content_bytes):
    os.makedirs(os.path.dirname(path) if os.path.dirname(path) else '.', exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content_bytes.decode('utf-8'))

def run_update(verbose=True):
    cfg  = load_config()
    repo = cfg.get('github_repo', 'Bsgoncalves822/fam-app')

    if verbose:
        print(f'[UPDATE] Conectando ao repositorio: {repo}')

    updated = []
    failed  = []

    for filepath in FILES_TO_UPDATE:
        abs_path = os.path.join(BASE_DIR, filepath.replace('/', os.sep))
        try:
            content = fetch_file(repo, filepath)
            backup_file(abs_path)
            write_file(abs_path, content)
            updated.append(filepath)
            if verbose:
                print(f'  [OK] {filepath}')
        except urllib.error.HTTPError as e:
            failed.append(filepath)
            if verbose:
                print(f'  [ERRO] {filepath} — HTTP {e.code}')
        except Exception as e:
            failed.append(filepath)
            if verbose:
                print(f'  [ERRO] {filepath} — {e}')

    cfg['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cfg['last_update_files'] = updated
    save_config(cfg)

    if verbose:
        print(f'\n[UPDATE] {len(updated)} arquivos atualizados, {len(failed)} falhas.')
        if failed:
            print(f'[UPDATE] Falhas: {", ".join(failed)}')
        print('[UPDATE] Backups salvos como *.bak')

    return len(updated) > 0

if __name__ == '__main__':
    print('=' * 60)
    print(' FAM App — Atualizador')
    print('=' * 60)

    success = run_update(verbose=True)
    if success:
        print('\nAtualizacao concluida. Reinicie o FAM App.')
    else:
        print('\nNenhum arquivo atualizado. Verifique sua conexao.')

    input('\nPressione Enter para fechar...')
