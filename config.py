import os, json

DATA_DIR = os.path.join(os.getenv('APPDATA', os.getcwd()), "AdvancedPDF_Data")
os.makedirs(DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(DATA_DIR, "library.db")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
TREE_FILE = os.path.join(DATA_DIR, "tree.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return {'root_paths': [], 'theme_mode': 0}

def save_config(c):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(c, f, ensure_ascii=False, indent=4)