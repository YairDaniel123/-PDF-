import os
import json
import webbrowser
import threading
import sys
import time
import logging

# === מניעת קריסה ב-EXE (הכי חשוב!) ===
# מבטל הדפסות כשהתוכנה רצה כ-EXE כדי למנוע קריסה
if getattr(sys, 'frozen', False):
    if sys.stdout is None: sys.stdout = open(os.devnull, 'w')
    if sys.stderr is None: sys.stderr = open(os.devnull, 'w')

try:
    from flask import Flask, render_template, jsonify, send_from_directory, request
except ImportError:
    os.system("py -m pip install flask")
    from flask import Flask, render_template, jsonify, send_from_directory, request

APP_TITLE = "מציג PDF מתקדם"

def get_base_path():
    if getattr(sys, 'frozen', False): return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'): return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

DATA_DIR = os.path.join(os.getenv('APPDATA'), "AdvancedPDF_Data")
if not os.path.exists(DATA_DIR):
    try: os.makedirs(DATA_DIR)
    except: DATA_DIR = get_base_path()

CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
INDEX_FILE = os.path.join(DATA_DIR, "file_index.json")

template_dir = get_resource_path('templates')
static_dir = get_resource_path('static')

app = Flask(__name__, template_folder=template_dir, static_folder=static_dir)
logging.getLogger('werkzeug').setLevel(logging.ERROR)

data_lock = threading.Lock()
scan_status = { 'is_scanning': False, 'total_files': 0, 'current_file': '', 'done': False }
cached_data = None
current_config = {'root_paths': [], 'favorites': [], 'last_read': '', 'dark_mode': False}
last_heartbeat = time.time() + 15 

# === לוגיקה ===

def load_config():
    global current_config
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                current_config.update(data)
        except: pass

def save_config():
    try:
        with data_lock:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(current_config, f, ensure_ascii=False)
    except: pass

def prune_empty_folders(node):
    if node['type'] == 'file': return True
    node['children'] = [child for child in node['children'] if prune_empty_folders(child)]
    node['children'].sort(key=lambda x: (1 if x['type'] == 'file' else 0, x['name'].lower()))
    return len(node['children']) > 0

def scan_worker():
    global scan_status, cached_data
    
    with data_lock:
        scan_status['is_scanning'] = True
        scan_status['done'] = False
        scan_status['total_files'] = 0
        scan_status['current_file'] = 'מתחיל...'
    
    try:
        file_tree = {'name': 'root', 'type': 'folder', 'path': '', 'children': []}
        flat_list = []
        count = 0
        
        # רשימת התעלמות - חובה כדי לא להיתקע
        SKIP_DIRS = {'Windows', 'Program Files', 'Program Files (x86)', 'System Volume Information', '$RECYCLE.BIN', 'AppData'}
        
        for root_folder in current_config['root_paths']:
            if not os.path.exists(root_folder): continue
            
            root_norm = os.path.normpath(root_folder)
            nodes_map = {'.': file_tree}

            for dirpath, dirnames, filenames in os.walk(root_norm):
                # סינון תיקיות מערכת
                dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith('.')]
                
                # נותן למערכת לנשום
                time.sleep(0.0001)

                pdf_files = [f for f in filenames if f.lower().endswith('.pdf')]
                if not pdf_files: continue

                try:
                    try:
                        rel_path = os.path.relpath(dirpath, root_norm)
                    except:
                        rel_path = dirpath.replace(':', '')

                    if rel_path == '.':
                        current_node = file_tree
                    else:
                        parts = rel_path.split(os.sep)
                        parent_path = '.'
                        parent_node = file_tree
                        
                        for part in parts:
                            curr_path = os.path.join(parent_path, part) if parent_path != '.' else part
                            
                            if curr_path not in nodes_map:
                                existing_child = next((c for c in parent_node['children'] if c['name'] == part and c['type'] == 'folder'), None)
                                if existing_child:
                                    nodes_map[curr_path] = existing_child
                                    parent_node = existing_child
                                else:
                                    new_folder = {'name': part, 'type': 'folder', 'children': []}
                                    parent_node['children'].append(new_folder)
                                    nodes_map[curr_path] = new_folder
                                    parent_node = new_folder
                            else:
                                parent_node = nodes_map[curr_path]
                        current_node = parent_node

                    for file in pdf_files:
                        full_path = os.path.join(dirpath, file).replace('\\', '/')
                        item = {'name': file, 'type': 'file', 'path': full_path}
                        current_node['children'].append(item)
                        flat_list.append({'n': file, 'p': full_path, 'f': os.path.basename(dirpath)})
                        
                        count += 1
                        if count % 20 == 0:
                            scan_status['total_files'] = count
                            scan_status['current_file'] = file
                except: continue

        scan_status['current_file'] = 'מסדר נתונים...'
        prune_empty_folders(file_tree)
        
        with data_lock:
            cached_data = {'tree': file_tree, 'flat': flat_list}
            
        try:
            with open(INDEX_FILE, 'w', encoding='utf-8') as f:
                json.dump(cached_data, f, ensure_ascii=False)
        except: pass

    except Exception: pass
    
    finally:
        with data_lock:
            scan_status['is_scanning'] = False
            scan_status['done'] = True
            scan_status['total_files'] = count if 'count' in locals() else 0

load_config()

# === נתיבים ===

@app.route('/')
def home(): return render_template('index.html', app_title=APP_TITLE)

@app.route('/favicon.ico')
def favicon(): return send_from_directory(static_dir, 'app_icon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/heartbeat')
def heartbeat():
    global last_heartbeat
    last_heartbeat = time.time()
    return "OK"

@app.route('/get_init_data')
def get_init_data():
    return jsonify({
        'paths': current_config.get('root_paths', []),
        'favorites': current_config.get('favorites', []),
        'last_read': current_config.get('last_read', ''),
        'dark_mode': current_config.get('dark_mode', False)
    })

@app.route('/save_preferences', methods=['POST'])
def save_preferences():
    data = request.json
    if 'dark_mode' in data: current_config['dark_mode'] = data['dark_mode']
    save_config()
    return jsonify({'status': 'saved'})

@app.route('/toggle_favorite', methods=['POST'])
def toggle_favorite():
    path = request.json.get('path')
    if path in current_config['favorites']: current_config['favorites'].remove(path)
    else: current_config['favorites'].append(path)
    save_config()
    return jsonify({'status': 'ok'})

@app.route('/set_last_read', methods=['POST'])
def set_last_read():
    current_config['last_read'] = request.json.get('path')
    save_config()
    return jsonify({'status': 'ok'})

@app.route('/browse_folder')
def browse_folder():
    # פונקציה יציבה יותר לבחירת תיקייה
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.lift()
        root.focus_force()
        
        path = filedialog.askdirectory(title="בחר ספרייה לסריקה")
        root.destroy()
        
        if path: return jsonify({'status': 'ok', 'path': path.replace('\\', '/')})
    except: pass
    return jsonify({'status': 'cancelled'})

@app.route('/update_paths', methods=['POST'])
def update_paths():
    # שומר ומאפס את המטמון כדי להכריח סריקה חדשה
    global cached_data
    paths = [p.replace('\\', '/') for p in request.json.get('paths', [])]
    current_config['root_paths'] = paths
    save_config()
    cached_data = None 
    return jsonify({'status': 'saved'})

@app.route('/start_scan')
def start_scan():
    global cached_data
    force = request.args.get('force') == 'true'
    
    if cached_data and not force: return jsonify({'status': 'ready'})
    if not scan_status['is_scanning']: 
        threading.Thread(target=scan_worker, daemon=True).start()
    return jsonify({'status': 'started'})

@app.route('/progress')
def progress(): return jsonify(scan_status)

@app.route('/get_result')
def get_result():
    global cached_data
    if cached_data: return jsonify(cached_data)
    if os.path.exists(INDEX_FILE):
        try:
            with open(INDEX_FILE, 'r', encoding='utf-8') as f:
                cached_data = json.load(f)
            return jsonify(cached_data)
        except: pass
    return jsonify({'tree': {'children': []}, 'flat': []})

@app.route('/open_pdf')
def open_pdf():
    path = request.args.get('path')
    if path and os.path.exists(path):
        return send_from_directory(os.path.dirname(path), os.path.basename(path))
    return "File not found", 404

@app.route('/shutdown')
def shutdown():
    func = request.environ.get('werkzeug.server.shutdown')
    if func: func()
    os._exit(0)
    return "BYE"

def monitor_browser():
    while True:
        time.sleep(2)
        if time.time() - last_heartbeat > 5: os._exit(0)

def open_browser():
    time.sleep(1)
    webbrowser.open_new('http://127.0.0.1:5000/')

if __name__ == '__main__':
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        threading.Thread(target=open_browser, daemon=True).start()
        threading.Thread(target=monitor_browser, daemon=True).start()
    app.run(port=5000, threaded=True)