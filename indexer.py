import os, json, threading, re, fitz
import database, config

class PDFIndexer:
    def __init__(self, db_manager, tree_file):
        self.db = db_manager
        self.tree_file = tree_file
        self.status = {'is_scanning': False, 'total_files': 0, 'done': False}
        self.idx_status = {'is_indexing': False, 'processed': 0, 'total_files': 0, 'current': '', 'done': False}
        self.stop_requested = False

    def scan(self, root_paths):
        self.status.update({'is_scanning': True, 'done': False, 'total_files': 0})
        flat_list = []
        
        for r_p in root_paths:
            r_p = os.path.normpath(r_p)
            if not os.path.exists(r_p): continue
            for root, _, files in os.walk(r_p):
                for f in files:
                    if f.lower().endswith('.pdf'):
                        full_p = os.path.join(root, f).replace('\\', '/')
                        flat_list.append({'n': f, 'p': full_p, 'base': r_p})
                        self.status['total_files'] = len(flat_list)

        with self.db.get_conn() as conn:
            for item in flat_list:
                conn.execute("INSERT OR IGNORE INTO files (path, filename, mod_time) VALUES (?,?,?)",
                             (item['p'], item['n'], os.path.getmtime(item['p'])))

        tree = []
        for item in flat_list:
            try:
                rel = os.path.relpath(item['p'], os.path.dirname(item['base'])).replace('\\', '/')
            except:
                rel = item['p']
                
            parts = rel.split('/')
            curr = {"children": tree}
            for part in parts[:-1]:
                found = next((child for child in curr['children'] if child.get('name') == part), None)
                if not found:
                    found = {"name": part, "type": "folder", "children": []}
                    curr['children'].append(found)
                curr = found
            curr['children'].append({"name": item['n'], "type": "file", "path": item['p']})

        with open(self.tree_file, 'w', encoding='utf-8') as f:
            json.dump({'tree': tree, 'flat': flat_list}, f, ensure_ascii=False)
        
        self.status.update({'is_scanning': False, 'done': True})

    def index_worker(self):
        if self.idx_status['is_indexing']: return
        self.stop_requested = False
        self.idx_status.update({'is_indexing': True, 'processed': 0, 'done': False})
        
        with self.db.get_conn() as conn:
            tasks = conn.execute("SELECT id, path, filename FROM files WHERE scanned=0").fetchall()
        
        self.idx_status['total_files'] = len(tasks)
        for t in tasks:
            if self.stop_requested: break
            self.idx_status['current'] = t['filename']
            try:
                with fitz.open(t['path']) as doc:
                    batch = []
                    for i, page in enumerate(doc):
                        if self.stop_requested: break
                        text = re.sub(r'\s+', ' ', page.get_text()).strip()
                        if text:
                            batch.append((t['id'], i+1, text))
                        
                        # שמירה כל 50 עמודים כדי לא לפוצץ את ה-RAM
                        if len(batch) >= 50:
                            with self.db.get_conn() as conn:
                                conn.executemany("INSERT INTO pages (file_id, page_num, content) VALUES (?,?,?)", batch)
                            batch = []
                    
                    if batch and not self.stop_requested:
                        with self.db.get_conn() as conn:
                            conn.executemany("INSERT INTO pages (file_id, page_num, content) VALUES (?,?,?)", batch)
                            conn.execute("UPDATE files SET scanned=1 WHERE id=?", (t['id'],))
            except Exception as e:
                print(f"Error indexing {t['path']}: {e}")
            
            self.idx_status['processed'] += 1
            
        self.idx_status.update({'is_indexing': False, 'done': True, 'current': ''})