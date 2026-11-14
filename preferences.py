"""仅保存非敏感界面配置。"""
import json
import os
from pathlib import Path
import tempfile
PREF_PATH = Path.home() / '.workbuddy' / 'salary-consolidation' / 'gui_prefs_v2.json'
KEYS = {'base_dir','source','mapping','output','month','amount_field','mode'}

def load_preferences(path=None):
    try:
        data=json.loads(Path(path or PREF_PATH).read_text(encoding='utf-8'))
        return {k:v for k,v in data.items() if k in KEYS and isinstance(v,str)}
    except (OSError,ValueError,AttributeError):
        return {}

def save_preferences(values,path=None):
    path=Path(path or PREF_PATH)
    path.parent.mkdir(parents=True,exist_ok=True)
    fd,temp=tempfile.mkstemp(dir=path.parent,suffix='.json')
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as dest:
            json.dump({k:v for k,v in values.items() if k in KEYS and isinstance(v,str)},dest,ensure_ascii=False,indent=2)
        os.replace(temp,path)
    finally:
        Path(temp).unlink(missing_ok=True)
