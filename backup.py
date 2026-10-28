# -*- coding: utf-8 -*-
"""Bazani xavfsiz nusxalash:  python backup.py   (backups/ papkasiga sanali fayl yozadi, oxirgi 14 tasi saqlanadi)"""
import os, sqlite3, time

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.environ.get("DB_PATH", os.path.join(BASE, "instance", "monitoring.db"))
OUT = os.environ.get("BACKUP_DIR", os.path.join(BASE, "backups"))
os.makedirs(OUT, exist_ok=True)
dst = os.path.join(OUT, "monitoring_%s.db" % time.strftime("%Y%m%d_%H%M%S"))
src = sqlite3.connect(DB)
dstc = sqlite3.connect(dst)
src.backup(dstc)  # ishlab turgan bazani ham buzmasdan nusxalaydi
dstc.close(); src.close()
files = sorted(f for f in os.listdir(OUT) if f.startswith("monitoring_") and f.endswith(".db"))
for old in files[:-14]:
    os.remove(os.path.join(OUT, old))
print("Nusxa saqlandi:", dst)
