#!/usr/bin/env python3
"""Opcional: gera data/pax.js para abrir o index.html com duplo clique (sem servidor).
O site publicado NÃO usa este arquivo — lê sempre data/pax.csv."""
import json
csv = open("data/pax.csv", encoding="utf-8").read()
cfg = json.load(open("data/config.json", encoding="utf-8"))
open("data/pax.js", "w", encoding="utf-8").write(
    "window.PAX_CSV=" + json.dumps(csv) + ";\nwindow.PAX_CONFIG=" + json.dumps(cfg, ensure_ascii=False) + ";\n")
print("data/pax.js atualizado")
