#!/usr/bin/env python3
"""Valida data/pax.csv antes de publicar. Erro = bloqueia a publicação; aviso = só informa."""
import csv, json, sys
from collections import defaultdict
from datetime import date

ATRACOES = {"Bondinho", "AquaRio", "Trem Corcovado", "Paineiras", "BioParque", "YupStar"}
OBRIGATORIAS = {"Bondinho", "AquaRio", "Trem Corcovado", "Paineiras"}   # BioParque/YupStar podem faltar
LIMITE_YOY = 0.35   # aviso se o mês novo variar mais de 35% vs. mesmo mês do ano anterior

erros, avisos = [], []
dados = defaultdict(dict)      # (ano, mes) -> {atracao: pax}
vistos = set()

with open("data/pax.csv", encoding="utf-8", newline="") as f:
    rd = csv.DictReader(f)
    if rd.fieldnames != ["data", "atracao", "pax_mil"]:
        sys.exit("ERRO: cabeçalho deve ser exatamente: data,atracao,pax_mil")
    for n, row in enumerate(rd, start=2):
        try:
            d = date.fromisoformat(row["data"])
        except ValueError:
            erros.append(f"linha {n}: data inválida '{row['data']}' (use AAAA-MM-01)"); continue
        if d.day != 1:
            erros.append(f"linha {n}: a data deve ser o dia 1 do mês ({row['data']})")
        a = row["atracao"].strip()
        if a not in ATRACOES:
            erros.append(f"linha {n}: atração desconhecida '{a}' (use: {', '.join(sorted(ATRACOES))})"); continue
        try:
            v = float(row["pax_mil"])
        except ValueError:
            erros.append(f"linha {n}: PAX não numérico '{row['pax_mil']}' (use ponto decimal, em mil)"); continue
        if v < 0:
            erros.append(f"linha {n}: PAX negativo ({a} {row['data']})")
        if v > 5000:
            avisos.append(f"linha {n}: {a} {row['data']} = {v} mil — confira se o valor está em MIL (ex.: 223.819 = 223,819)")
        if (d, a) in vistos:
            erros.append(f"linha {n}: {a} duplicado em {row['data']}")
        vistos.add((d, a)); dados[(d.year, d.month)][a] = v

for (y, m), r in sorted(dados.items()):
    falta = OBRIGATORIAS - set(r)
    if falta and y >= 2022:
        erros.append(f"{m:02d}/{y}: faltam {', '.join(sorted(falta))}")

ordem = sorted(dados)
if ordem:
    ya, ma = ordem[-1]                      # só o mês mais recente: compara com o mesmo mês do ano anterior
    ant = dados.get((ya - 1, ma), {})
    for a, z in dados[(ya, ma)].items():
        x = ant.get(a)
        if x and z and abs(z / x - 1) > LIMITE_YOY:
            avisos.append(f"{ma:02d}/{ya} {a}: {z / x - 1:+.0%} vs {ma:02d}/{ya - 1} — confira se não é erro de digitação")
        if x and z == 0:
            avisos.append(f"{ma:02d}/{ya} {a}: valor 0 (dado pendente? deixe a linha de fora em vez de lançar 0)")

try:
    cfg = json.load(open("data/config.json", encoding="utf-8"))
    ym = cfg.get("bioparque_zero_desde", "")
    if ym and (len(ym) != 7 or ym[4] != "-"):
        erros.append("config.json: bioparque_zero_desde deve ser AAAA-MM (ex.: 2026-01)")
except Exception as e:
    erros.append(f"config.json inválido: {e}")

ult = ordem[-1] if ordem else None
print(f"{len(vistos)} registros · {len(dados)} meses · último mês: {ult[1]:02d}/{ult[0]}" if ult else "sem dados")
for w in avisos: print("AVISO:", w)
for e in erros: print("ERRO: ", e)
sys.exit(1 if erros else 0)
