#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hidroweb_precip.py
==================
Trata séries diárias de chuva do Hidroweb (ANA) e entrega, por estação, a
precipitação total mensal organizada em colunas (Jan..Dez), as médias mensais
e anual, e os insumos para o índice de erosão (EI) e o fator R.

Fluxo
-----
1. Lê um .txt com os códigos das estações (um por linha, ou separados por
   vírgula/espaço; '#' inicia comentário).
2. Obtém os dados diários de chuva:
     --fonte ws     baixa do WebService público da ANA (HidroSerieHistorica),
                    com retentativas e espera crescente;
     --fonte local  lê arquivos já baixados do portal Hidroweb (.zip, .csv,
                    .txt, .xlsx) em --pasta; zips são descompactados.
3. Escolhe, mês a mês, a melhor série (nível 2 = consistido; nível 1 = bruto
   só quando o consistido não tem leitura).
4. Monta SEMPRE o calendário completo (todos os anos x 12 meses) e marca o que
   falta. Assim, meses ausentes na fonte aparecem como falha em vez de sumirem.
5. Aplica critérios eliminatórios (existência de dados, anos de histórico,
   % de meses com falha). A janela de anos (ano inicial a final) e os limites
   dos dois últimos critérios são definidos pelo usuário: por opção na linha de
   comando ou, se omitidos, perguntados no terminal.
6. Grava uma planilha .xlsx por estação aprovada + resumo_geral.xlsx/.csv.
   Meses sem dado ficam em branco e listados na aba Falhas; nenhum valor é
   estimado e as médias usam apenas meses completos.

Uso
---
    python hidroweb_precip.py estacoes.txt --fonte ws --saida saida
    python hidroweb_precip.py estacoes.txt --fonte local --pasta baixados --saida saida
    python hidroweb_precip.py estacoes.txt --fonte local --pasta baixados \\
        --ano-inicial 1991 --ano-final 2020 --min-anos 20 --max-falhas-pct 10
"""
from __future__ import annotations

import argparse
import calendar
import io
import logging
import os
import re
import sys
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

import numpy as np
import pandas as pd

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
WS_URL = os.environ.get(
    "HIDROWEB_WS_URL", "https://telemetriaws1.ana.gov.br/ServiceANA.asmx/HidroSerieHistorica"
)
COLS_DIARIO = ["cod", "ano", "mes", "dia", "chuva", "status", "nivel", "total_arq"]
LOG = logging.getLogger("hidroweb")


# --------------------------------------------------------------------------- #
# Parâmetros
# --------------------------------------------------------------------------- #
@dataclass
class Criterios:
    min_anos: float = 0.0           # anos de histórico mínimos (1º ao último mês com dado)
    max_falhas_pct: float = 100.0   # % máxima de meses sem total mensal válido
    # (no programa, ambos vêm do usuário; 0 e 100 = "sem exigência" para uso como biblioteca)
    max_falhas_dia: int = 0         # dias sem leitura tolerados para validar um mês
    ano_inicio: int | None = None
    ano_fim: int | None = None
    so_consistido: bool = False     # True = ignora dados brutos (nível 1)
    ei_a: float | None = None       # EI = a * (p^2/P)^b  (coeficientes regionais)
    ei_b: float | None = None


@dataclass
class Resultado:
    cod: int
    situacao: str = "Sem dados"     # Aprovada | Reprovada | Sem dados | Erro
    motivos: list = field(default_factory=list)
    diario: pd.DataFrame | None = None
    mensal: pd.DataFrame | None = None   # índice (ano, mes)
    mat: pd.DataFrame | None = None      # anos x 1..12, totais observados
    janela: pd.DataFrame | None = None   # anos x 1..12, bool (dentro da janela)
    criterios: list = field(default_factory=list)
    info: dict = field(default_factory=dict)


def fmt(cod: int) -> str:
    return f"{cod:08d}"


def txt_janela_anos(crit: "Criterios") -> str:
    if crit.ano_inicio is None and crit.ano_fim is None:
        return "sem limite"
    return f"{crit.ano_inicio if crit.ano_inicio is not None else 'início'} a " \
           f"{crit.ano_fim if crit.ano_fim is not None else 'fim'}"


# --------------------------------------------------------------------------- #
# Entrada: lista de estações
# --------------------------------------------------------------------------- #
def ler_estacoes(caminho: Path) -> list[int]:
    texto = Path(caminho).read_text(encoding="utf-8-sig", errors="replace")
    codigos: list[int] = []
    for linha in texto.splitlines():
        linha = linha.split("#", 1)[0]
        for tok in re.split(r"[;,\s]+", linha.strip()):
            if not tok:
                continue
            if not tok.isdigit() or len(tok) > 8:
                LOG.warning("Código ignorado (não parece código ANA): %r", tok)
                continue
            cod = int(tok)
            if cod not in codigos:
                codigos.append(cod)
    return codigos


# --------------------------------------------------------------------------- #
# Conversões
# --------------------------------------------------------------------------- #
def _num(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip().str.replace(",", ".", regex=False)
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan, "NaN": np.nan})
    return pd.to_numeric(s, errors="coerce")


def _parse_data(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    iso = s.str.match(r"^\d{4}-\d{2}-\d{2}")
    d1 = pd.to_datetime(s.where(iso).str[:10], format="%Y-%m-%d", errors="coerce")
    d2 = pd.to_datetime(s.where(~iso).str[:10], format="%d/%m/%Y", errors="coerce")
    return d1.fillna(d2)


def wide_para_diario(w: pd.DataFrame | None) -> pd.DataFrame:
    """Formato do Hidroweb (uma linha por mês, colunas Chuva01..Chuva31) -> diário."""
    vazio = pd.DataFrame(columns=COLS_DIARIO)
    if w is None or w.empty:
        return vazio
    w = w.reset_index(drop=True).copy()
    w.columns = [str(c).strip() for c in w.columns]
    low = {c.lower(): c for c in w.columns}
    c_cod = low.get("estacaocodigo")
    c_data = low.get("data") or low.get("datahora")
    if not c_cod or not c_data:
        raise ValueError("colunas EstacaoCodigo/Data não encontradas")
    c_niv = low.get("nivelconsistencia")
    c_tot = low.get("total")

    datas = _parse_data(w[c_data])
    base = pd.DataFrame(
        {
            "cod": pd.to_numeric(w[c_cod].astype(str).str.strip(), errors="coerce"),
            "ano": datas.dt.year,
            "mes": datas.dt.month,
            "nivel": _num(w[c_niv]).fillna(1) if c_niv else 1,
        }
    )
    base = base[base.cod.notna() & base.ano.notna()].astype(
        {"cod": int, "ano": int, "mes": int, "nivel": int}
    )
    # total mensal consolidado que vem no arquivo (usado quando não há leituras diárias)
    base["total_arq"] = _num(w.loc[base.index, c_tot]) if c_tot else np.nan

    cols_val: dict[int, str] = {}
    cols_st: dict[int, str] = {}
    for c in w.columns:
        m = re.fullmatch(r"chuva(\d{2})(status)?", c.lower())
        if m and 1 <= int(m.group(1)) <= 31:
            (cols_st if m.group(2) else cols_val)[int(m.group(1))] = c
    partes = []
    for dia, c in sorted(cols_val.items()):
        p = base.copy()
        p["dia"] = dia
        p["chuva"] = _num(w.loc[p.index, c])
        p["status"] = _num(w.loc[p.index, cols_st[dia]]) if dia in cols_st else np.nan
        partes.append(p)
    if not partes:
        return vazio
    d = pd.concat(partes, ignore_index=True)
    ok = pd.to_datetime({"year": d.ano, "month": d.mes, "day": d.dia}, errors="coerce").notna()
    return d.loc[ok, COLS_DIARIO].reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Fonte 1: WebService público da ANA
# --------------------------------------------------------------------------- #
def baixar_ws(cod: int, pasta_bruto: Path, refazer: bool, tentativas: int = 6) -> tuple[bytes, bool]:
    """Retorna (conteúdo, veio_da_rede). Reaproveita XML já salvo, salvo --refazer."""
    destino = pasta_bruto / f"{fmt(cod)}.xml"
    if destino.exists() and not refazer:
        return destino.read_bytes(), False
    import requests  # só é necessário para --fonte ws

    params = {
        "codEstacao": str(cod),
        "dataInicio": "",
        "dataFim": "",
        "tipoDados": "2",          # 2 = chuva
        "nivelConsistencia": "",   # vazio = bruto e consistido
    }
    espera, ultimo = 1.0, ""
    for t in range(1, tentativas + 1):
        try:
            r = requests.get(WS_URL, params=params, timeout=120,
                             headers={"User-Agent": "hidroweb-precip/1.0"})
            if r.status_code == 200:
                destino.write_bytes(r.content)
                return r.content, True
            if r.status_code not in (429, 500, 502, 503, 504):
                raise RuntimeError(f"HTTP {r.status_code} ao baixar {fmt(cod)}")
            ultimo = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            ultimo = str(e)
        LOG.warning("Estação %s: tentativa %d/%d falhou (%s); aguardando %.0fs",
                    fmt(cod), t, tentativas, ultimo, min(60.0, espera))
        time.sleep(min(60.0, espera))
        espera *= 2
    raise RuntimeError(f"falha ao baixar {fmt(cod)} após {tentativas} tentativas ({ultimo})")


def xml_para_largo(conteudo: bytes, cod: int) -> pd.DataFrame:
    if b"<!DOCTYPE" in conteudo.upper() or b"<!ENTITY" in conteudo.upper():
        raise ValueError("XML recusado (contém DOCTYPE/ENTITY)")
    raiz = ET.fromstring(conteudo)
    linhas = []
    for el in raiz.iter():
        if el.tag.rsplit("}", 1)[-1] == "SerieHistorica":
            linhas.append({f.tag.rsplit("}", 1)[-1]: (f.text or "").strip() for f in el})
    if not linhas:
        return pd.DataFrame()
    w = pd.DataFrame(linhas)
    w["EstacaoCodigo"] = cod
    if "Data" not in w.columns and "DataHora" in w.columns:
        w["Data"] = w["DataHora"]
    return w


# --------------------------------------------------------------------------- #
# Fonte 2: arquivos baixados do portal (zip/csv/txt/xlsx)
# --------------------------------------------------------------------------- #
def extrair_zips(pasta: Path, destino: Path, limite_bytes: int = 500 * 1024 * 1024) -> None:
    fila = [(z, 0) for z in sorted(pasta.rglob("*.zip"))]
    while fila:
        z, prof = fila.pop(0)
        alvo = destino / f"{z.stem}_{abs(hash(str(z))) % 10**6}"
        alvo.mkdir(parents=True, exist_ok=True)
        try:
            with zipfile.ZipFile(z) as zf:
                for m in zf.infolist():
                    if m.is_dir():
                        continue
                    if m.file_size > limite_bytes:
                        LOG.warning("%s: item %s grande demais, ignorado", z.name, m.filename)
                        continue
                    nome = Path(m.filename).name        # achata: evita caminhos maliciosos
                    if not nome:
                        continue
                    saida = alvo / nome
                    saida.write_bytes(zf.read(m))
                    if saida.suffix.lower() == ".zip" and prof < 3:
                        fila.append((saida, prof + 1))
        except zipfile.BadZipFile:
            LOG.warning("Zip inválido ignorado: %s", z)


def _ler_texto(caminho: Path) -> str:
    bruto = caminho.read_bytes()
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return bruto.decode(enc)
        except UnicodeDecodeError:
            continue
    return bruto.decode("latin-1", errors="replace")


def ler_arquivo_local(caminho: Path) -> pd.DataFrame | None:
    suf = caminho.suffix.lower()
    if suf in (".xlsx", ".xls"):
        bruto = pd.read_excel(caminho, header=None, dtype=str)
        for i in range(min(60, len(bruto))):
            if bruto.iloc[i].astype(str).str.lower().str.contains("estacaocodigo").any():
                df = bruto.iloc[i + 1:].copy()
                df.columns = [str(c).strip() for c in bruto.iloc[i]]
                return df.reset_index(drop=True)
        return None
    linhas = _ler_texto(caminho).splitlines()
    idx = next((i for i, l in enumerate(linhas) if "estacaocodigo" in l.lower()), None)
    if idx is None:
        return None
    cab = linhas[idx]
    sep = max([";", "\t", ","], key=cab.count)
    return pd.read_csv(io.StringIO("\n".join(linhas[idx:])), sep=sep, dtype=str, index_col=False)


def carregar_local(pasta: Path, extraidos: Path, codigos: set[int]) -> dict[int, pd.DataFrame]:
    extrair_zips(pasta, extraidos)
    arquivos = []
    for base in (pasta, extraidos):
        arquivos += [p for p in sorted(base.rglob("*")) if p.is_file()]
    por_cod: dict[int, list] = {}
    for p in arquivos:
        suf = p.suffix.lower()
        if suf == ".mdb":
            LOG.warning("%s: .mdb não é lido (exporte como .csv no Hidroweb)", p.name)
            continue
        if suf not in (".csv", ".txt", ".xlsx", ".xls"):
            continue
        try:
            w = ler_arquivo_local(p)
            if w is None:
                continue
            d = wide_para_diario(w)
        except Exception as e:  # arquivo fora do padrão: avisa e segue
            LOG.warning("%s: não foi possível interpretar (%s)", p.name, e)
            continue
        for cod, g in d.groupby("cod"):
            if int(cod) in codigos:
                por_cod.setdefault(int(cod), []).append(g)
                LOG.info("Arquivo %s: estação %s (%d leituras diárias)", p.name, fmt(int(cod)), len(g))
    return {c: pd.concat(v, ignore_index=True) for c, v in por_cod.items()}


# --------------------------------------------------------------------------- #
# Núcleo: qualidade, calendário completo, totais mensais
# --------------------------------------------------------------------------- #
def escolher_serie(d: pd.DataFrame, so_consistido: bool) -> pd.DataFrame:
    if so_consistido:
        d = d[d.nivel == 2]
    if d.empty:
        return d
    nv = (d.assign(ok=d.chuva.notna(), tot=d.total_arq.notna())
            .groupby(["ano", "mes", "nivel"], as_index=False)
            .agg(nv=("ok", "sum"), tot=("tot", "max")))
    nv["tem"] = (nv.nv > 0) | nv.tot.astype(bool)
    nv = nv.sort_values(["ano", "mes", "tem", "nivel", "nv"],
                        ascending=[True, True, False, False, False])
    melhor = nv.drop_duplicates(["ano", "mes"])[["ano", "mes", "nivel"]]
    d = d.merge(melhor, on=["ano", "mes", "nivel"], how="inner")
    d = d.sort_values("chuva", na_position="last").drop_duplicates(["ano", "mes", "dia"], keep="first")
    return d.sort_values(["ano", "mes", "dia"]).reset_index(drop=True)


def processar_estacao(cod: int, diario: pd.DataFrame | None, crit: Criterios) -> Resultado:
    res = Resultado(cod)
    if diario is None or diario.empty:
        res.motivos.append("nenhuma leitura da estação na fonte")
        return res
    d = escolher_serie(diario, crit.so_consistido)
    if crit.ano_inicio is not None:
        d = d[d.ano >= crit.ano_inicio]
    if crit.ano_fim is not None:
        d = d[d.ano <= crit.ano_fim]
    if d.empty or not d.chuva.notna().any():
        res.motivos.append("sem leituras de chuva no período analisado")
        return res
    d = d.copy()
    d["ok"] = d.chuva.notna()
    d["acum"] = d.status.eq(4)
    d["duv"] = d.status.eq(3)
    res.diario = d[COLS_DIARIO]

    ag = d.groupby(["ano", "mes"]).agg(
        dias_validos=("ok", "sum"), soma=("chuva", "sum"),
        dias_acumulados=("acum", "sum"), dias_duvidosos=("duv", "sum"),
        tot_arq=("total_arq", "max"))
    com_dado = ag[(ag.dias_validos > 0) | ag.tot_arq.notna()]
    pos_ok = com_dado.index.get_level_values("ano") * 12 + com_dado.index.get_level_values("mes") - 1
    ini, fim = int(pos_ok.min()), int(pos_ok.max())

    idx = pd.MultiIndex.from_product([range(ini // 12, fim // 12 + 1), range(1, 13)],
                                     names=["ano", "mes"])
    m = ag.reindex(idx)
    presente = idx.isin(ag.index)
    tot_arq = m.tot_arq
    m = m.fillna(0)
    m["tot_arq"] = tot_arq
    m["dias_esperados"] = [calendar.monthrange(a, mm)[1] for a, mm in idx]
    pos = idx.get_level_values("ano") * 12 + idx.get_level_values("mes") - 1
    m["na_janela"] = np.asarray((pos >= ini) & (pos <= fim))
    dias_ok = (m.dias_validos >= m.dias_esperados - crit.max_falhas_dia) & (m.dias_validos > 0)
    usa_arq = (~dias_ok) & m.tot_arq.notna()     # nível 2 só com total mensal, sem diário
    completo = dias_ok | usa_arq
    m["total"] = np.where(dias_ok, m.soma, np.where(usa_arq, m.tot_arq, np.nan)).astype(float)
    m["do_arquivo"] = usa_arq
    motivos = []
    for pres, dv, esp, comp in zip(presente, m.dias_validos, m.dias_esperados, completo):
        if comp:
            motivos.append("")
        elif not pres:
            motivos.append("mês ausente na fonte")
        elif dv == 0:
            motivos.append("sem leituras no mês")
        else:
            motivos.append(f"{int(esp - dv)} dias sem leitura")
    m["motivo"] = motivos
    res.mensal = m
    res.mat = m["total"].unstack("mes")
    res.janela = m["na_janela"].unstack("mes").astype(bool)

    n_win = fim - ini + 1
    n_falhas = int((m.na_janela & m.total.isna()).sum())
    n_ok = n_win - n_falhas
    falhas_pct = 100.0 * n_falhas / n_win
    anos_hist = n_win / 12.0
    res.criterios = [
        ("Existência de dados (meses completos)", n_ok, "≥ 1", n_ok >= 1),
        ("Anos de histórico (1º ao último mês com dado)", round(anos_hist, 1),
         f"≥ {crit.min_anos:g}", anos_hist >= crit.min_anos),
        ("Meses com falha na janela (%)", round(falhas_pct, 1),
         f"≤ {crit.max_falhas_pct:g}", falhas_pct <= crit.max_falhas_pct),
    ]
    res.motivos = [f"{n}: {v} (exigido {lim})" for n, v, lim, ok in res.criterios if not ok]
    res.situacao = "Aprovada" if not res.motivos else "Reprovada"

    def mmaaaa(p: int) -> str:
        return f"{p % 12 + 1:02d}/{p // 12}"

    res.info = {
        "periodo": f"{mmaaaa(ini)} a {mmaaaa(fim)}",
        "anos_hist": round(anos_hist, 1),
        "meses_janela": n_win,
        "meses_validos": n_ok,
        "meses_falha": n_falhas,
        "falhas_pct": round(falhas_pct, 1),
        "dias_acumulados": int(m.dias_acumulados.sum()),
        "dias_duvidosos": int(m.dias_duvidosos.sum()),
        "meses_do_arquivo": int((m.do_arquivo & m.na_janela).sum()),
    }
    return res


def estatisticas(mat: pd.DataFrame) -> dict:
    med = mat.mean(axis=0, skipna=True)
    n = mat.count(axis=0)
    dp = mat.std(axis=0, ddof=1)
    anual = float(med.sum()) if int(n.min()) > 0 else np.nan
    completos = mat.dropna()
    return {
        "med": med, "n": n, "dp": dp, "anual": anual,
        "anual_xcheck": float(completos.sum(axis=1).mean()) if len(completos) else np.nan,
        "n_completos": len(completos),
    }


def indices_erosividade(est: dict, crit: Criterios) -> dict:
    anual = est["anual"]
    out = {"c": pd.Series(np.nan, index=est["med"].index), "ifm": np.nan, "ei": None, "R": np.nan}
    if not (anual and anual > 0):
        return out
    c = est["med"] ** 2 / anual            # coeficiente de Fournier mensal (p²/P)
    out["c"], out["ifm"] = c, float(c.sum())
    if crit.ei_a is not None and crit.ei_b is not None:
        ei = crit.ei_a * c ** crit.ei_b
        out["ei"], out["R"] = ei, float(ei.sum())
    return out


# --------------------------------------------------------------------------- #
# Saída
# --------------------------------------------------------------------------- #
def _w(ws, r, c, v, f=None):
    if v is None or (isinstance(v, (float, np.floating)) and np.isnan(v)):
        ws.write_blank(r, c, None, f)
    elif isinstance(v, (bool, np.bool_)):
        ws.write(r, c, "sim" if v else "não", f)
    elif isinstance(v, (int, float, np.integer, np.floating)):
        ws.write_number(r, c, float(v), f)
    else:
        ws.write(r, c, str(v), f)


def escrever_planilha(res: Resultado, crit: Criterios, caminho: Path, fonte: str) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(str(caminho))
    F = {
        "tit": wb.add_format({"bold": True, "font_size": 14}),
        "cab": wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1, "align": "center",
                              "valign": "vcenter", "text_wrap": True}),
        "txt": wb.add_format({"border": 1}),
        "int": wb.add_format({"border": 1, "num_format": "0"}),
        "n1": wb.add_format({"border": 1, "num_format": "0.0"}),
        "n2": wb.add_format({"border": 1, "num_format": "0.00"}),
        "bold": wb.add_format({"bold": True, "border": 1, "num_format": "0.0", "bg_color": "#F2F2F2"}),
        "boldt": wb.add_format({"bold": True, "border": 1, "bg_color": "#F2F2F2"}),
        "falha": wb.add_format({"border": 1, "bg_color": "#FFE699"}),
        "fora": wb.add_format({"border": 1, "bg_color": "#EDEDED"}),
        "ok": wb.add_format({"border": 1, "font_color": "#1E6B2E", "bold": True}),
        "ruim": wb.add_format({"border": 1, "font_color": "#B00020", "bold": True}),
    }
    est = estatisticas(res.mat)
    ind = indices_erosividade(est, crit)

    # ---------------- Resumo ----------------
    ws = wb.add_worksheet("Resumo")
    ws.set_column(0, 0, 46)
    ws.set_column(1, 6, 18)
    ws.write(0, 0, f"Estação {fmt(res.cod)} — precipitação (Hidroweb/ANA)", F["tit"])
    linhas = [
        ("Situação", res.situacao), ("Fonte dos dados", fonte), ("Janela analisada", res.info["periodo"]),
        ("Janela de anos solicitada", txt_janela_anos(crit)),
        ("Anos de histórico", res.info["anos_hist"]), ("Meses na janela", res.info["meses_janela"]),
        ("Meses com total válido", res.info["meses_validos"]),
        ("Meses com falha", res.info["meses_falha"]), ("% de meses com falha", res.info["falhas_pct"]),
        ("Dias com status 'acumulado' (4)", res.info["dias_acumulados"]),
        ("Dias com status 'duvidoso' (3)", res.info["dias_duvidosos"]),
        ("Tolerância de dias sem leitura por mês", crit.max_falhas_dia),
        ("Meses com total consolidado do nível 2 (sem leituras diárias)", res.info["meses_do_arquivo"]),
    ]
    r = 2
    for k, v in linhas:
        ws.write(r, 0, k, F["txt"])
        _w(ws, r, 1, v, F["txt"])
        r += 1
    r += 1
    ws.write_row(r, 0, ["Critério eliminatório", "Valor", "Limite", "Atende"], F["cab"])
    r += 1
    for nome, val, lim, ok in res.criterios:
        ws.write(r, 0, nome, F["txt"])
        _w(ws, r, 1, val, F["n1"])
        ws.write(r, 2, lim, F["txt"])
        ws.write(r, 3, "sim" if ok else "não", F["ok"] if ok else F["ruim"])
        r += 1
    r += 1
    cab = ["Mês", "Média observada (mm)", "N anos (observados)", "Desvio padrão (mm)"]
    cab.append("Coef. Fournier p²/P (mm)")
    if ind["ei"] is not None:
        cab.append("EI mensal (unid. do coef. informado)")
    ws.set_row(r, 45)
    ws.write_row(r, 0, cab, F["cab"])
    r += 1
    r0 = r
    for i, mes in enumerate(est["med"].index):
        c = 0
        ws.write(r, c, MESES[i], F["txt"]); c += 1
        _w(ws, r, c, est["med"][mes], F["n1"]); c += 1
        _w(ws, r, c, int(est["n"][mes]), F["int"]); c += 1
        _w(ws, r, c, est["dp"][mes], F["n1"]); c += 1
        _w(ws, r, c, ind["c"][mes], F["n1"]); c += 1
        if ind["ei"] is not None:
            _w(ws, r, c, ind["ei"][mes], F["n1"])
        r += 1
    c = 0
    ws.write(r, c, "ANUAL (soma das 12 médias)", F["boldt"]); c += 1
    _w(ws, r, c, est["anual"], F["bold"]); c += 1
    ws.write_blank(r, c, None, F["boldt"]); c += 1
    ws.write_blank(r, c, None, F["boldt"]); c += 1
    _w(ws, r, c, ind["ifm"], F["bold"]); c += 1
    if ind["ei"] is not None:
        _w(ws, r, c, ind["R"], F["bold"])
    r += 2
    ws.write(r, 0, f"Conferência: média dos totais anuais (anos com 12 meses; n={est['n_completos']})", F["txt"])
    _w(ws, r, 1, est["anual_xcheck"], F["n1"])
    r += 2
    notas = [
        "As médias usam só meses completos; meses com falha ficam de fora (nenhum valor é estimado).",
        "Coef. Fournier mensal = p²/P (p = média do mês; P = média anual). Coluna 'ANUAL' dessa coluna = Índice de Fournier modificado.",
    ]
    if ind["ei"] is not None:
        notas.append(f"EI = {crit.ei_a:g} × (p²/P)^{crit.ei_b:g}; R = soma dos 12 EI (coeficientes informados pelo usuário).")
    else:
        notas.append("EI/fator R não calculados: informe --ei-a e --ei-b (coeficientes da equação da sua região).")
    for t in notas:
        ws.write(r, 0, t); r += 1

    # ---------------- Mensal ----------------
    wm = wb.add_worksheet("Mensal")
    wm.set_column(0, 0, 18)
    wm.set_column(1, 12, 8)
    wm.set_column(13, 14, 14)
    wm.write_row(0, 0, ["Ano"] + MESES + ["Total anual (mm)", "Meses observados"], F["cab"])
    wm.freeze_panes(1, 1)
    for i, ano in enumerate(res.mat.index):
        rr = i + 1
        wm.write_number(rr, 0, int(ano), F["int"])
        for mes in range(1, 13):
            obs = res.mat.at[ano, mes]
            if pd.notna(obs):
                wm.write_number(rr, mes, float(obs), F["n1"])
            elif bool(res.janela.at[ano, mes]):
                wm.write_blank(rr, mes, None, F["falha"])
            else:
                wm.write_blank(rr, mes, None, F["fora"])
        linha = res.mat.loc[ano]
        _w(wm, rr, 13, float(linha.sum()) if linha.notna().all() else np.nan, F["n1"])
        wm.write_number(rr, 14, int(res.mat.loc[ano].notna().sum()), F["int"])
    rr = len(res.mat) + 2
    stats = [("Média (observado)", est["med"], "bold"), ("N anos (observado)", est["n"], "int"),
             ("Desvio padrão", est["dp"], "n1")]
    for nome, serie, f in stats:
        wm.write(rr, 0, nome, F["boldt"])
        for mes in range(1, 13):
            _w(wm, rr, mes, serie[mes], F[f])
        if nome == "Média (observado)":
            _w(wm, rr, 13, est["anual"], F["bold"])
        rr += 1
    rr += 1
    wm.write(rr, 0, "Legenda:")
    wm.write(rr, 1, "falha", F["falha"])
    wm.write(rr, 2, "fora da janela", F["fora"])

    # ---------------- Falhas ----------------
    wf = wb.add_worksheet("Falhas")
    wf.set_column(0, 1, 8)
    wf.set_column(2, 3, 14)
    wf.set_column(4, 4, 28)
    wf.write_row(0, 0, ["Ano", "Mês", "Dias com leitura", "Dias no mês", "Motivo"], F["cab"])
    falhas = res.mensal[res.mensal.na_janela & res.mensal.total.isna()]
    for i, ((ano, mes), row) in enumerate(falhas.iterrows(), start=1):
        wf.write_number(i, 0, int(ano), F["int"])
        wf.write(i, 1, MESES[mes - 1], F["txt"])
        wf.write_number(i, 2, int(row.dias_validos), F["int"])
        wf.write_number(i, 3, int(row.dias_esperados), F["int"])
        wf.write(i, 4, row.motivo, F["txt"])
    wb.close()


def linha_resumo(res: Resultado, crit: Criterios) -> dict:
    row = {"Código": fmt(res.cod), "Situação": res.situacao, "Motivos": "; ".join(res.motivos)}
    if res.mat is None:
        return row
    est = estatisticas(res.mat)
    ind = indices_erosividade(est, crit)
    row.update({
        "Período analisado": res.info["periodo"], "Anos de histórico": res.info["anos_hist"],
        "Meses com falha": res.info["meses_falha"], "% falhas": res.info["falhas_pct"],
    })
    for i, mes in enumerate(est["med"].index):
        row[f"P_{MESES[i]} (mm)"] = round(float(est["med"][mes]), 2) if pd.notna(est["med"][mes]) else np.nan
    row["P_anual (soma das médias, mm)"] = round(est["anual"], 2) if pd.notna(est["anual"]) else np.nan
    row["P_anual (média dos anos completos, mm)"] = (
        round(est["anual_xcheck"], 2) if pd.notna(est["anual_xcheck"]) else np.nan)
    row["IFM Fournier modificado (mm)"] = round(ind["ifm"], 2) if pd.notna(ind["ifm"]) else np.nan
    if ind["ei"] is not None:
        row["R (soma dos EI)"] = round(ind["R"], 2)
    return row


# --------------------------------------------------------------------------- #
# Programa principal
# --------------------------------------------------------------------------- #
def _perguntar(texto, converter, valido, ajuda, opcional=False, entrada=input):
    """Pergunta no terminal até receber um valor válido (Enter vazio = None, se opcional)."""
    while True:
        try:
            bruto = entrada(texto).strip().replace(",", ".")
        except EOFError:
            raise SystemExit("\nEntrada interrompida.")
        if not bruto:
            if opcional:
                return None
            print("  Resposta obrigatória.")
            continue
        try:
            v = converter(bruto)
        except ValueError:
            v = None
        if v is None or not valido(v):
            print(f"  Valor inválido. {ajuda}")
            continue
        return v


def completar_parametros(a, interativo: bool, entrada=input) -> list[str]:
    """Pergunta o que não veio nas opções (janela de anos, anos de histórico, % de falhas).

    Sem terminal interativo, a janela fica sem limite e os dois limites faltantes
    são devolvidos na lista de retorno (o chamador deve recusar a execução)."""
    faltando: list[str] = []
    if interativo and None in (a.ano_inicio, a.ano_fim, a.min_anos, a.max_falhas_pct):
        print("\nDefina os parâmetros da análise.")
    if interativo:
        ano_ok = lambda v: 1800 <= v <= 2100  # noqa: E731
        while True:
            if a.ano_inicio is None:
                a.ano_inicio = _perguntar("Ano inicial da janela (Enter = sem limite): ", int,
                                          ano_ok, "Digite um ano, por exemplo 1991.", True, entrada)
            if a.ano_fim is None:
                a.ano_fim = _perguntar("Ano final da janela (Enter = sem limite): ", int,
                                       ano_ok, "Digite um ano, por exemplo 2020.", True, entrada)
            if a.ano_inicio is not None and a.ano_fim is not None and a.ano_inicio > a.ano_fim:
                print("  O ano inicial é maior que o final. Informe a janela de novo.")
                a.ano_inicio = a.ano_fim = None
                continue
            break
    if a.min_anos is None:
        if interativo:
            a.min_anos = _perguntar("Mínimo de anos de histórico (1º ao último mês com dado), ex. 20: ",
                                    float, lambda v: 0 < v <= 200, "Digite um número maior que 0.",
                                    False, entrada)
        else:
            faltando.append("--min-anos")
    if a.max_falhas_pct is None:
        if interativo:
            a.max_falhas_pct = _perguntar("Máximo de meses com falha (%), ex. 10: ", float,
                                          lambda v: 0 <= v <= 100, "Digite um número entre 0 e 100.",
                                          False, entrada)
        else:
            faltando.append("--max-falhas-pct")
    return faltando


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("estacoes", help=".txt com os códigos das estações")
    ap.add_argument("--fonte", choices=["ws", "local"], default="ws")
    ap.add_argument("--pasta", help="pasta com .zip/.csv/.txt/.xlsx baixados do Hidroweb (--fonte local)")
    ap.add_argument("--saida", default="saida")
    ap.add_argument("--refazer", action="store_true", help="baixa de novo mesmo que o XML já esteja salvo")
    ap.add_argument("--pausa", type=float, default=3.0, help="segundos entre downloads (padrão 3)")
    ap.add_argument("--ano-inicial", "--ano-inicio", dest="ano_inicio", type=int,
                    help="primeiro ano da janela de análise (se omitido, é perguntado no terminal)")
    ap.add_argument("--ano-final", "--ano-fim", dest="ano_fim", type=int,
                    help="último ano da janela de análise (se omitido, é perguntado no terminal)")
    ap.add_argument("--min-anos", type=float,
                    help="mínimo de anos de histórico, do 1º ao último mês com dado (se omitido, é perguntado)")
    ap.add_argument("--max-falhas-pct", type=float,
                    help="máximo de meses com falha, em %% (se omitido, é perguntado)")
    ap.add_argument("--max-falhas-dia", type=int, default=0)
    ap.add_argument("--so-consistido", action="store_true")
    ap.add_argument("--ei-a", type=float, help="coeficiente a de EI = a*(p²/P)^b")
    ap.add_argument("--ei-b", type=float, help="expoente b de EI = a*(p²/P)^b")
    a = ap.parse_args(argv)

    if (a.ei_a is None) != (a.ei_b is None):
        ap.error("--ei-a e --ei-b devem ser informados juntos")
    if a.fonte == "local" and not a.pasta:
        ap.error("--fonte local exige --pasta")
    if a.min_anos is not None and a.min_anos <= 0:
        ap.error("--min-anos deve ser maior que 0")
    if a.max_falhas_pct is not None and not 0 <= a.max_falhas_pct <= 100:
        ap.error("--max-falhas-pct deve estar entre 0 e 100")
    if a.ano_inicio is not None and a.ano_fim is not None and a.ano_inicio > a.ano_fim:
        ap.error("--ano-inicial não pode ser maior que --ano-final")
    interativo = sys.stdin is not None and sys.stdin.isatty()
    faltando = completar_parametros(a, interativo)
    if faltando:
        ap.error("informe " + " e ".join(faltando) + " (ou rode em um terminal interativo para ser perguntado)")

    saida = Path(a.saida)
    for sub in ("estacoes", "dados_brutos", "dados_diarios", "extraidos"):
        (saida / sub).mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler(sys.stdout),
                  logging.FileHandler(saida / "log_execucao.txt", "w", encoding="utf-8")])

    crit = Criterios(min_anos=a.min_anos, max_falhas_pct=a.max_falhas_pct,
                     max_falhas_dia=a.max_falhas_dia, ano_inicio=a.ano_inicio, ano_fim=a.ano_fim,
                     so_consistido=a.so_consistido, ei_a=a.ei_a, ei_b=a.ei_b)
    codigos = ler_estacoes(Path(a.estacoes))
    if not codigos:
        LOG.error("Nenhum código válido em %s", a.estacoes)
        return 1
    LOG.info("%d estações na lista; fonte: %s", len(codigos), a.fonte)
    LOG.info("Parâmetros: janela de anos %s; mínimo de anos de histórico %g; máximo de meses com falha %g%%",
             txt_janela_anos(crit), crit.min_anos, crit.max_falhas_pct)

    # 1) obtenção dos dados diários
    diarios: dict[int, pd.DataFrame] = {}
    erros: dict[int, str] = {}
    if a.fonte == "local":
        diarios = carregar_local(Path(a.pasta), saida / "extraidos", set(codigos))
    else:
        for cod in codigos:
            try:
                conteudo, da_rede = baixar_ws(cod, saida / "dados_brutos", a.refazer)
                diarios[cod] = wide_para_diario(xml_para_largo(conteudo, cod))
                LOG.info("Estação %s: %d leituras diárias", fmt(cod), len(diarios[cod]))
                if da_rede:
                    time.sleep(a.pausa)
            except Exception as e:
                erros[cod] = str(e)
                LOG.error("Estação %s: %s", fmt(cod), e)

    # 2) qualidade + calendário completo
    resultados: dict[int, Resultado] = {}
    for cod in codigos:
        if cod in erros:
            resultados[cod] = Resultado(cod, "Erro", [erros[cod]])
            continue
        try:
            r = processar_estacao(cod, diarios.get(cod), crit)
        except Exception as e:
            LOG.exception("Estação %s: erro no processamento", fmt(cod))
            r = Resultado(cod, "Erro", [f"erro no processamento: {e}"])
        if r.situacao == "Sem dados" and a.fonte == "local" and cod not in diarios:
            r.motivos = ["estação não encontrada nos arquivos de --pasta"]
        resultados[cod] = r
        LOG.info("Estação %s: %s%s", fmt(cod), r.situacao, f" — {'; '.join(r.motivos)}" if r.motivos else "")

    com_mat = {c: r for c, r in resultados.items() if r.mat is not None}
    aprovadas = [c for c, r in com_mat.items() if r.situacao == "Aprovada"]

    # 3) saídas
    for c in aprovadas:
        escrever_planilha(resultados[c], crit, saida / "estacoes" / f"{fmt(c)}.xlsx",
                          "WebService ANA" if a.fonte == "ws" else f"arquivos locais (pasta '{Path(a.pasta).name}')")
    for c, r in com_mat.items():
        r.diario.to_csv(saida / "dados_diarios" / f"{fmt(c)}.csv", index=False)
    resumo = pd.DataFrame([linha_resumo(resultados[c], crit) for c in codigos])
    resumo.to_csv(saida / "resumo_geral.csv", index=False, sep=";", decimal=",", encoding="utf-8-sig")
    resumo.to_excel(saida / "resumo_geral.xlsx", index=False)
    cont = resumo["Situação"].value_counts().to_dict()
    LOG.info("Concluído: %s. Saída em %s", cont, saida.resolve())
    return 0


if __name__ == "__main__":
    sys.exit(main())
