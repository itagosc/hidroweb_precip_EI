#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Teste de ponta a ponta do hidroweb_precip.py com dados SINTÉTICOS (inventados).

Gera séries no formato do Hidroweb (csv latin-1 com vírgula decimal e linhas '//',
xlsx, zip dentro de zip, XML do WebService), roda o programa nos dois modos
(--fonte local e --fonte ws contra um servidor falso) e confere as saídas contra
valores calculados de forma independente.

    python testar_sintetico.py            # cria ./teste_sintetico
"""
import calendar
import http.server
import io
import os
import shutil
import subprocess
import sys
import threading
import urllib.parse
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import hidroweb_precip as hp  # noqa: E402

SCRIPT = AQUI / "hidroweb_precip.py"
RAIZ = AQUI / "teste_sintetico"
FALHAS: list[str] = []


def check(cond: bool, msg: str) -> None:
    print(("  OK   " if cond else "  FALHOU ") + msg)
    if not cond:
        FALHAS.append(msg)


# --------------------------------------------------------------------------- #
# Geração dos dados sintéticos
# --------------------------------------------------------------------------- #
BASE = np.array([250, 220, 180, 90, 60, 30, 25, 30, 70, 130, 190, 240.0])
A, B, C, D, E, F = 1944009, 2047001, 2047002, 2347003, 2447004, 2547005


def gerar():
    rng = np.random.default_rng(42)
    fator = {(a, m): rng.lognormal(0, 0.35) for a in range(1990, 2021) for m in range(1, 13)}

    def dias(total, nd):
        k = int(rng.integers(6, 20))
        v = np.zeros(nd)
        v[rng.choice(nd, size=k, replace=False)] = rng.dirichlet(np.ones(k)) * total
        return np.round(v, 1)

    todos = [(a, m) for a in range(1990, 2021) for m in range(1, 13)]
    e_ausentes = {todos[i] for i in rng.choice(len(todos), size=150, replace=False)}
    cfg = {
        A: dict(escala=1.00, anos=(1990, 2020)),
        B: dict(escala=0.90, anos=(1990, 2020),
                vazios={(1995, 3), (1995, 4), (2003, 12), (2008, 7)}, ausentes={(1998, 5), (2011, 1)},
                dias_falt={(2005, 9): 3, (2015, 2): 2}, acum={(2001, 6)}),
        C: dict(escala=1.10, anos=(1990, 2020), vazios={(1992, 6), (2001, 10), (2002, 11), (2016, 2), (2019, 8)}),
        D: dict(escala=1.00, anos=(1999, 2010)),
        E: dict(escala=0.95, anos=(1990, 2020), ausentes=e_ausentes),
    }
    registros, verdade, esperado = {}, {}, {}
    for cod, c in cfg.items():
        regs, ver, esp = [], {}, {}
        for a in range(c["anos"][0], c["anos"][1] + 1):
            for m in range(1, 13):
                nd = calendar.monthrange(a, m)[1]
                alvo = BASE[m - 1] * c["escala"] * fator[(a, m)] * rng.lognormal(0, 0.10)
                vals = dias(alvo, nd)
                ver[(a, m)] = float(vals.sum())
                status = np.where(np.isnan(vals), np.nan, 1.0)
                valido = True
                if (a, m) in c.get("ausentes", set()):
                    esp[(a, m)] = np.nan
                    continue
                if (a, m) in c.get("vazios", set()):
                    vals = np.full(nd, np.nan); status = np.full(nd, np.nan); valido = False
                elif (a, m) in c.get("dias_falt", {}):
                    n = c["dias_falt"][(a, m)]
                    vals = vals.copy(); vals[[3, 7, 11][:n]] = np.nan
                    status = np.where(np.isnan(vals), np.nan, 1.0)
                    valido = False
                if (a, m) in c.get("acum", set()):
                    status = status.copy(); status[9] = 4.0
                total = ver[(a, m)] if (not valido and cod == C and (a, m) in c.get("vazios", set())) else None
                esp[(a, m)] = float(np.nansum(vals)) if valido else (total if total is not None else np.nan)
                regs.append(dict(cod=cod, nivel=2, ano=a, mes=m, vals=vals, status=status, total=total))
        registros[cod], verdade[cod], esperado[cod] = regs, ver, esp

    # Estação A: duplicatas de nível 1 (bruto) em 2000 e mês consistido vazio em 06/1996
    extra = []
    for r in registros[A]:
        if r["ano"] == 2000:
            v = np.round(r["vals"] * 1.10, 1)
            extra.append(dict(r, nivel=1, vals=v, status=np.where(np.isnan(v), np.nan, 1.0)))
        if (r["ano"], r["mes"]) == (1996, 6):
            bruto = np.round(r["vals"] * 1.05, 1)
            extra.append(dict(r, nivel=1, vals=bruto, status=np.where(np.isnan(bruto), np.nan, 1.0)))
            esperado[A][(1996, 6)] = float(bruto.sum())
            r["vals"] = np.full(len(r["vals"]), np.nan); r["status"] = np.full(len(r["vals"]), np.nan)
    registros[A] += extra
    return registros, verdade, esperado


CAB = (["EstacaoCodigo", "NivelConsistencia", "Data", "Hora", "MediaDiaria", "Maxima", "Total", "TotalAnual"]
       + [f"Chuva{d:02d}" for d in range(1, 32)] + [f"Chuva{d:02d}Status" for d in range(1, 32)])


def csv_texto(regs) -> str:
    out = ["// Série histórica de chuvas - arquivo SINTÉTICO de teste", "// gerado por testar_sintetico.py", ";".join(CAB) + ";"]
    for r in regs:
        nd = len(r["vals"])
        v = [("" if np.isnan(x) else f"{x:.1f}".replace(".", ",")) for x in r["vals"]] + [""] * (31 - nd)
        s = [("" if np.isnan(x) else str(int(x))) for x in r["status"]] + [""] * (31 - nd)
        out.append(";".join([f"{r['cod']:08d}", str(r["nivel"]), f"01/{r['mes']:02d}/{r['ano']}", "00:00:00",
                             "", "", ("" if r.get("total") is None else f"{r['total']:.1f}".replace(".", ",")), "", *v, *s]) + ";")
    return "\n".join(out) + "\n"


def xlsx_escrever(regs, caminho: Path) -> None:
    linhas = [["// SINTÉTICO"], ["// planilha de teste"], CAB]
    for r in regs:
        nd = len(r["vals"])
        linhas.append([f"{r['cod']:08d}", r["nivel"], f"01/{r['mes']:02d}/{r['ano']}", "00:00:00", None, None, r.get("total"), None,
                       *list(r["vals"]), *[None] * (31 - nd), *list(r["status"]), *[None] * (31 - nd)])
    pd.DataFrame(linhas).to_excel(caminho, header=False, index=False)


def xml_texto(regs) -> str:
    out = ['<?xml version="1.0" encoding="utf-8"?>',
           '<DataSet xmlns="http://MRCS/"><diffgr:diffgram xmlns:msdata="urn:schemas-microsoft-com:xml-msdata" '
           'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"><DocumentElement xmlns="">']
    for i, r in enumerate(regs):
        nd = len(r["vals"])
        out.append(f'<SerieHistorica diffgr:id="SerieHistorica{i + 1}" msdata:rowOrder="{i}">'
                   f"<EstacaoCodigo>{r['cod']:08d}</EstacaoCodigo><NivelConsistencia>{r['nivel']}</NivelConsistencia>"
                   f"<DataHora>{r['ano']}-{r['mes']:02d}-01 00:00:00</DataHora><Total>{'' if r.get('total') is None else round(r['total'], 1)}</Total>")
        for d in range(nd):
            x, s = r["vals"][d], r["status"][d]
            out.append(f"<Chuva{d + 1:02d}>{'' if np.isnan(x) else x}</Chuva{d + 1:02d}>"
                       f"<Chuva{d + 1:02d}Status>{'' if np.isnan(s) else int(s)}</Chuva{d + 1:02d}Status>")
        out.append("</SerieHistorica>")
    out.append("</DocumentElement></diffgr:diffgram></DataSet>")
    return "".join(out)


XML_VAZIO = ('<?xml version="1.0" encoding="utf-8"?><DataSet xmlns="http://MRCS/"><diffgr:diffgram '
             'xmlns:diffgr="urn:schemas-microsoft-com:xml-diffgram-v1"/></DataSet>')


def itens_api(regs):
    """Itens no formato JSON da API HidroSerieChuva v1 (Chuva_01, Chuva_01_Status, ...)."""
    out = []
    for r in regs:
        nd = len(r["vals"])
        it = {}
        for d in range(1, 32):
            if d <= nd and not np.isnan(r["vals"][d - 1]):
                it[f"Chuva_{d:02d}"] = f"{r['vals'][d - 1]:.1f}"
                it[f"Chuva_{d:02d}_Status"] = "0" if np.isnan(r["status"][d - 1]) else str(int(r["status"][d - 1]))
            else:
                it[f"Chuva_{d:02d}"] = None
                it[f"Chuva_{d:02d}_Status"] = "0"
        it.update({"Data_Hora_Dado": f"{r['ano']}-{r['mes']:02d}-01 00:00:00.0",
                   "Nivel_Consistencia": str(r["nivel"]),
                   "Total": None if r.get("total") is None else f"{r['total']:.1f}",
                   "codigoestacao": str(r["cod"])})
        out.append(it)
    return out


def montar_jwt(exp: int) -> str:
    import base64, json
    b = lambda d: base64.urlsafe_b64encode(json.dumps(d).encode()).rstrip(b"=").decode()  # noqa: E731
    return f"{b({'alg': 'HS512'})}.{b({'exp': exp, 'iat': exp - 3600})}.assinatura-falsa"


# Janeiro/2020 da estação 1646001 exatamente como a API real devolveu (colado pelo usuário em 09/11/2026)
REAL_JAN2020 = [0.6, 0.0, 14.2, 29.8, 25.5, 2.0, 0.0, 7.2, 0.0, 0.0, 1.4, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.9,
                0.0, 0.0, 4.0, 70.0, 53.4, 21.7, 14.1, 49.5, 5.3, 10.3, 11.5, 3.8, 0.0]


def rodar(args: list[str], env=None, criterios: bool = True) -> subprocess.CompletedProcess:
    """Executa o programa sem terminal (stdin fechado). Os limites agora vêm do usuário:
    se o teste não os informa, usa 20 anos e 10 %."""
    if criterios and "--min-anos" not in args:
        args = args + ["--min-anos", "20", "--max-falhas-pct", "10"]
    e = dict(os.environ)
    e.update(env or {})
    return subprocess.run([sys.executable, str(SCRIPT), *args], capture_output=True, text=True, env=e,
                          stdin=subprocess.DEVNULL)


def esperado_df(esp: dict) -> pd.DataFrame:
    s = pd.Series(esp)
    s.index = pd.MultiIndex.from_tuples(s.index, names=["ano", "mes"])
    return s.unstack("mes")


def main() -> int:
    if RAIZ.exists():
        shutil.rmtree(RAIZ)
    (RAIZ / "baixados").mkdir(parents=True)
    registros, verdade, esperado = gerar()

    # ---- arquivos de entrada (formatos diferentes) ----
    pasta = RAIZ / "baixados"
    (pasta / "chuvas_A_B.csv").write_bytes(csv_texto(registros[A] + registros[B]).encode("latin-1"))
    xlsx_escrever(registros[C] + registros[D], pasta / "chuvas_C_D.xlsx")
    interno = io.BytesIO()
    with zipfile.ZipFile(interno, "w") as zi:
        zi.writestr("chuvas_E.csv", csv_texto(registros[E]))
    with zipfile.ZipFile(pasta / "E.zip", "w") as zf:
        zf.writestr("subpasta/E_interno.zip", interno.getvalue())
    (pasta / "lixo.zip").write_bytes(b"isto nao e um zip")
    (pasta / "leiame.txt").write_text("arquivo qualquer, sem séries\n", encoding="utf-8")
    est_txt = RAIZ / "estacoes.txt"
    est_txt.write_text("# estações de teste\n1944009, 2047001   # A e B\n02047002\n2347003 2447004\n"
                       "2547005\n2547005\nabc123\n", encoding="utf-8")

    print("\n[1] Lista de estações")
    cods = hp.ler_estacoes(est_txt)
    check(cods == [A, B, C, D, E, F], f"lista lida sem duplicatas/ruído: {cods}")

    print("\n[2] Execução --fonte local (csv latin-1, xlsx, zip dentro de zip, zip inválido)")
    r = rodar([str(est_txt), "--fonte", "local", "--pasta", str(pasta), "--saida", str(RAIZ / "out_local")])
    check(r.returncode == 0, f"código de saída 0 (stderr: {r.stderr[-300:]!r})")
    out = RAIZ / "out_local"
    res = pd.read_csv(out / "resumo_geral.csv", sep=";", decimal=",", encoding="utf-8-sig", dtype={"Código": str})
    sit = dict(zip(res["Código"], res["Situação"]))
    check(sit.get("01944009") == "Aprovada", "A aprovada")
    check(sit.get("02047001") == "Aprovada", "B aprovada")
    check(sit.get("02047002") == "Aprovada", "C aprovada")
    check(sit.get("02347003") == "Reprovada", "D reprovada (histórico curto)")
    check(sit.get("02447004") == "Reprovada", "E reprovada (muitas falhas)")
    check(sit.get("02547005") == "Sem dados", "F sem dados (não está nos arquivos)")
    mot = dict(zip(res["Código"], res["Motivos"].fillna("")))
    check("Anos de histórico" in mot["02347003"], "motivo de D cita anos de histórico")
    check("Meses com falha" in mot["02447004"], "motivo de E cita % de meses com falha")
    check("não encontrada" in mot["02547005"], "motivo de F: não encontrada nos arquivos")
    check(sorted(p.name for p in (out / "estacoes").glob("*.xlsx")) == ["01944009.xlsx", "02047001.xlsx", "02047002.xlsx"],
          "uma planilha por estação aprovada, nenhuma para reprovadas")

    print("\n[3] Médias mensais/anual contra cálculo independente")
    for cod in (A, B, C):
        ex = esperado_df(esperado[cod])
        med = ex.mean(axis=0)
        linha = res[res["Código"] == hp.fmt(cod)].iloc[0]
        ok_m = all(abs(linha[f"P_{hp.MESES[i]} (mm)"] - med[i + 1]) < 0.006 for i in range(12))
        check(ok_m, f"{hp.fmt(cod)}: 12 médias mensais conferem")
        check(abs(linha["P_anual (soma das médias, mm)"] - med.sum()) < 0.02, f"{hp.fmt(cod)}: média anual = soma das médias")
        compl = ex.dropna()
        check(abs(linha["P_anual (média dos anos completos, mm)"] - compl.sum(axis=1).mean()) < 0.02,
              f"{hp.fmt(cod)}: conferência pela média dos anos completos")

    print("\n[4] Planilha da estação A (nível 2 preferido; 06/1996 cai para o bruto)")
    wb = load_workbook(out / "estacoes" / "01944009.xlsx", data_only=True)
    ws = wb["Mensal"]
    ex = esperado_df(esperado[A])
    bad = 0
    for i, ano in enumerate(ex.index, start=2):
        for m in range(1, 13):
            v = ws.cell(row=i, column=m + 1).value
            if v is None or abs(v - ex.at[ano, m]) > 0.051:
                bad += 1
    check(bad == 0, f"372 células mensais de A iguais ao esperado (divergentes: {bad})")
    check(sorted(wb.sheetnames) == ["Falhas", "Mensal", "Resumo"], f"abas: {wb.sheetnames}")

    print("\n[5] Estação B: meses ausentes/vazios/incompletos (sem estimativa)")
    wbB = load_workbook(out / "estacoes" / "02047001.xlsx", data_only=True)
    wf = wbB["Falhas"]
    linhas = [[c.value for c in row] for row in wf.iter_rows(min_row=2)]
    linhas = [l for l in linhas if l[0] is not None]
    pares = {(l[0], hp.MESES.index(l[1]) + 1): l for l in linhas}
    check(set(pares) == {(1995, 3), (1995, 4), (2003, 12), (2008, 7), (1998, 5), (2011, 1), (2005, 9), (2015, 2)},
          f"8 falhas listadas corretamente ({len(pares)})")
    check(pares[(1998, 5)][4] == "mês ausente na fonte", "mês removido do arquivo = 'mês ausente na fonte'")
    check(pares[(1995, 3)][4] == "sem leituras no mês", "mês em branco = 'sem leituras no mês'")
    check("3 dias sem leitura" in pares[(2005, 9)][4], "mês com 3 dias faltando identificado")
    # mês com dia acumulado (status 4) não vira falha, só é contado
    check((2001, 6) not in pares, "mês com status 'acumulado' continua válido")
    check(wf.max_column == 5, "aba Falhas sem coluna de valor estimado")
    todos = [str(c.value) for s in wbB.worksheets for row in s.iter_rows() for c in row if c.value is not None]
    check(not any("preench" in t.lower() or "valor estimado" in t.lower() for t in todos),
          "planilha não menciona preenchimento nem valor estimado")
    linhaB = res[res["Código"] == "02047001"].iloc[0]
    check(int(linhaB["Meses com falha"]) == 8, "resumo informa 8 meses com falha")
    check("Meses preenchidos" not in res.columns, "resumo_geral sem a coluna 'Meses preenchidos'")
    ex_obs = esperado_df(esperado[B]).mean(axis=0)
    check(all(abs(linhaB[f"P_{hp.MESES[i]} (mm)"] - ex_obs[i + 1]) < 0.006 for i in range(12)),
          "médias de B só com meses completos (as 8 falhas ficam de fora)")
    r = rodar([str(est_txt), "--fonte", "local", "--pasta", str(pasta), "--saida", str(RAIZ / "out_x"),
               "--preencher", "nenhum"])
    check(r.returncode == 2, "a opção --preencher não existe mais")

    print("\n[6] Tolerância de dias faltantes e filtro de nível")
    dd = hp.carregar_local(pasta, RAIZ / "extr_tmp", {A, B, C})
    rc = hp.processar_estacao(C, dd[C], hp.Criterios())
    ok_c = all(abs(rc.mat.at[k[0], k[1]] - verdade[C][k]) < 0.06 for k in [(2019, 8), (1992, 6), (2001, 10)])
    check(ok_c and rc.info["meses_do_arquivo"] == 5,
          "nível 2 só com total (sem diário) usa o total do arquivo (5 meses)")
    crit = hp.Criterios(max_falhas_dia=3)
    rb = hp.processar_estacao(B, dd[B], crit)
    check(pd.notna(rb.mat.at[2005, 9]) and pd.notna(rb.mat.at[2015, 2]), "com --max-falhas-dia 3, meses incompletos passam a valer")
    ra = hp.processar_estacao(A, dd[A], hp.Criterios(so_consistido=True))
    ra2 = hp.processar_estacao(A, dd[A], hp.Criterios())
    check(pd.isna(ra.mat.at[1996, 6]) and pd.notna(ra2.mat.at[1996, 6]), "--so-consistido descarta 06/1996 (só havia bruto)")

    print("\n[7] XML do WebService (mesmo conteúdo) equivale ao CSV")
    xml_a = xml_texto(registros[A]).encode("utf-8")
    dx = hp.wide_para_diario(hp.xml_para_largo(xml_a, A))
    k = ["nivel", "ano", "mes", "dia"]
    d1 = dx.sort_values(k).reset_index(drop=True)
    d2 = dd[A].sort_values(k).reset_index(drop=True)
    check(len(d1) == len(d2) and np.allclose(d1.chuva, d2.chuva, equal_nan=True)
          and np.allclose(d1.status, d2.status, equal_nan=True), f"{len(d1)} leituras diárias idênticas")
    try:
        hp.xml_para_largo(b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><x>&a;</x>', 1)
        check(False, "XML com DOCTYPE/ENTITY deveria ser recusado")
    except ValueError:
        check(True, "XML com DOCTYPE/ENTITY é recusado")

    print("\n[8] --fonte ws contra servidor falso (com um 503 para testar retentativa) e cache")
    base_xml = {A: xml_texto(registros[A]), B: xml_texto(registros[B]), C: xml_texto(registros[C])}
    chamadas: dict[int, int] = {}

    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            cod = int(q["codEstacao"][0])
            chamadas[cod] = chamadas.get(cod, 0) + 1
            if cod == A and chamadas[cod] == 1:
                self.send_response(503); self.end_headers(); return
            corpo = base_xml.get(cod, XML_VAZIO).encode("utf-8")
            self.send_response(200); self.send_header("Content-Type", "text/xml"); self.end_headers()
            self.wfile.write(corpo)

        def log_message(self, *a):
            pass

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    url = f"http://127.0.0.1:{srv.server_address[1]}/HidroSerieHistorica"
    ws_txt = RAIZ / "estacoes_ws.txt"
    ws_txt.write_text("1944009\n2047001\n2047002\n2547005\n", encoding="utf-8")
    env = {"HIDROWEB_WS_URL": url, "NO_PROXY": "127.0.0.1,localhost", "no_proxy": "127.0.0.1,localhost"}
    r = rodar([str(ws_txt), "--fonte", "ws", "--saida", str(RAIZ / "out_ws"), "--pausa", "0"], env)
    check(r.returncode == 0, f"execução ws terminou (stderr: {r.stderr[-300:]!r})")
    check(chamadas.get(A) == 2, f"estação A foi rebaixada após o 503 ({chamadas.get(A)} chamadas)")
    rw = pd.read_csv(RAIZ / "out_ws" / "resumo_geral.csv", sep=";", decimal=",", encoding="utf-8-sig", dtype={"Código": str})
    sw = dict(zip(rw["Código"], rw["Situação"]))
    check(sw == {"01944009": "Aprovada", "02047001": "Aprovada", "02047002": "Aprovada", "02547005": "Sem dados"},
          f"situações via ws: {sw}")
    la = rw[rw["Código"] == "01944009"].iloc[0]
    lb = res[res["Código"] == "01944009"].iloc[0]
    check(all(abs(la[f"P_{m} (mm)"] - lb[f"P_{m} (mm)"]) < 1e-9 for m in hp.MESES), "médias de A via ws = via arquivos locais")
    srv.shutdown()
    n_antes = dict(chamadas)
    r = rodar([str(ws_txt), "--fonte", "ws", "--saida", str(RAIZ / "out_ws"), "--pausa", "0"],
              {"HIDROWEB_WS_URL": "http://127.0.0.1:9/x", "NO_PROXY": "127.0.0.1", "no_proxy": "127.0.0.1"})
    check(r.returncode == 0 and chamadas == n_antes, "segunda execução usa os XML salvos (sem rede)")

    print("\n[9] EI e fator R com coeficientes informados")
    r = rodar([str(est_txt), "--fonte", "local", "--pasta", str(pasta), "--saida", str(RAIZ / "out_ei"),
               "--ei-a", "67.355", "--ei-b", "0.85"])
    check(r.returncode == 0, "execução com EI terminou")
    re_ = pd.read_csv(RAIZ / "out_ei" / "resumo_geral.csv", sep=";", decimal=",", encoding="utf-8-sig", dtype={"Código": str})
    med = esperado_df(esperado[A]).mean(axis=0)
    c = med ** 2 / med.sum()
    R_esp = float((67.355 * c ** 0.85).sum())
    R_obt = float(re_[re_["Código"] == "01944009"]["R (soma dos EI)"].iloc[0])
    check(abs(R_obt - R_esp) < 0.05, f"R de A = {R_obt:.2f} (esperado {R_esp:.2f})")
    check(abs(float(re_[re_["Código"] == "01944009"]["IFM Fournier modificado (mm)"].iloc[0]) - c.sum()) < 0.02,
          "índice de Fournier modificado confere")

    print("\n[10] Janela de anos, limites definidos pelo usuário e critério por mês removido")
    r = rodar([str(est_txt), "--fonte", "local", "--pasta", str(pasta), "--saida", str(RAIZ / "out_win"),
               "--ano-inicial", "2000", "--ano-final", "2010", "--min-anos", "10", "--max-falhas-pct", "10"], criterios=False)
    check(r.returncode == 0, f"execução com janela 2000–2010 terminou (stderr: {r.stderr[-300:]!r})")
    rj = pd.read_csv(RAIZ / "out_win" / "resumo_geral.csv", sep=";", decimal=",", encoding="utf-8-sig",
                     dtype={"Código": str})
    sj = dict(zip(rj["Código"], rj["Situação"]))
    check(sj.get("01944009") == "Aprovada" and sj.get("02347003") == "Aprovada",
          "D (dados 1999–2010) passa a ser aprovada: só 2000–2010 conta, 11 anos ≥ 10")
    check(sj.get("02447004") == "Reprovada", "E continua reprovada por falhas dentro da janela")
    lj = rj[rj["Código"] == "01944009"].iloc[0]
    check(lj["Período analisado"] == "01/2000 a 12/2010" and abs(lj["Anos de histórico"] - 11.0) < 1e-9,
          f"A: período {lj['Período analisado']}, {lj['Anos de histórico']} anos")
    for cod in (A, D):
        med = esperado_df(esperado[cod]).loc[2000:2010].mean(axis=0)
        li = rj[rj["Código"] == hp.fmt(cod)].iloc[0]
        check(all(abs(li[f"P_{hp.MESES[i]} (mm)"] - med[i + 1]) < 0.006 for i in range(12)),
              f"{hp.fmt(cod)}: médias mensais calculadas só com 2000–2010")
    wbj = load_workbook(RAIZ / "out_win" / "estacoes" / "01944009.xlsx", data_only=True)
    textos = [str(c.value) for row in wbj["Resumo"].iter_rows() for c in row if c.value is not None]
    check("2000 a 2010" in textos, "planilha informa a janela de anos solicitada")
    check(not any("Mín. de anos válidos" in t for t in textos), "critério por mês do calendário não aparece mais")
    check(sum(1 for t in textos if t in ("Existência de dados (meses completos)",
                                          "Anos de histórico (1º ao último mês com dado)",
                                          "Meses com falha na janela (%)")) == 3, "restam os 3 critérios")
    check("Mín. anos por mês" not in rj.columns, "resumo_geral sem a coluna do critério removido")

    r = rodar([str(est_txt), "--fonte", "local", "--pasta", str(pasta), "--saida", str(RAIZ / "out_x"),
               "--ano-inicial", "2010", "--ano-final", "2000"])
    check(r.returncode == 2 and "ano-inicial" in r.stderr, "janela invertida é recusada")
    r = rodar([str(est_txt), "--fonte", "local", "--pasta", str(pasta), "--saida", str(RAIZ / "out_x")],
              criterios=False)
    check(r.returncode == 2 and "--min-anos" in r.stderr and "--max-falhas-pct" in r.stderr,
          "sem terminal e sem limites informados: recusa e diz o que falta")

    # perguntas interativas (entrada simulada, com respostas inválidas no meio)
    import argparse as _ap
    respostas = iter(["abc", "1991", "1990", "1991", "2020", "", "-3", "20,5", "150", "10"])
    a = _ap.Namespace(ano_inicio=None, ano_fim=None, min_anos=None, max_falhas_pct=None)
    falt = hp.completar_parametros(a, True, entrada=lambda _t: next(respostas))
    check((a.ano_inicio, a.ano_fim, a.min_anos, a.max_falhas_pct) == (1991, 2020, 20.5, 10.0) and falt == [],
          "perguntas: rejeita texto, janela invertida, vazio, negativo e >100; aceita vírgula decimal")
    a = _ap.Namespace(ano_inicio=None, ano_fim=None, min_anos=None, max_falhas_pct=None)
    respostas = iter(["", "", "15", "5"])
    hp.completar_parametros(a, True, entrada=lambda _t: next(respostas))
    check(a.ano_inicio is None and a.ano_fim is None and a.min_anos == 15 and a.max_falhas_pct == 5,
          "Enter em branco nos anos = sem limite")
    a = _ap.Namespace(ano_inicio=1991, ano_fim=2020, min_anos=20.0, max_falhas_pct=10.0)

    def _nao_deve_perguntar(_t):
        raise AssertionError("não deveria perguntar")
    check(hp.completar_parametros(a, True, entrada=_nao_deve_perguntar) == [],
          "tudo informado por opção: não pergunta nada")

    print("\n[11] --fonte api: formato real, limite de 366 dias, token (servidor simulado)")
    import json as _json
    import datetime as _dt
    import time as _time
    # (a) item real da API -> leituras diárias
    item = {f"Chuva_{d:02d}{suf}": (f"{REAL_JAN2020[d - 1]}" if not suf else "1")
            for d in range(1, 32) for suf in ("", "_Status")}
    item.update({"Data_Hora_Dado": "2020-01-01 00:00:00.0", "Data_Ultima_Alteracao": "2020-07-24 00:00:00.0",
                 "Dia_Maxima": "22", "Maxima": "70.0", "Nivel_Consistencia": "1", "Numero_Dias_de_Chuva": "18",
                 "Tipo_Medicao_Chuvas": "1", "Total": "325.2", "Total_Anual": None, "Total_Anual_Status": "0",
                 "Total_Status": "1", "codigoestacao": "1646001"})
    dj = hp.wide_para_diario(hp.itens_para_largo([item], 1646001))
    check(len(dj) == 31 and abs(dj.chuva.sum() - 325.2) < 1e-6 and dj.nivel.eq(1).all()
          and dj.total_arq.eq(325.2).all() and (dj.ano == 2020).all() and (dj.mes == 1).all(),
          "item real da API (jan/2020, estação 1646001): 31 dias somam o Total 325,2")
    # (b) token
    JWT_OK = montar_jwt(int(_time.time()) + 3600)
    exp = hp.expiracao_jwt(JWT_OK)
    check(exp is not None and abs(exp.timestamp() - (_time.time() + 3600)) < 5, "expiração lida do JWT")
    check(hp.expiracao_jwt("isto-nao-e-jwt") is None, "token ilegível não quebra a leitura da expiração")
    # (c) servidor simulado
    dados_api = {A: itens_api(registros[A]), B: itens_api(registros[B]), C: itens_api(registros[C])}
    pedidos: list = []
    chamadas_api = [0]          # todas as chamadas, inclusive as recusadas com 401

    class HA(http.server.BaseHTTPRequestHandler):
        def _json(self, corpo):
            b = _json.dumps(corpo).encode()
            self.send_response(corpo["code"]); self.send_header("Content-Type", "application/json")
            self.end_headers(); self.wfile.write(b)

        def do_GET(self):
            chamadas_api[0] += 1
            if self.headers.get("Authorization") != f"Bearer {JWT_OK}":
                return self._json({"status": "UNAUTHORIZED", "code": 401, "message": "Token inválido", "items": None})
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            cod = int(q["Código da Estação"][0])
            ini = _dt.date.fromisoformat(q["Data Inicial (yyyy-MM-dd)"][0])
            fim = _dt.date.fromisoformat(q["Data Final (yyyy-MM-dd)"][0])
            assert q["Tipo Filtro Data"] == ["DATA_LEITURA"]
            pedidos.append((cod, ini, fim))
            if (fim - ini).days > 366:
                return self._json({"status": "NOT_ACCEPTABLE", "code": 406, "items": None,
                                   "message": "Intervalo entre datas superior ao limite permitido de 366 dias. Verifique!"})
            itens = [it for it in dados_api.get(cod, [])
                     if ini <= _dt.date.fromisoformat(it["Data_Hora_Dado"][:10]) <= fim]
            self._json({"status": "OK", "code": 200, "message": "Sucesso", "items": itens or None})

        def log_message(self, *a):
            pass

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), HA)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    envA = {"HIDROWEB_API_URL": f"http://127.0.0.1:{srv.server_address[1]}/HidroSerieChuva/v1",
            "HIDROWEB_TOKEN": JWT_OK, "NO_PROXY": "127.0.0.1,localhost", "no_proxy": "127.0.0.1,localhost"}
    base = ["--fonte", "api", "--ano-inicial", "2000", "--ano-final", "2010", "--min-anos", "10",
            "--max-falhas-pct", "10", "--pausa", "0"]
    r = rodar([str(ws_txt), *base, "--saida", str(RAIZ / "out_api")], envA, criterios=False)
    check(r.returncode == 0, f"execução api terminou (stdout: {r.stdout[-300:]!r})")
    check(len(pedidos) == 44, f"44 requisições = 4 estações x 11 anos ({len(pedidos)})")
    check(all((f - i).days <= 366 for _, i, f in pedidos) and pedidos[0][1:] == (_dt.date(2000, 1, 1), _dt.date(2000, 12, 31)),
          "cada requisição cobre 1 ano (dentro do limite de 366 dias)")
    ra = pd.read_csv(RAIZ / "out_api" / "resumo_geral.csv", sep=";", decimal=",", encoding="utf-8-sig",
                     dtype={"Código": str})
    sa = dict(zip(ra["Código"], ra["Situação"]))
    check(sa == {"01944009": "Aprovada", "02047001": "Aprovada", "02047002": "Aprovada", "02547005": "Sem dados"},
          f"situações via api: {sa}")
    for cod in (A, B, C):
        la, lw = ra[ra["Código"] == hp.fmt(cod)].iloc[0], rj[rj["Código"] == hp.fmt(cod)].iloc[0]
        check(all(abs(la[f"P_{m} (mm)"] - lw[f"P_{m} (mm)"]) < 1e-9 for m in hp.MESES)
              and la["Meses com falha"] == lw["Meses com falha"],
              f"{hp.fmt(cod)}: médias e falhas via api = via arquivos locais (mesma janela)")
    todos_arqs = [p for p in (RAIZ / "out_api").rglob("*") if p.is_file() and p.suffix in (".txt", ".json", ".csv")]
    check(not any(JWT_OK in p.read_text(encoding="utf-8", errors="ignore") for p in todos_arqs),
          "o token não aparece em nenhum arquivo de saída nem no log")
    n = len(pedidos)
    envMorto = dict(envA, HIDROWEB_API_URL="http://127.0.0.1:9/x")
    r = rodar([str(ws_txt), *base, "--saida", str(RAIZ / "out_api")], envMorto, criterios=False)
    check(r.returncode == 0 and len(pedidos) == n, "segunda execução reaproveita o JSON salvo (sem rede)")
    # (d) falhas de token: interrompem a execução com mensagem clara
    n = len(pedidos)
    r = rodar([str(ws_txt), *base, "--saida", str(RAIZ / "out_api_t1")], dict(envA, HIDROWEB_TOKEN=montar_jwt(int(_time.time()) + 3600) + "x"),
              criterios=False)
    check(r.returncode == 1 and "recusou o token" in r.stdout and "Traceback" not in r.stderr,
          "token recusado (401): interrompe com mensagem clara")
    antes = chamadas_api[0]
    r = rodar([str(ws_txt), *base, "--saida", str(RAIZ / "out_api_t2")], dict(envA, HIDROWEB_TOKEN=montar_jwt(int(_time.time()) - 60)),
              criterios=False)
    check(r.returncode == 1 and "expirado" in r.stdout and chamadas_api[0] == antes,
          "token expirado: recusado antes de chamar a API")
    r = rodar([str(ws_txt), *base, "--saida", str(RAIZ / "out_api_t3")], dict(envA, HIDROWEB_TOKEN=""), criterios=False)
    check(r.returncode == 1 and "HIDROWEB_TOKEN" in r.stdout, "sem token e sem terminal: pede a variável HIDROWEB_TOKEN")
    srv.shutdown()

    # amostra para o usuário (dados inventados)
    shutil.copy(out / "estacoes" / "02047001.xlsx", AQUI / "exemplo_SINTETICO_estacao_02047001.xlsx")
    print("\n" + ("TODOS OS TESTES PASSARAM" if not FALHAS else f"{len(FALHAS)} FALHA(S): " + "; ".join(FALHAS)))
    return 1 if FALHAS else 0


if __name__ == "__main__":
    sys.exit(main())
