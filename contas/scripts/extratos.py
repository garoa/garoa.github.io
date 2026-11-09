"""Lê os extratos em PDF da pasta extratos/ (fora do git): conta corrente e investimentos do C6
e conta do Mercado Pago. O tipo e o mês saem do conteúdo do PDF, não do nome do arquivo.

Também lê o CSV de transações exportado do Conta Aberta (open finance), que cobre vários meses
da conta corrente. Ele só é usado nos meses sem extrato em PDF (na prática, o mês em andamento).

Senhas: números no nome do arquivo (ex.: 2026-09_141246.pdf) e as linhas de extratos/senhas.txt.
"""
import csv
import re
from pathlib import Path

import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect
from pdfplumber.utils.exceptions import PdfminerException

PASTA = Path(__file__).resolve().parent.parent / "extratos"
VALOR = r"(-?)R\$ ?(-?[\d.]+,\d\d)"
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
         "setembro", "outubro", "novembro", "dezembro"]


def br(s):
    return float(s.replace(".", "").replace(",", "."))


def valor(sinal, numero):
    v = br(numero)
    return -abs(v) if sinal or v < 0 else v


def senhas(arquivo):
    extras = PASTA / "senhas.txt"
    lista = re.findall(r"\d{4,}", arquivo.stem)[1:]  # o primeiro é o ano
    if extras.exists():
        lista += extras.read_text().split()
    return [""] + lista


def texto(arquivo):
    for senha in senhas(arquivo):
        try:
            with pdfplumber.open(arquivo, password=senha) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages)
        except PdfminerException as e:
            if not isinstance(e.args[0], PDFPasswordIncorrect):
                raise

    raise SystemExit(f"{arquivo.name}: nenhuma senha funcionou (veja extratos/senhas.txt)")


def conta_c6(t):
    """Extrato da conta corrente: lançamentos, totais do cabeçalho e saldo informado no dia da exportação.
    Aceita os dois layouts do C6 (até set/2026 "Extrato Período • ...", depois "Período da Exportação")."""
    per = re.search(r"\( \d\d/(\d\d)/(\d{4}) - (\d\d)/\d\d/\d{4} \) Entradas", t)
    tot = re.search(r"Entradas: " + VALOR + r" • Saídas: " + VALOR, t)
    hoje = re.search(r"Saldo do dia • (\d+) de (\w+) de (\d{4}) • " + VALOR, t)
    if hoje:
        hoje = {"data": f"{hoje[3]}-{MESES.index(hoje[2].lower()) + 1:02d}-{int(hoje[1]):02d}",
                "valor": valor(hoje[4], hoje[5])}
    else:
        exp = re.search(r"Exportado em (\d\d)/(\d\d)/(\d\d)", t)
        sal = re.search(r"\d\d/\d\d/\d{4} a \d\d/\d\d/\d{4} " + VALOR, t)
        hoje = {"data": f"20{exp[3]}-{exp[2]}-{exp[1]}", "valor": valor(sal[1], sal[2])} if exp and sal else None
    lancamentos = []
    for linha in t.splitlines():
        m = re.match(r"^(\d\d)/(\d\d) \d\d/\d\d (.*?) " + VALOR + "$", linha)
        if m:
            lancamentos.append({"dia": int(m[1]), "mes": int(m[2]), "descricao": m[3], "valor": valor(m[4], m[5])})
    return {
        "tipo": "conta",
        "mes": f"{per[2]}-{per[1]}",
        "origem": "extrato do banco",
        "ate": f"{per[2]}-{per[1]}-{per[3]}",
        "entradas": valor(tot[1], tot[2]),
        "saidas": abs(valor(tot[3], tot[4])),
        "lancamentos": lancamentos,
        "saldoExportacao": hoje,
    }


def investimentos_c6(t):
    """Extrato de aplicações: totais do período por título e da conta."""
    per = re.search(r"Período: de (\d\d)/(\d\d)/(\d{4}) a (\d\d)/(\d\d)/(\d{4})", t)
    campos = ["saldoAnterior", "aplicacao", "rendimento", "resgate", "ir", "iof", "difResgate",
              "saldoFinal", "irSaldo", "iofSaldo"]
    tot = re.search(r"Total Conta [\d/]+: ((?:[\d.]+,\d\d ?){10})", t)
    titulos = re.findall(r"Título:(.+?)\n", t)
    # O mês é o do fim do período; um período que termina no dia 1º (ex.: 30/04 a 01/06) é do mês anterior.
    ano, mes = int(per[6]), int(per[5])
    if per[4] == "01":
        ano, mes = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
    return {
        "tipo": "investimentos",
        "mes": f"{ano}-{mes:02d}",
        "inicio": f"{per[3]}-{per[2]}-{per[1]}",
        "fim": f"{per[6]}-{per[5]}-{per[4]}",
        "titulos": sorted({re.sub(r"\s+", " ", x).strip() for x in titulos if "%" in x}),
        **dict(zip(campos, (br(x) for x in tot[1].split()))),
    }


def mercado_pago(t):
    per = re.search(r"Periodo: De \d\d-(\d\d)-(\d{4})", t)
    saldo = re.search(r"Saldo inicial: " + VALOR + r" Saldo final: " + VALOR, t)
    ent = re.search(r"Entradas: " + VALOR, t)
    sai = re.search(r"Saidas: " + VALOR, t)
    return {
        "tipo": "mercadoPago",
        "mes": f"{per[2]}-{per[1]}",
        "saldoInicial": valor(saldo[1], saldo[2]),
        "saldoFinal": valor(saldo[3], saldo[4]),
        "entradas": valor(ent[1], ent[2]),
        "saidas": abs(valor(sai[1], sai[2])),
    }


def ler(arquivo):
    t = texto(arquivo)
    if "Extrato de Aplicações" in t:
        return investimentos_c6(t)
    if re.search(r"\) Entradas: R\$", t):
        return conta_c6(t)
    if "mercadopago" in t.lower():
        return mercado_pago(t)
    raise SystemExit(f"{arquivo.name}: tipo de extrato desconhecido")


def conta_aberta(arquivo):
    """CSV do Conta Aberta: {mes: extrato da conta}, no mesmo formato de conta_c6. Não traz saldo."""
    meses = {}
    with open(arquivo, encoding="utf-8-sig", newline="") as f:
        for linha in csv.DictReader(f, delimiter=";"):
            dia, mes, ano = (int(x) for x in linha["Data"].split("/"))
            meses.setdefault(f"{ano}-{mes:02d}", []).append(
                {"dia": dia, "mes": mes, "descricao": linha["Descrição"].strip(), "valor": float(linha["Valor"]),
                 "categoria": linha["Categoria"].strip()})
    saida = {}
    for mes, lanc in meses.items():
        lanc.sort(key=lambda x: x["dia"])
        saida[mes] = {
            "tipo": "conta",
            "mes": mes,
            "origem": "Conta Aberta",
            "ate": f"{mes}-{lanc[-1]['dia']:02d}",
            "entradas": round(sum(x["valor"] for x in lanc if x["valor"] > 0), 2),
            "saidas": round(-sum(x["valor"] for x in lanc if x["valor"] < 0), 2),
            "lancamentos": lanc,
            "saldoExportacao": None,
        }
    return saida


def eh_conta_aberta(arquivo):
    return arquivo.suffix.lower() == ".csv" and arquivo.read_bytes()[:40].decode("utf-8-sig", "ignore") \
        .startswith("Data;Descrição;Tipo;Categoria;Valor")


def todos():
    """{tipo: {mes: extrato}} de todos os PDFs e CSVs do Conta Aberta da pasta (inclusive subpastas).
    Vazio se a pasta não existe. No mês que tem PDF da conta corrente, o PDF vale e o CSV é ignorado."""
    saida, csvs = {}, []
    if not PASTA.exists():
        return saida
    for arquivo in sorted(PASTA.rglob("*")):
        if not arquivo.is_file():
            continue
        if arquivo.read_bytes()[:5] == b"%PDF-":  # há PDFs salvos como .jpg
            e = ler(arquivo)
            saida.setdefault(e["tipo"], {})[e["mes"]] = e
        elif eh_conta_aberta(arquivo):
            csvs.append(arquivo)
    pdfs = set(saida.get("conta", {}))
    for arquivo in sorted(csvs, key=lambda a: a.stat().st_mtime):  # o mais recente vale
        for mes, e in conta_aberta(arquivo).items():
            if mes not in pdfs:
                saida.setdefault("conta", {})[mes] = e
    return saida
