"""Lê os extratos em PDF da pasta extratos/ (fora do git): conta corrente e investimentos do C6
e conta do Mercado Pago. O tipo e o mês saem do conteúdo do PDF, não do nome do arquivo.

Senhas: números no nome do arquivo (ex.: 2026-09_141246.pdf) e as linhas de extratos/senhas.txt.
"""
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
    """Extrato da conta corrente: lançamentos, totais do cabeçalho e saldo informado no dia da exportação."""
    per = re.search(r"Extrato Período • \d+ de (\w+) de (\d{4})", t)
    tot = re.search(r"Entradas: " + VALOR + r" • Saídas: " + VALOR, t)
    hoje = re.search(r"Saldo do dia • (\d+) de (\w+) de (\d{4}) • " + VALOR, t)
    lancamentos = []
    for linha in t.splitlines():
        m = re.match(r"^(\d\d)/(\d\d) \d\d/\d\d (.*?) " + VALOR + "$", linha)
        if m:
            lancamentos.append({"dia": int(m[1]), "mes": int(m[2]), "descricao": m[3], "valor": valor(m[4], m[5])})
    return {
        "tipo": "conta",
        "mes": f"{per[2]}-{MESES.index(per[1].lower()) + 1:02d}",
        "entradas": valor(tot[1], tot[2]),
        "saidas": abs(valor(tot[3], tot[4])),
        "lancamentos": lancamentos,
        "saldoExportacao": {"data": f"{hoje[3]}-{MESES.index(hoje[2].lower()) + 1:02d}-{int(hoje[1]):02d}",
                            "valor": valor(hoje[4], hoje[5])} if hoje else None,
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
    if "Extrato Período" in t:
        return conta_c6(t)
    if "mercadopago" in t.lower():
        return mercado_pago(t)
    raise SystemExit(f"{arquivo.name}: tipo de extrato desconhecido")


def todos():
    """{tipo: {mes: extrato}} de todos os PDFs da pasta (inclusive subpastas). Vazio se a pasta não existe."""
    saida = {}
    if not PASTA.exists():
        return saida
    for arquivo in sorted(PASTA.rglob("*")):
        if arquivo.is_file() and arquivo.read_bytes()[:5] == b"%PDF-":  # há PDFs salvos como .jpg
            e = ler(arquivo)
            saida.setdefault(e["tipo"], {})[e["mes"]] = e
    return saida
