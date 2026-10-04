"""Gera dados.js (agregado, sem dados pessoais) a partir da planilha de prestação de contas.

Uso (na pasta contas/):
    python3 -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt
    .venv/bin/python scripts/extrair.py

Para incluir mais meses, acrescente as abas em MESES e rode de novo.
"""
import json
import re
from datetime import date
from pathlib import Path

import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
PLANILHA = RAIZ / "planilhas" / "Garoa HC - Prestação de Contas 2026.xlsx"
SAIDA = RAIZ / "dados.js"
ANO = 2026
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set"]

# Projeção dos meses que ainda não fecharam (usada em "quanto falta arrecadar").
# Aluguel: valor conhecido do boleto. IPTU (R$ 429,67) vem de fev a nov, dezembro não tem.
# Contabilidade: honorário fixo mensal. Água e internet: último valor pago. Luz: média mensal.
# Impostos e "outros" não entram na previsão.
IPTU = 429.67
ALUGUEL_BASE = 4290.27 - IPTU
ALUGUEL_PREVISTO = {10: ALUGUEL_BASE + IPTU, 11: ALUGUEL_BASE + IPTU, 12: ALUGUEL_BASE}
FIXOS_PREVISTOS = {"Contabilidade": 398.20}
CATEGORIAS_PELA_MEDIA = ["Luz"]
CATEGORIAS_PELO_ULTIMO = ["Água", "Internet"]

# Previsão de receitas dos meses restantes:
# mensalidades = média dos últimos N meses fechados; doações/lojinha = média do ano sem eventos
# extraordinários; anuidades = só as que já foram recebidas nas abas dos meses seguintes.
RECORRENTES = ["Mensalidades (depósito)", "Mensalidades (PayPal)", "Apoia-se"]
MESES_BASE_RECORRENTES = 3
DOACOES = "Doações, eventos e lojinha"
EXTRAORDINARIO = r"bsides"
ABAS_SEGUINTES = ["Out", "Nov", "Dez"]

# Ordem importa: a primeira regra que casar define a categoria.
REGRAS_RECEITA = [
    (r"anua?idade", "Anuidades"),
    (r"mensalidade", "Mensalidades (depósito)"),
    (r"paypal", "Mensalidades (PayPal)"),
    (r"apoia-se", "Apoia-se"),
    (r"mercado ?livre|\bml\b|venda", "Outras vendas (Mercado Livre)"),
    (r"doa[çc][ãa]o|bsides|oficina|lojinha", "Doações, eventos e lojinha"),
]
REGRAS_DESPESA = [
    (r"aluguel", "Aluguel"),
    (r"arlete|contador|contabil", "Contabilidade"),
    (r"enel", "Luz"),
    (r"sabesp", "Água"),
    (r"vivo", "Internet"),
    (r"cofins|tfe|certificado|imposto|taxa", "Impostos e taxas"),
]


def classificar(texto, regras, padrao):
    t = texto.lower()
    for padrao_re, cat in regras:
        if re.search(padrao_re, t):
            return cat
    return padrao


def num(v):
    return v if isinstance(v, (int, float)) else None


def main():
    wb = openpyxl.load_workbook(PLANILHA, data_only=True)
    resumo = wb["Resumo"]
    # Linha 4 da aba Resumo = janeiro
    linhas_resumo = list(resumo.iter_rows(min_row=4, max_row=3 + len(MESES), values_only=True))

    meses = []
    for i, (aba, linha) in enumerate(zip(MESES, linhas_resumo), start=1):
        ws = wb[aba]
        receitas, despesas = {}, {}
        investimento = extraordinario = 0.0
        for data, oque, valor, tipo, obs in ws.iter_rows(min_row=2, max_col=5, values_only=True):
            valor = num(valor) or 0.0
            if not oque or not valor:
                continue
            texto = f"{oque} {obs or ''}"
            if tipo == "Investimento":
                investimento += valor  # + resgate, - aplicação
            elif valor > 0:
                cat = classificar(texto, REGRAS_RECEITA, "Outras receitas")
                receitas[cat] = round(receitas.get(cat, 0) + valor, 2)
                if re.search(EXTRAORDINARIO, texto.lower()):
                    extraordinario += valor
            else:
                cat = classificar(texto, REGRAS_DESPESA, "Outros")
                despesas[cat] = round(despesas.get(cat, 0) - valor, 2)

        rec, desp = sum(receitas.values()), sum(despesas.values())
        if abs(rec - (num(linha[1]) or 0)) > 0.01 or abs(desp + (num(linha[3]) or 0)) > 0.01:
            raise SystemExit(f"{aba}: total diverge da aba Resumo ({rec:.2f} / {desp:.2f})")

        meses.append({
            "mes": f"{ANO}-{i:02d}",
            "receitas": receitas,
            "despesas": despesas,
            "investimento": round(investimento, 2),
            "receitaExtraordinaria": round(extraordinario, 2),
            "saldoConta": num(ws["I4"].value),          # "Saldo hoje" da aba do mês
            "saldoInvestimentos": num(linha[6]),        # bruto
            "saldoInvestimentosLiquido": num(linha[8]),  # após IR, quando informado
            "observacao": linha[9] or "",
        })

    # Projeção dos meses restantes
    medias = {cat: round(sum(m["despesas"].get(cat, 0) for m in meses) / len(meses), 2)
              for cat in CATEGORIAS_PELA_MEDIA}
    ultimos = {cat: next(m["despesas"][cat] for m in reversed(meses) if m["despesas"].get(cat))
               for cat in CATEGORIAS_PELO_ULTIMO}
    projecao = []
    for mes in range(len(MESES) + 1, 13):
        despesas = {"Aluguel": round(ALUGUEL_PREVISTO.get(mes, ALUGUEL_BASE + IPTU), 2)}
        despesas.update(FIXOS_PREVISTOS)
        despesas.update(medias)
        despesas.update(ultimos)
        projecao.append({"mes": f"{ANO}-{mes:02d}", "despesas": despesas})

    base = meses[-MESES_BASE_RECORRENTES:]
    receita_mensal = {cat: round(sum(m["receitas"].get(cat, 0) for m in base) / len(base), 2)
                      for cat in RECORRENTES}
    receita_mensal[DOACOES] = round(
        sum(m["receitas"].get(DOACOES, 0) - m["receitaExtraordinaria"] for m in meses) / len(meses), 2)
    anuidades_confirmadas = 0.0
    for aba in ABAS_SEGUINTES[: 12 - len(MESES)]:
        for data, oque, valor, tipo, obs in wb[aba].iter_rows(min_row=2, max_col=5, values_only=True):
            if oque and data and (num(valor) or 0) > 0 and re.search(r"anua?idade", str(oque).lower()):
                anuidades_confirmadas += valor
    previsao_receitas = {
        "mensal": receita_mensal,
        "mesesBase": [m["mes"] for m in base],
        "anuidadesConfirmadas": round(anuidades_confirmadas, 2),
    }

    dados = {
        "ano": ANO,
        "atualizadoEm": date.today().isoformat(),
        "saldoInicialConta": num(wb["Jan"]["I3"].value),
        "saldoInicialInvestimentos": num(resumo["G3"].value),
        "meses": meses,
        "projecao": projecao,
        "previsaoReceitas": previsao_receitas,
    }
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write("// Gerado por scripts/extrair.py — apenas valores agregados por categoria.\n")
        f.write("window.DADOS = ")
        json.dump(dados, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    for m in meses:
        print(m["mes"], f"receitas {sum(m['receitas'].values()):10.2f}",
              f"despesas {sum(m['despesas'].values()):10.2f}", "ok (confere com Resumo)")
    for p in projecao:
        print(p["mes"], f"despesa prevista {sum(p['despesas'].values()):10.2f}",
              f"(aluguel {p['despesas'].get('Aluguel', 0):.2f})")
    print("receita prevista/mês", receita_mensal, "anuidades confirmadas", anuidades_confirmadas)
    print("Gerado", SAIDA.relative_to(RAIZ))


if __name__ == "__main__":
    main()
