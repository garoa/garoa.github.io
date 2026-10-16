"""Gera dados.js (agregado, sem dados pessoais) a partir da planilha de prestação de contas.

Uso (na pasta contas/):
    python3 -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt
    .venv/bin/python scripts/extrair.py

Para incluir mais meses, acrescente as abas em MESES e rode de novo.
"""
import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
PLANILHA = RAIZ / "planilhas" / "Garoa HC - Prestação de Contas 2026.xlsx"
SAIDA = RAIZ / "dados.js"
ANO = 2026
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set"]
# Mês em andamento (ou None): entram só os lançamentos com data, os sem data ainda não caíram.
MES_PARCIAL = "Out"

# Projeção dos meses que ainda não fecharam (usada em "quanto falta arrecadar").
# Aluguel: valor conhecido do boleto. IPTU (R$ 429,67) vem de fev a nov, dezembro e janeiro não têm.
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
ABAS_SEGUINTES = ["Out", "Nov", "Dez"]  # anuidades pagas adiantadas (depois do mês parcial)

# Cenários do fechamento: pessimista = menor valor de cada mensalidade nos meses-base, sem doações;
# otimista = maior valor de cada mensalidade nos meses-base, doações pela média do ano sem eventos.

# Badge r4.2 (ESP32-S3): lote de 30, vendido na GambiConf e na 307.
# Custo por unidade da lista de materiais (BadgeGaroa_r4.2-ESP32-S3 lista mat, Rev. 0 / 24-09-26),
# com a PCB pela cotação de 30 placas. O clube paga os componentes de 20; os outros 10 são doados.
# Itens marcados como "Estoque" na lista não têm custo. Frete e embalagem: +10% sobre o custo.
BADGE = {
    "lote": 30,
    "unidadesPagasPeloClube": 20,
    "custoUnidade": 87.7655,
    "frete": 0.10,
    "preco": 150.0,
    "taxaCartao": 0.05,  # supõe que parte das vendas é no crédito/débito
    "eventos": ["GambiConf", "307"],  # ordem de exibição (o JS ordena chaves numéricas primeiro)
    "vendas": {
        "pessimista": {"GambiConf": 4, "307": 8},
        "otimista": {"GambiConf": 10, "307": 20},
    },
}

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
    (r"bradesco|seguro", "Seguro"),
    (r"cofins|tfe|certificado|imposto|taxa", "Impostos e taxas"),
]


# "Onde comemos bola" (dos extratos):
# Dinheiro parado: o que ficou na conta acima de um colchão que cobre a maior saída líquida dentro de um mês
# (sem resgates/aplicações), com 15% de folga, arredondado para cima em R$ 500. O rendimento perdido usa
# o CDB 100% CDI de liquidez diária (R$ 6,90 por dia útil sobre R$ 13.357,93 em set/2026) e IR de 22,5%.
RENDIMENTO_DIA_UTIL = 6.90 / 13357.93
IR_CURTO_PRAZO = 0.225
# Aluguel pago com atraso: mês → valor do boleto sem multa (o pago a mais é multa e juros).
ALUGUEL_ATRASADO = {"2026-09": 4290.27}


def dinheiro_parado(conta, saldo_inicial):
    from datetime import timedelta
    mov, necessidade = {}, {}
    for mes, c in sorted(conta.items()):
        acc = pior = 0.0
        for x in sorted(c["lancamentos"], key=lambda x: x["dia"]):
            d = date(int(mes[:4]), x["mes"], x["dia"])
            mov[d] = mov.get(d, 0) + x["valor"]
            if not re.search(r"cdb|recusado|estornado", x["descricao"], re.I):
                acc += x["valor"]
                pior = min(pior, acc)
        necessidade[mes] = round(-pior, 2)
    colchao = -(-max(necessidade.values()) * 1.15 // 500) * 500
    ultimo = max(conta)
    dia = min(mov).replace(day=1)
    fim = date(int(ultimo[:4]) + int(ultimo[5:]) // 12, int(ultimo[5:]) % 12 + 1, 1) - timedelta(days=1)
    saldo, soma_saldo, dias, perda = saldo_inicial, 0.0, 0, 0.0
    while dia <= fim:
        saldo += mov.get(dia, 0)
        soma_saldo += saldo
        dias += 1
        if dia.weekday() < 5:
            perda += max(0.0, saldo - colchao) * RENDIMENTO_DIA_UTIL
        dia += timedelta(days=1)
    return {
        "colchao": colchao,
        "maiorNecessidade": max(necessidade.values()),
        "mesMaiorNecessidade": max(necessidade, key=necessidade.get),
        "saldoMedio": round(soma_saldo / dias, 2),
        "perdaBruta": round(perda, 2),
        "perdaLiquida": round(perda * (1 - IR_CURTO_PRAZO), 2),
        "taxaDiaUtil": RENDIMENTO_DIA_UTIL,
        "de": min(conta), "ate": max(conta),
    }


def aluguel_atrasado(conta):
    saida = []
    for mes, boleto in ALUGUEL_ATRASADO.items():
        for x in conta.get(mes, {}).get("lancamentos", []):
            if re.search(r"imobiliaria", x["descricao"], re.I) and x["valor"] < 0 and "recusado" not in x["descricao"].lower():
                saida.append({"mes": mes, "dia": x["dia"], "pago": -x["valor"], "boleto": boleto,
                              "extra": round(-x["valor"] - boleto, 2)})
    return saida


def conferidos(fechados, conta, inv, saldo_inicial):
    """Meses em que o saldo do fim do mês na planilha bate com o extrato (conta corrente e investimentos).
    A conferência lançamento a lançamento fica em scripts/conferir_extratos.py."""
    ok_conta, ok_inv, saldo = [], [], saldo_inicial
    for m in fechados:
        c = conta.get(m["mes"])
        if not c:
            break
        saldo = round(saldo + c["entradas"] - c["saidas"], 2)
        if abs(saldo - (m["saldoConta"] or 0)) > 0.01:
            raise SystemExit(f"{m['mes']}: saldo da planilha ({m['saldoConta']}) não bate com o extrato ({saldo})")
        ok_conta.append(m["mes"])
    for m in fechados:
        e = inv.get(m["mes"])
        if e and m["saldoInvestimentos"] is not None:
            if abs(e["saldoFinal"] - m["saldoInvestimentos"]) > 0.01:
                raise SystemExit(f"{m['mes']}: investimentos da planilha não batem com o extrato")
            ok_inv.append(m["mes"])
    return {"conta": ok_conta, "investimentos": ok_inv}


def dados_dos_extratos(fechados, saldo_inicial):
    """Rendimento dos investimentos e saldo do Mercado Pago, lidos dos PDFs de extratos/ (fora do git).
    Sem a pasta (ou sem pdfplumber), mantém o que já estava publicado no dados.js."""
    try:
        import extratos
        ext = extratos.todos()
    except ImportError:
        ext = {}
    if not ext:
        anterior = SAIDA.read_text(encoding="utf-8") if SAIDA.exists() else ""
        m = re.search(r'"extratos": (\{.*?\n  \})', anterior, re.S)
        return json.loads(m[1]) if m else None
    meses_fechados = fechados
    fechados = [m["mes"] for m in fechados]
    inv = {m: e for m, e in ext.get("investimentos", {}).items() if m in fechados}
    mp = ext.get("mercadoPago", {})
    ult = mp[max(mp)] if mp else None
    conta = {m: c for m, c in ext.get("conta", {}).items() if m in fechados}
    return {
        "conferidos": conferidos(meses_fechados, conta, inv, saldo_inicial),
        "dinheiroParado": dinheiro_parado(conta, saldo_inicial) if conta else None,
        "aluguelAtrasado": aluguel_atrasado(conta),
        "rendimento": {m: e["rendimento"] for m, e in sorted(inv.items())},
        "irRetido": round(sum(e["ir"] for e in inv.values()), 2),
        "mercadoPago": {"mes": ult["mes"], "saldo": ult["saldoFinal"],
                        "saldoInicial": mp[min(mp)]["saldoInicial"]} if ult else None,
    }


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

    abas = MESES + ([MES_PARCIAL] if MES_PARCIAL else [])
    linhas_resumo = list(resumo.iter_rows(min_row=4, max_row=3 + len(abas), values_only=True))
    meses = []
    for i, (aba, linha) in enumerate(zip(abas, linhas_resumo), start=1):
        ws = wb[aba]
        parcial = aba == MES_PARCIAL
        receitas, despesas = {}, {}
        investimento = extraordinario = 0.0
        ultima_data = None
        for data, oque, valor, tipo, obs in ws.iter_rows(min_row=2, max_col=5, values_only=True):
            valor = num(valor) or 0.0
            if not oque or not valor:
                continue
            if parcial:
                if not isinstance(data, datetime):
                    continue
                ultima_data = max(ultima_data or data, data)
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
        if parcial:
            # A aba Resumo soma também os lançamentos sem data; confere com o "Saldo hoje".
            if abs(num(ws["I3"].value) + rec - desp + investimento - num(ws["I4"].value)) > 0.01:
                raise SystemExit(f"{aba}: lançamentos com data não batem com o saldo de hoje")
        elif abs(rec - (num(linha[1]) or 0)) > 0.01 or abs(desp + (num(linha[3]) or 0)) > 0.01:
            raise SystemExit(f"{aba}: total diverge da aba Resumo ({rec:.2f} / {desp:.2f})")

        meses.append({
            "mes": f"{ANO}-{i:02d}",
            **({"parcial": True, "ate": ultima_data.date().isoformat()} if parcial else {}),
            "receitas": receitas,
            "despesas": despesas,
            "investimento": round(investimento, 2),
            "receitaExtraordinaria": round(extraordinario, 2),
            "saldoConta": num(ws["I4"].value),          # "Saldo hoje" da aba do mês
            "saldoInvestimentos": num(linha[6]),        # bruto
            "saldoInvestimentosLiquido": num(linha[8]),  # após IR, quando informado
            "observacao": linha[9] or "",
        })

    # Médias e previsões usam só os meses fechados.
    fechados = [m for m in meses if not m.get("parcial")]
    parcial = meses[-1] if MES_PARCIAL else None

    # Projeção dos meses restantes; no mês parcial, só o que ainda não foi pago.
    medias = {cat: round(sum(m["despesas"].get(cat, 0) for m in fechados) / len(fechados), 2)
              for cat in CATEGORIAS_PELA_MEDIA}
    ultimos = {cat: next(m["despesas"][cat] for m in reversed(fechados) if m["despesas"].get(cat))
               for cat in CATEGORIAS_PELO_ULTIMO}
    def previstas(aluguel):
        return {"Aluguel": round(aluguel, 2)} | FIXOS_PREVISTOS | medias | ultimos

    projecao = []
    for mes in range(len(MESES) + 1, 13):
        despesas = previstas(ALUGUEL_PREVISTO.get(mes, ALUGUEL_BASE + IPTU))
        if parcial and mes == len(MESES) + 1:
            despesas = {cat: round(v - parcial["despesas"].get(cat, 0), 2) for cat, v in despesas.items()
                        if v - parcial["despesas"].get(cat, 0) > 0.01}
        projecao.append({"mes": f"{ANO}-{mes:02d}", "despesas": despesas})
    # Contas de janeiro do ano seguinte (sem IPTU): devem estar na conta ao fim de dezembro,
    # já que as anuidades só entram em fevereiro.
    contas_janeiro = {"mes": f"{ANO + 1}-01", "despesas": previstas(ALUGUEL_BASE)}

    base = fechados[-MESES_BASE_RECORRENTES:]
    receita_mensal = {cat: round(sum(m["receitas"].get(cat, 0) for m in base) / len(base), 2)
                      for cat in RECORRENTES}
    receita_mensal[DOACOES] = round(
        sum(m["receitas"].get(DOACOES, 0) - m["receitaExtraordinaria"] for m in fechados) / len(fechados), 2)
    anuidades_confirmadas = 0.0
    for aba in ABAS_SEGUINTES[len(abas) - len(MESES): 12 - len(MESES)]:
        for data, oque, valor, tipo, obs in wb[aba].iter_rows(min_row=2, max_col=5, values_only=True):
            if oque and data and (num(valor) or 0) > 0 and re.search(r"anua?idade", str(oque).lower()):
                anuidades_confirmadas += valor
    cenarios = {
        "pessimista": {cat: min(m["receitas"].get(cat, 0) for m in base) for cat in RECORRENTES}
                      | {DOACOES: 0.0},
        "otimista": {cat: max(m["receitas"].get(cat, 0) for m in base) for cat in RECORRENTES}
                    | {DOACOES: receita_mensal[DOACOES]},
    }
    previsao_receitas = {
        "mensal": receita_mensal,
        "mesesBase": [m["mes"] for m in base],
        "anuidadesConfirmadas": round(anuidades_confirmadas, 2),
        "cenarios": cenarios,
    }
    if parcial:
        # O que ainda deve entrar no mês parcial: cenário menos o que já entrou.
        previsao_receitas["restanteParcial"] = {
            c: {cat: round(max(0.0, v - parcial["receitas"].get(cat, 0)), 2) for cat, v in cen.items()}
            for c, cen in cenarios.items()}

    custo_clube = BADGE["unidadesPagasPeloClube"] * BADGE["custoUnidade"]
    preco_liquido = BADGE["preco"] * (1 - BADGE["taxaCartao"])
    badge = {
        **BADGE,
        "componentes": round(custo_clube, 2),
        "custoClube": round(custo_clube * (1 + BADGE["frete"]), 2),
        "precoLiquido": round(preco_liquido, 2),
        "receita": {c: round(sum(v.values()) * preco_liquido, 2) for c, v in BADGE["vendas"].items()},
    }

    dados = {
        "ano": ANO,
        "atualizadoEm": date.today().isoformat(),
        "saldoInicialConta": num(wb["Jan"]["I3"].value),
        "saldoInicialInvestimentos": num(resumo["G3"].value),
        "meses": meses,
        "projecao": projecao,
        "contasJaneiro": contas_janeiro,
        "previsaoReceitas": previsao_receitas,
        "badge": badge,
        "extratos": dados_dos_extratos(fechados, num(wb["Jan"]["I3"].value)),
    }
    with open(SAIDA, "w", encoding="utf-8") as f:
        f.write("// Gerado por scripts/extrair.py — apenas valores agregados por categoria.\n")
        f.write("window.DADOS = ")
        json.dump(dados, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    for m in meses:
        print(m["mes"], f"receitas {sum(m['receitas'].values()):10.2f}",
              f"despesas {sum(m['despesas'].values()):10.2f}",
              f"parcial até {m['ate']} (confere com saldo hoje {m['saldoConta']:.2f})" if m.get("parcial")
              else "ok (confere com Resumo)")
    for p in projecao:
        print(p["mes"], f"despesa prevista {sum(p['despesas'].values()):10.2f}",
              f"(aluguel {p['despesas'].get('Aluguel', 0):.2f})")
    print(contas_janeiro["mes"], f"contas de janeiro {sum(contas_janeiro['despesas'].values()):10.2f}")
    print("receita prevista/mês", receita_mensal, "anuidades confirmadas", anuidades_confirmadas)
    for c, v in previsao_receitas["cenarios"].items():
        print(f"cenário {c}: {sum(v.values()):.2f}/mês, badge receita {badge['receita'][c]:.2f}")
    print(f"badge custo do clube {badge['custoClube']:.2f}")
    print("Gerado", SAIDA.relative_to(RAIZ))


if __name__ == "__main__":
    main()
