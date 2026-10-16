"""Confere a planilha de prestação de contas com os extratos em PDF de extratos/.

Uso (na pasta contas/):
    .venv/bin/python scripts/conferir_extratos.py

Para cada mês fechado: lançamentos da conta corrente (valor e data), totais e saldo do fim do mês,
e saldo dos investimentos. Também mostra o que está no Mercado Pago, que a planilha não acompanha.
Sai com código 1 se algo não bater.
"""
from collections import Counter
from datetime import datetime

import openpyxl

import extratos
from extrair import ANO, MES_PARCIAL, MESES, PLANILHA, num

brl = lambda v: f"R$ {v:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def main():
    wb = openpyxl.load_workbook(PLANILHA, data_only=True)
    resumo = list(wb["Resumo"].iter_rows(min_row=4, max_row=3 + len(MESES), values_only=True))
    ext = extratos.todos()
    conta, inv, mp = ext.get("conta", {}), ext.get("investimentos", {}), ext.get("mercadoPago", {})
    problemas = 0
    saldo = num(wb["Jan"]["I3"].value)

    def aviso(texto):
        nonlocal problemas
        problemas += 1
        print("  ✗", texto)

    for i, aba in enumerate(MESES, start=1):
        mes = f"{ANO}-{i:02d}"
        ws = wb[aba]
        print(f"== {aba}")
        plan = [(d, o, round(v, 2)) for d, o, v, *_ in ws.iter_rows(min_row=2, max_col=5, values_only=True)
                if o and num(v)]
        c = conta.get(mes)
        if not c:
            print("  - sem extrato da conta corrente")
        else:
            lanc = c["lancamentos"]
            # Pix recusado e estornado no mesmo valor se anulam; a planilha não os registra.
            estorno = Counter(round(-x["valor"], 2) for x in lanc if "estornado" in x["descricao"].lower())
            lanc = [x for x in lanc if not ("recusado" in x["descricao"].lower() and estorno[round(x["valor"], 2)])
                    and "estornado" not in x["descricao"].lower()]
            do_extrato = Counter(round(x["valor"], 2) for x in lanc)
            da_planilha = Counter(v for *_, v in plan)
            for v, n in (do_extrato - da_planilha).items():
                for x in [x for x in lanc if round(x["valor"], 2) == v][:n]:
                    aviso(f"só no extrato: {x['dia']:02d}/{x['mes']:02d} {x['descricao'][:60]} {brl(v)}")
            for v, n in (da_planilha - do_extrato).items():
                for d, o, _ in [p for p in plan if p[2] == v][:n]:
                    aviso(f"só na planilha: {d:%d/%m} {o} {brl(v)}" if isinstance(d, datetime) else f"só na planilha: {o} {brl(v)}")
            for d, o, v in plan:
                datas = [(x["dia"], x["mes"]) for x in lanc if round(x["valor"], 2) == v]
                if isinstance(d, datetime) and datas and (d.day, d.month) not in datas:
                    print(f"  · data diferente: {o} {brl(v)}: planilha {d:%d/%m}, extrato "
                          + ", ".join(f"{a:02d}/{b:02d}" for a, b in datas))
            saldo = round(saldo + c["entradas"] - c["saidas"], 2)
            if abs(saldo - num(ws["I4"].value)) > 0.01:
                aviso(f"saldo do fim do mês: extrato {brl(saldo)}, planilha {brl(num(ws['I4'].value))}")
            else:
                print(f"  ✓ {len(lanc)} lançamentos, saldo {brl(saldo)}")

        e = inv.get(mes)
        if e:
            p = num(resumo[i - 1][6])
            if p is None or abs(e["saldoFinal"] - p) > 0.01:
                aviso(f"investimentos: extrato {brl(e['saldoFinal'])}, planilha {brl(p) if p else 'vazio'}")
            else:
                print(f"  ✓ investimentos {brl(e['saldoFinal'])} (rendeu {brl(e['rendimento'])}"
                      + (f", IR retido {brl(e['ir'])}" if e["ir"] else "") + ")")
        else:
            print("  - sem extrato dos investimentos")

    if MES_PARCIAL:
        c = conta.get(f"{ANO}-{len(MESES):02d}")
        if c and c["saldoExportacao"]:
            s = c["saldoExportacao"]
            print(f"== {MES_PARCIAL} (parcial): saldo em {s['data']} no extrato exportado: {brl(s['valor'])}, "
                  f"planilha hoje: {brl(num(wb[MES_PARCIAL]['I4'].value))}")
    if mp:
        ult = mp[max(mp)]
        print(f"== Mercado Pago (fora da planilha): saldo {brl(ult['saldoFinal'])} em {ult['mes']}")
        for m, e in sorted(mp.items()):
            if e["entradas"] or e["saidas"]:
                print(f"  {m}: entrou {brl(e['entradas'])}, saiu {brl(e['saidas'])}")

    print("Tudo confere." if not problemas else f"{problemas} divergência(s).")
    raise SystemExit(1 if problemas else 0)


if __name__ == "__main__":
    main()
