"""Extrai o logotipo (guarda-chuva) do Garoa do cartaz SVG da wiki e salva em assets/garoa-logo.svg.

Fonte: https://garoa.net.br/mediawiki/images/Garoa_Cartaz_inline.svg
O guarda-chuva usa a classe "guarda-chuva" e as letras a classe "letras",
para a página colorir cada parte via CSS.

Uso (na pasta contas/): python3 scripts/extrair_logo.py
"""
import re
import urllib.request
from pathlib import Path

URL = "https://garoa.net.br/mediawiki/images/Garoa_Cartaz_inline.svg"
SAIDA = Path(__file__).resolve().parent.parent / "assets" / "garoa-logo.svg"

svg = urllib.request.urlopen(URL, timeout=30).read().decode("utf-8")
inicio = svg.index("<!-- Logotipo do Garoa Hacker Clube -->")
grupo = svg[inicio:svg.rindex("</g>", 0, svg.rindex("</g>"))]  # até o fim do bloco do logo
grupo = re.sub(r"<!--.*?-->", "", grupo, flags=re.S)
grupo = grupo.replace('transform="translate(405.6, 631.9)" ', "")
grupo = grupo.replace('fill="#131718"', 'class="guarda-chuva"').replace('fill="#ffff4d"', 'class="letras"')
grupo = re.sub(r"\s+", " ", grupo).strip()
grupo += "</g>" * (grupo.count("<g") - grupo.count("</g>"))  # fecha o grupo externo

SAIDA.parent.mkdir(exist_ok=True)
SAIDA.write_text(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 162 187">{grupo}</svg>\n', encoding="utf-8")
print("Gerado", SAIDA)
