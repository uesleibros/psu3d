# -*- coding: utf-8 -*-
"""Gera a copia do wiki a partir de docs/.

A documentacao mora em docs/, onde os links apontam para arquivo .md e funcionam
navegando o repositorio. O wiki do GitHub quer link por nome de pagina, sem .md,
e quer a Home chamada Home.md. Este script faz essa traducao.

Uso:

    python tools/wiki-sync.py ../psu3d-wiki

Depois, dentro da pasta de saida:

    git add -A && git commit -m "atualiza wiki" && git push

O repositorio do wiki so passa a existir depois que a primeira pagina e criada
pela interface do GitHub, em Wiki e depois Create the first page. Uma vez.
"""
import io
import os
import re
import sys
import glob

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(os.path.dirname(HERE), "docs")

SIDEBAR = """**Psu3D**

[Home](Home)

**Comecando**

[Instalacao](Instalacao)
[Primeiros passos](Primeiros-passos)
[Conceitos](Conceitos)
[Como o 3D funciona](Como-o-3D-funciona)

**Guias**

[Canvas](Canvas)
[Camera](Camera)
[Renderer e primitivas](Renderer-e-primitivas)
[Cena](Cena)
[Materiais](Materiais)
[Luz e nevoa](Luz-e-nevoa)
[Corpo e fisica](Corpo-e-fisica)
[Fases em JSON](Fases-em-JSON)

**Por dentro**

[Ordenacao por profundidade](Ordenacao-por-profundidade)
[Indice espacial](Indice-espacial)
[Performance](Performance)
[Limites conhecidos](Limites-conhecidos)

**Pratica**

[Receitas](Receitas)
[Autoteste](Autoteste)

**Referencia**

[PCore](API-PCore)
[PLighting](API-PLighting)
[PMaterial](API-PMaterial)
[PMaterials](API-PMaterials)
[PCanvas](API-PCanvas)
[PCamera](API-PCamera)
[PRenderer](API-PRenderer)
[PScene](API-PScene)
[PBody](API-PBody)
[PLevel](API-PLevel)
[Psu3D](API-Psu3D)
[PSelfTest](API-PSelfTest)
"""


def main(out_dir):
    if not os.path.isdir(DOCS):
        print("docs/ nao encontrado em", DOCS)
        return 1

    os.makedirs(out_dir, exist_ok=True)
    written = 0

    for path in sorted(glob.glob(os.path.join(DOCS, "*.md"))):
        name = os.path.basename(path)
        text = io.open(path, encoding="utf-8", newline="").read().replace("\r\n", "\n")

        # link de arquivo vira link de pagina
        text = re.sub(r"\]\((?!http)([A-Za-z0-9._-]+)\.md(#[^)]*)?\)",
                      lambda m: "](%s%s)" % (m.group(1), m.group(2) or ""), text)
        # o indice do repositorio e a Home do wiki
        text = text.replace("](README)", "](Home)")

        target = "Home.md" if name == "README.md" else name
        io.open(os.path.join(out_dir, target), "w", encoding="utf-8", newline="\n").write(text)
        written += 1

    io.open(os.path.join(out_dir, "_Sidebar.md"), "w", encoding="utf-8", newline="\n").write(SIDEBAR)
    print("%d paginas escritas em %s" % (written + 1, out_dir))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
