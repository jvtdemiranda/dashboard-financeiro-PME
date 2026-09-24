"""
Monta a pasta site/: uma cópia estática, pronta pra publicar (Vercel,
Render, Netlify, GitHub Pages, qualquer host de arquivo estático), de
cada projeto do portfólio.

Convenção — qualquer pasta na raiz do repo entra automaticamente,
detectada de dois jeitos (um projeto novo não precisa de nenhuma mudança
neste script, só rodar ele de novo depois de gerar/editar o projeto):

- **Projeto de dado/dashboard**: tem `dashboard/*.html` (arquivo único,
  gerado por script a partir de dados) -> copiado como
  `site/<projeto>/index.html`.
- **Projeto de site estático**: tem `index.html` na própria raiz da
  pasta (referenciando `css/`, `js/`, etc.) -> a pasta inteira é copiada
  pra `site/<projeto>/`, exceto README.md e docs/ (que são só pro
  GitHub).

Rodar depois de qualquer mudança num projeto já commitado:
    python build_site.py

O CI (.github/workflows/pipeline.yml) roda isso e falha se o resultado
não bater com o que está commitado, pra pegar quem esquecer de rodar
antes de commitar.
"""

import html
import os
import shutil
from pathlib import Path

RAIZ = Path(__file__).parent
SITE = RAIZ / "site"

IGNORAR = {".git", ".github", "site", "node_modules"}
IGNORAR_NA_COPIA = shutil.ignore_patterns("README.md", "docs", ".*")


def encontrar_projetos():
    """Retorna, em ordem alfabética, cada projeto detectado e como publicá-lo.

    Cada item é (nome_da_pasta, "arquivo" | "pasta", origem).
    """
    projetos = []
    for item in sorted(RAIZ.iterdir()):
        if not item.is_dir() or item.name in IGNORAR or item.name.startswith("."):
            continue

        pasta_dashboard = item / "dashboard"
        if pasta_dashboard.is_dir():
            htmls = sorted(pasta_dashboard.glob("*.html"))
            if htmls:
                projetos.append((item.name, "arquivo", htmls[0]))
                continue

        if (item / "index.html").is_file():
            projetos.append((item.name, "pasta", item))

    return projetos


def montar_pagina_inicial(projetos: list[tuple[str, str, Path]]) -> str:
    itens = "\n".join(
        f'    <a href="./{html.escape(nome)}/">{html.escape(nome)}</a>'
        for nome, _, _ in projetos
    )
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Portfólio</title>
  <style>
    :root{{color-scheme:light dark; --bg:#f9f9f7; --surface:#fcfcfb; --ink:#0b0b0b; --ink-2:#52514e; --border:#e1e0d9; --blue:#2a78d6;}}
    @media (prefers-color-scheme: dark){{
      :root{{--bg:#181816; --surface:#221f1c; --ink:#f7f5f0; --ink-2:#c3c2b7; --border:#33312b; --blue:#3987e5;}}
    }}
    *{{box-sizing:border-box;}}
    body{{font-family:system-ui,-apple-system,"Segoe UI",sans-serif; max-width:640px; margin:0 auto; padding:48px 20px; line-height:1.5; background:var(--bg); color:var(--ink);}}
    h1{{font-size:24px; margin:0 0 6px;}}
    p.sub{{color:var(--ink-2); margin:0 0 28px; font-size:14px;}}
    a{{display:block; padding:16px 18px; margin:10px 0; border:1px solid var(--border); border-radius:12px; text-decoration:none; color:var(--ink); background:var(--surface); font-weight:600;}}
    a:hover{{border-color:var(--blue); color:var(--blue);}}
    footer{{margin-top:32px; font-size:12px; color:var(--ink-2);}}
  </style>
</head>
<body>
  <h1>Portfólio</h1>
  <p class="sub">Cada link abre o projeto completo, direto no navegador (funciona no celular).</p>
{itens}
  <footer>Gerado por build_site.py a partir de cada projeto do repositório.</footer>
</body>
</html>
"""


def main():
    if SITE.exists():
        shutil.rmtree(SITE)
    SITE.mkdir()

    projetos = encontrar_projetos()

    for nome, tipo, origem in projetos:
        pasta_destino = SITE / nome
        if tipo == "arquivo":
            pasta_destino.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(origem, pasta_destino / "index.html")
        else:
            shutil.copytree(origem, pasta_destino, ignore=IGNORAR_NA_COPIA)

    (SITE / "index.html").write_text(montar_pagina_inicial(projetos), encoding="utf-8")

    print(f"site/ montado com {len(projetos)} projeto(s):")
    for nome, tipo, origem in projetos:
        origem_relativa = origem.relative_to(RAIZ)
        print(f"  - {nome}  <-  {origem_relativa} ({tipo})")


if __name__ == "__main__":
    main()
