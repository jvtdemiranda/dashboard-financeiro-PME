# Portfólio de Dashboards

[![Pipeline](https://github.com/jvtdemiranda/teste01/actions/workflows/pipeline.yml/badge.svg)](https://github.com/jvtdemiranda/teste01/actions/workflows/pipeline.yml)

Projetos de portfólio simulando o fluxo real de trabalhos freelance de
análise/engenharia de dados: receber uma base bruta (com os problemas que
ela realmente tem), tratar as inconsistências, e entregar um dashboard
gerencial a partir dela.

| # | Projeto | Cenário | Stack | Status |
|---|---------|---------|-------|--------|
| 1 | [`dashboard-financeiro-pme`](dashboard-financeiro-pme/) | Fluxo de caixa, contas a pagar/receber e DRE simplificado para uma pequena empresa/prestador de serviço | Python (pandas) + Excel (openpyxl) | ✅ Concluído |

Cada projeto tem seu próprio README com o cenário simulado, as decisões de
tratamento de dados (e por quê) e os bugs reais encontrados no processo —
a ideia é documentar o processo inteiro, não só o resultado final.

## Site com todos os dashboards

`build_site.py` (na raiz) junta o dashboard HTML de cada projeto numa
pasta `site/` — uma página inicial listando todos, e cada dashboard em
`site/<projeto>/`. É só uma pasta de arquivos estáticos, então funciona em
qualquer host gratuito (Vercel, Render, Netlify, GitHub Pages...); nenhum
build step é necessário do lado do host.

```bash
python build_site.py   # regenera site/ a partir dos dashboards já commitados
```

O CI confere a cada push se `site/` está em dia com os dashboards — se
esquecer de rodar o comando acima depois de mudar um dashboard, o badge
fica vermelho avisando.

### Publicado no Vercel

Já conectado ao GitHub — todo push na `main` atualiza o link sozinho.
Root Directory = `site`, sem build (HTML puro), proteção de acesso
desligada (senão o link pediria login do Vercel pra abrir). Não precisa
repetir nada disso pros próximos projetos: como `site/` já lista qualquer
pasta nova automaticamente, um dashboard novo aparece no mesmo link assim
que for commitado na `main`.
