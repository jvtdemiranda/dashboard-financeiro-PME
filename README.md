# Portfólio de Dashboards

[![Pipeline](https://github.com/jvtdemiranda/teste01/actions/workflows/pipeline.yml/badge.svg)](https://github.com/jvtdemiranda/teste01/actions/workflows/pipeline.yml)

Projetos de portfólio simulando o fluxo real de trabalhos freelance de
análise/engenharia de dados: receber uma base bruta (com os problemas que
ela realmente tem), tratar as inconsistências, e entregar um dashboard
gerencial a partir dela.

| # | Projeto | Cenário | Stack | Link | Status |
|---|---------|---------|-------|------|--------|
| 1 | [`dashboard-financeiro-pme`](dashboard-financeiro-pme/) | Fluxo de caixa, contas a pagar/receber e DRE simplificado para uma pequena empresa/prestador de serviço | Python (pandas) + Excel (openpyxl) | [dashboard-financeiro-pme.vercel.app](https://dashboard-financeiro-pme.vercel.app) | ✅ Concluído |

Cada projeto tem seu próprio README com o cenário simulado, as decisões de
tratamento de dados (e por quê) e os bugs reais encontrados no processo —
a ideia é documentar o processo inteiro, não só o resultado final.

## Publicação: 1 link dedicado por projeto

Cada dashboard tem seu **próprio projeto no Vercel**, isolado dos outros —
o link de um projeto abre só aquele dashboard, direto, sem listagem nem
acesso aos demais. Faz sentido porque cada dashboard pode ir pra um
cliente diferente: ninguém que abre o link de um projeto deve enxergar os
outros.

`build_site.py` (na raiz) ainda gera a pasta `site/` — cada dashboard em
`site/<projeto>/` — mas agora é só matéria-prima: cada `site/<projeto>/`
vira o Root Directory de um projeto Vercel separado.

```bash
python build_site.py   # regenera site/ a partir dos dashboards já commitados
```

O CI confere a cada push se `site/` está em dia com os dashboards — se
esquecer de rodar o comando acima depois de mudar um dashboard, o badge
fica vermelho avisando.

### Publicar um projeto novo no Vercel

Pra cada projeto novo do portfólio, cria-se um projeto Vercel novo
apontando só pra pasta dele:

1. [vercel.com](https://vercel.com) → **Add New... → Project** → selecione este repositório (`jvtdemiranda/teste01`) de novo — o Vercel deixa importar o mesmo repo várias vezes como projetos diferentes.
2. Dê um nome pro projeto (ex.: `dashboard-financeiro-pme-2`, ou o nome da pasta do projeto).
3. **Root Directory** = `site/<nome-da-pasta-do-projeto>`.
4. Framework: **Other**, sem build.
5. Nas configurações do projeto (Settings → Deployment Protection), desligue **Vercel Authentication** — por padrão ele vem ligado e pede login do Vercel pra abrir o link.
6. **Deploy**.

O link fica em `https://<nome-do-projeto>.vercel.app`, atualiza sozinho a
cada push na `main`, e não enxerga nada dos outros projetos.

> O projeto `teste01` no Vercel (a pasta `site/` inteira, com a listagem de
> todos os dashboards) continua existindo — é só pra seu uso pessoal
> conferir tudo de uma vez, **não é o link que vai pro cliente**.
