# Portfólio de Projetos Freelance

[![Pipeline](https://github.com/jvtdemiranda/teste01/actions/workflows/pipeline.yml/badge.svg)](https://github.com/jvtdemiranda/teste01/actions/workflows/pipeline.yml)

Projetos de portfólio simulando trabalhos reais de freelancer em dados e
web, escolhidos com base numa pesquisa de vagas reais no Workana e
99Freelas (não achismo — ver o README de cada projeto pra evidência).

| # | Projeto | Cenário | Stack | Link | Status |
|---|---------|---------|-------|------|--------|
| 1 | [`dashboard-financeiro-pme`](dashboard-financeiro-pme/) | Fluxo de caixa, contas a pagar/receber e DRE simplificado para uma pequena empresa/prestador de serviço | Python (pandas) + Excel (openpyxl) | [dashboard-financeiro-pme.vercel.app](https://dashboard-financeiro-pme.vercel.app) | ✅ Concluído |
| 2 | [`landing-prestador-servico`](landing-prestador-servico/) | Landing page com captação via WhatsApp e calculadora interativa, pro nicho mais recorrente encontrado na pesquisa (prestador de serviço local) | HTML/CSS/JS puro | _(publicar)_ | ✅ Concluído |
| 3 | Dashboard com dado ao vivo via API pública | Dashboard financeiro/PME, mas consumindo API real em vez de CSV estático | Python + REST API | — | 🔜 Planejado |
| 4 | Landing page para advocacia | Mesmo padrão do projeto 2, focado no vertical jurídico (ângulo: compliance com regras de publicidade da OAB) | HTML/CSS/JS | — | 🔜 Planejado |
| 5 | Landing page de produto / e-commerce | Página de produto único voltada pra tráfego pago (Google/Meta Ads) | HTML/CSS/JS | — | 🔜 Planejado |

Cada projeto tem seu próprio README com o cenário simulado, as decisões
de projeto (e por quê) e os bugs reais encontrados no processo — a ideia
é documentar o processo inteiro, não só o resultado final.

## Publicação: 1 link dedicado por projeto

Cada projeto tem seu **próprio projeto no Vercel**, isolado dos outros —
o link de um projeto abre só aquele projeto, direto, sem listagem nem
acesso aos demais. Faz sentido porque cada peça pode ir pra um cliente
diferente: ninguém que abre o link de um projeto deve enxergar os outros.

`build_site.py` (na raiz) ainda gera a pasta `site/` — cada projeto em
`site/<projeto>/` — mas agora é só matéria-prima: cada `site/<projeto>/`
vira o Root Directory de um projeto Vercel separado. O script reconhece
dois formatos automaticamente: projeto de dado que gera `dashboard/*.html`,
ou site estático com `index.html` na raiz da própria pasta.

```bash
python build_site.py   # regenera site/ a partir do que já está commitado
```

O CI confere a cada push se `site/` está em dia — se esquecer de rodar o
comando acima depois de mudar um projeto, o badge fica vermelho avisando.

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
