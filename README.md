# Analisador Contabil

Conferencia automatica de saldos no razao SCI.

## O que o app faz

O sistema le o plano de contas e o razao contabil exportado do SCI em CSV separado por ponto e virgula. Depois identifica contas e participantes com saldo contrario ao comportamento esperado, resume sequencias repetidas e mostra somente os casos que precisam de revisao.

## Como rodar no Streamlit Cloud

1. Envie este projeto para um repositorio no GitHub.
2. No Streamlit Cloud, crie um novo app a partir desse repositorio.
3. Em **Main file path**, use:

```text
streamlit_app.py
```

4. O Streamlit instalara as dependencias do arquivo `requirements.txt`.

## Arquivos principais

```text
streamlit_app.py              App para hospedagem no Streamlit
core.py                       Logica de leitura, analise e exportacao
requirements.txt              Dependencias do projeto
logo_analisador_contabil.svg  Logo do app
app.py                        Servidor local usado para testes
```

## Arquivos que nao devem ir para o GitHub

Arquivos reais de empresas, planilhas exportadas, resultados e logs locais ficam fora do GitHub pelo `.gitignore`.

Exemplos:

```text
*.csv
*.xlsx
*.xls
*.txt
*.log
```

Assim, plano de contas, razao contabil e exports gerados pelo app nao sao enviados ao repositorio.

## Rodar localmente

Para testar a versao local:

```text
abrir_analisador.bat
```

Para testar a versao Streamlit localmente:

```text
abrir_streamlit.bat
```
