# Consulta MDU Óptica — versão web automática

Aplicação de consulta somente leitura para o FORECAST MDU.

## O que a consulta mostra

Somente os 6 campos definidos:

1. Cidade
2. Endereço
3. Node
4. EPO MDU responsável
5. Construção RI MDU – fim
6. Construção RI MDU – status

## Arquitetura recomendada

**GitHub** guarda apenas o código do site.  
**Streamlit Cloud** publica o site com um link fixo.  
**Google Drive / OneDrive** guarda a planilha grande do FORECAST.

Fluxo:

`FORECAST.xlsx -> Drive/OneDrive -> Consulta MDU -> usuários`

A aplicação verifica a fonte a cada 5 minutos. Quando a planilha é atualizada mantendo o mesmo link, a consulta passa a utilizar a nova base sem trocar o endereço do site.

> Não é recomendado versionar diariamente o XLSX de ~50 MB dentro do GitHub. O repositório cresceria muito rápido. Por isso a planilha fica fora do GitHub e somente o código fica no repositório.

## Base inicial

O projeto já contém `data/base_mdu.json` com a base da versão aprovada de 11/09/2026. Ela funciona como contingência caso a fonte automática esteja temporariamente indisponível.

## Publicação rápida

1. Crie um repositório no GitHub, por exemplo `consulta-mdu-optica`.
2. Envie todos os arquivos desta pasta para o repositório.
3. Acesse https://share.streamlit.io/ e crie um app a partir desse repositório.
4. Arquivo principal: `app.py`.
5. Em **App settings > Secrets**, configure:

```toml
MDU_SOURCE_URL = "LINK_DA_PLANILHA"
APP_PASSWORD = "SENHA_DO_SITE"
```

6. Salve. O app reinicia e passa a ler a planilha automaticamente.

## Como atualizar no dia a dia

O responsável não mexe no GitHub e não muda o site.

Atualize/substitua a planilha no Google Drive/OneDrive **mantendo o mesmo arquivo/link**. Em até 5 minutos a aplicação verifica novamente a fonte. Há também um botão **Verificar atualização** na própria tela.

## Planilha esperada

A aba deve continuar se chamando:

`MDU'S BROWNFIELD(CONSULTA)`

A aplicação localiza os campos pelo nome, portanto a posição das colunas pode mudar. Os seis cabeçalhos esperados são:

- `DSC_CIDADE`
- `ENDEREÇO MDU`
- `NODE`
- `EPO MDU Responsável`
- `Construção RI MDU Fim`
- `Construção RI MDU Status`

Acentos, espaços e quebras de linha nos cabeçalhos são normalizados automaticamente.

## Segurança

Se `APP_PASSWORD` estiver definido nos Secrets, o site exige senha antes da consulta. O XLSX não fica no repositório.

Para dados corporativos sensíveis, confirme as regras internas antes de publicar em um serviço externo. Se necessário, a mesma aplicação pode ser hospedada em infraestrutura interna.
