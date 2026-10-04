# Consulta MDU — Drive V3

Aplicação Streamlit de consulta somente leitura alimentada pelo FORECAST MDU no Google Drive.

## Fonte configurada
Pasta: **MDU-FORECAST**  
ID: `1h8Ium1WG9ZmZeuONQuBtAGScw7ukie9V`

Arquivo detectado na validação de 26/09/2026:
`FORECAST CONSOLIDADO CONSTRUÇÃO MDU_11_09_26.xlsx`

A aplicação verifica a pasta a cada **5 minutos**. Também há o botão **Verificar atualização** para forçar uma nova leitura.

### Regra principal
Mantenha **somente um arquivo .xlsx de produção** dentro da pasta MDU-FORECAST. Quando houver uma base mais recente, substitua o arquivo anterior. O site continua no mesmo link e lê a nova planilha.

## Campos exibidos
- Cidade
- Endereço
- Node
- EPO MDU responsável
- Construção RI MDU – fim
- Construção RI MDU – status

## Contingência
Se o Google Drive estiver temporariamente indisponível ou a nova planilha tiver erro estrutural, a aplicação mantém a base embarcada de contingência em `data/base_mdu.json`.

## Compartilhamento do Drive
Para o modo simples com `gdown`, a pasta precisa permitir leitura pelo ambiente do Streamlit. Se a base não puder ser pública por link, migre para acesso autenticado via Google Drive API/service account.

Veja `PASSO_A_PASSO.md`.

## V9 — campos SINERGIA
A consulta também exibe os campos Brownfield HPs, Brownfield inst. GPON, GPON + Híbrido e Brownfield % Penetração GPON, lidos diretamente da mesma aba MDU'S BROWNFIELD(CONSULTA).
