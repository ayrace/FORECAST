# Passo a passo — colocar a Consulta MDU no ar

## Etapa 1 — GitHub

1. Entre no GitHub.
2. Crie um novo repositório chamado `consulta-mdu-optica`.
3. Deixe o repositório **Private** se a política da empresa permitir o uso com Streamlit Cloud; caso contrário, valide o modelo de acesso com TI.
4. Faça upload de todos os arquivos deste pacote, preservando as pastas `.streamlit` e `data`.
5. Faça o commit.

## Etapa 2 — fonte da planilha

### Opção recomendada: Google Drive

1. Coloque o FORECAST em uma pasta controlada.
2. Use sempre o **mesmo arquivo**.
3. Compartilhe de forma que a aplicação consiga baixar o arquivo.
4. Copie o link desse arquivo.

Ao atualizar o FORECAST, substitua a versão do arquivo mantendo o mesmo ID/link.

### OneDrive/SharePoint

Também funciona desde que o endereço configurado permita download direto sem uma tela interativa de login. Links corporativos autenticados normalmente exigem uma integração adicional.

## Etapa 3 — Streamlit Cloud

1. Abra `share.streamlit.io`.
2. Clique em **Create app**.
3. Selecione o repositório `consulta-mdu-optica`.
4. Branch: `main`.
5. Main file path: `app.py`.
6. Deploy.

## Etapa 4 — configurar a atualização automática

No Streamlit, abra **Settings > Secrets** e cole:

```toml
MDU_SOURCE_URL = "COLE_O_LINK_DO_FORECAST"
APP_PASSWORD = "SUA_SENHA"
```

Salve.

A partir daí:

- o endereço do site não muda;
- os usuários recebem o link uma única vez;
- a planilha é verificada a cada 5 minutos;
- se a nova planilha falhar, a base inicial continua disponível;
- o botão **Verificar atualização** força uma nova leitura imediatamente.

## Rotina diária

**Responsável pela base:** atualiza o mesmo FORECAST no Drive/OneDrive.  
**Usuários:** apenas acessam o mesmo link da Consulta MDU.

Não é necessário reenviar HTML, gerar nova versão ou mandar um novo link.
