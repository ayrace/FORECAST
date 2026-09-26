# Passo a passo — Consulta MDU

## 1. Google Drive
A pasta já configurada é:
`https://drive.google.com/drive/folders/1h8Ium1WG9ZmZeuONQuBtAGScw7ukie9V?usp=drive_link`

Ela foi validada com o arquivo:
`FORECAST CONSOLIDADO CONSTRUÇÃO MDU_11_09_26.xlsx`

Mantenha somente um `.xlsx` de produção nessa pasta.

## 2. GitHub
Crie um repositório, por exemplo `consulta-mdu-optica`, e envie **o conteúdo da pasta do projeto**, não o ZIP fechado.

Na raiz do repositório devem aparecer pelo menos:
- `app.py`
- `requirements.txt`
- `README.md`
- pasta `data`
- pasta `.streamlit`

## 3. Streamlit Community Cloud
Crie o app apontando para:
- repositório: o criado acima
- branch: `main`
- arquivo principal: `app.py`

## 4. Secrets
A fonte já está embutida como padrão no `app.py`, então `MDU_SOURCE_URL` é opcional.

Se quiser configurar explicitamente em **Settings > Secrets**:

```toml
MDU_SOURCE_URL = "https://drive.google.com/drive/folders/1h8Ium1WG9ZmZeuONQuBtAGScw7ukie9V?usp=drive_link"
```

Para proteger a consulta com senha, adicione também:

```toml
APP_PASSWORD = "SUA_SENHA"
```

## 5. Atualização diária
Quando receber o FORECAST mais recente:
1. entre na pasta `MDU-FORECAST`;
2. substitua o `.xlsx` anterior pelo novo;
3. deixe somente uma planilha de produção na pasta;
4. não altere GitHub nem Streamlit.

A aplicação verifica automaticamente a cada 5 minutos. Para testar imediatamente, abra o site e pressione **Verificar atualização**.

## 6. Como confirmar
No cabeçalho da consulta você verá:
- data identificada no nome do FORECAST;
- horário da última verificação;
- total de registros;
- total de cidades;
- nome/identificação da fonte.

Se a nova planilha estiver inválida, a última base de contingência permanece disponível e a página exibirá um aviso.
