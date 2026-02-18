# Auto Teste - Coletor Local de Dados

MVP em Python para substituir fluxos de Google Apps Script no tratamento de dados.

## Objetivo

Sincronizar dados da API da Yampi para um banco local (SQLite) por analista, com exportacao para CSV.

Fluxo:

1. API Yampi
2. Banco local SQLite
3. Exportacao CSV
4. Importacao no Google Sheets

## Requisitos (Desenvolvimento)

- Python 3.11+
- `pip install -r requirements.txt`

## Configuracao

Fluxo plug and play para analistas (sem editar arquivos):

1. Baixe e extraia o `auto-teste-windows.zip` da release.
2. (Opcional, recomendado) copie a pasta `clientes` para dentro da pasta do app.
3. Execute `INICIAR_AUTO_TESTE.bat`.
4. Se necessario, no app clique em `1) Credenciais` e salve.
5. Clique em `Testar Conexao API`.
6. Clique em `2) Exportar Mensal`.

Observacoes:
- `config/clients.json` e `.env` sao criados automaticamente quando faltam.
- se existir `clientes/clients.json` e `clientes/.env`, o app usa esses arquivos automaticamente.
- no pacote com `.exe`, nao precisa Python instalado.
- o CSV mensal vai por padrao para `Downloads`.
- ao finalizar exportacao, a pasta do arquivo e aberta automaticamente.

Controles anti-erro na interface:
- botoes ficam bloqueados durante execucao em background
- seletores de cliente ficam bloqueados durante execucao
- `Banco local` e `CSV destino` ficam em modo somente leitura (usar botoes `Escolher` e `Salvar como`)
- validacao de formato `.csv` na exportacao
- `Reprocessar Mes` exige inicio e fim no mesmo mes e evita dado parcial antigo no banco local

## Comandos

Inicializar banco:

```powershell
python -m app.main init-db
```

Verificar/atualizar app pela ultima release do GitHub:

```powershell
python -m app.main update-app --check-only
python -m app.main update-app
```

O repositorio de update padrao ja vem configurado para `fbreseghello/auto-teste`.
Use `AUTO_TESTE_GITHUB_REPO` apenas se quiser sobrescrever para outro repositorio.

Sincronizar pedidos da Yampi para um cliente:

```powershell
python -m app.main sync-yampi --client unfair_pinkperfect
```

Sincronizar por janela de datas:

```powershell
python -m app.main sync-yampi --client unfair_pinkperfect --start-date 31/01/2026 --end-date 13/02/2026
```

Exportar agregado mensal (1:1 do Google Sheets):

```powershell
python -m app.main export-monthly --client unfair_pinkperfect --start-date 31/01/2026 --end-date 13/02/2026 --output exports\unfair_pinkperfect_mensal.csv
```

Colunas do CSV mensal:
- `Data` (YYYY-MM-01)
- `Nome Empresa`
- `Alias`
- `Vendas de Produto` = `value_products + value_shipment`
- `Descontos concedidos` = `value_discount`
- `Juros de Venda` = `value_tax`

Filtros aplicados no agregado mensal:
- pedidos criados no periodo selecionado
- apenas IDs unicos (`client_id + order_id`)
- `payment_date` preenchida
- `cancelled_date` vazia

Exportar pedidos para CSV:

```powershell
python -m app.main export-orders --client unfair_pinkperfect --output exports\unfair_pinkperfect_orders.csv
```

Listar unidades configuradas e status de token:

```powershell
python -m app.main list-clients
```

Mostrar arvore plataforma -> clientes:

```powershell
python -m app.main list-tree
```

Abrir menu interativo (plataforma -> cliente -> acao):

```powershell
python -m app.main menu
```

## Estrutura

- `app/config.py`: leitura de configuracao de clientes
- `app/database.py`: schema e operacoes SQLite
- `app/connectors/yampi.py`: cliente HTTP da Yampi
- `app/services.py`: sincronizacao incremental e exportacao
- `app/main.py`: interface de linha de comando
- `app/updater.py`: atualizacao automatica via GitHub Releases

## Distribuicao Para Usuarios

Fluxo recomendado:

1. Publicar release com tag (`0.2.0`, `0.2.1`, etc) no GitHub.
2. O workflow `.github/workflows/release-windows.yml` gera o arquivo `auto-teste-windows.zip` na release.
3. Usuarios baixam esse ZIP, extraem e executam `INICIAR_AUTO_TESTE.bat`.
4. O pacote ja inclui `AutoTeste.exe` e nao exige Python instalado na maquina do analista.

## Exemplo de cliente/unidade

```json
{
  "id": "unfair_pinkperfect",
  "company": "Unfair",
  "branch": "pinkperfect",
  "alias": "pinkperfect",
  "name": "Unfair - pinkperfect",
  "platform": "yampi",
  "base_url": "https://api.dooki.com.br/v2",
  "user_token_env": "YAMPI_UNFAIR_PINKPERFECT_USER_TOKEN",
  "user_secret_key_env": "YAMPI_UNFAIR_PINKPERFECT_USER_SECRET_KEY",
  "token_env": "YAMPI_UNFAIR_PINKPERFECT_TOKEN",
  "page_size": 100
}
```

## Notas de seguranca

- Nao versionar `config/clients.json` com segredos.
- Nao versionar `clientes/.env` com segredos.
- Credenciais ficam em variaveis de ambiente (`user_token_env`, `user_secret_key_env` e/ou `token_env`).
- Mantenha o `.env` local fora de repositorio.
