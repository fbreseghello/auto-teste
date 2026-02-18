# Auto Teste - Coletor Local de Dados

MVP em Python para substituir fluxos de Google Apps Script no tratamento de dados.

## Objetivo

Sincronizar dados da API da Yampi para um banco local (SQLite) por analista, com exportacao para CSV.

Fluxo:

1. API Yampi
2. Banco local SQLite
3. Exportacao CSV
4. Importacao no Google Sheets

## Requisitos

- Python 3.11+
- `pip install -r requirements.txt`

## Configuracao

1. Copie o arquivo de exemplo:

```powershell
Copy-Item config\clients.example.json config\clients.json
```

2. Ajuste o `config/clients.json` com suas unidades (empresa + filial/divisao).

3. Defina as chaves de API no ambiente (ou em `.env` local):

```env
YAMPI_EMPRESA_FILIAL_USER_TOKEN=seu_user_token
YAMPI_EMPRESA_FILIAL_USER_SECRET_KEY=sua_secret_key
YAMPI_EMPRESA_FILIAL_TOKEN=seu_token
```

Opcao simples para teste sem comandos:

1. Abra o arquivo `.env` e preencha as chaves.
2. Execute `executar_teste_aurha.bat` com duplo clique.

Opcao recomendada para analistas (sem terminal):

1. Abra `.env` e preencha as chaves.
2. Execute `abrir_app.bat` com duplo clique.
3. Na tela:
   - escolha `plataforma -> empresa -> alias`
   - use `Testar Conexao API` para validar credenciais
   - periodo ja vem preenchido automaticamente (1o dia do mes ate hoje)
   - clique em `Sincronizar Pedidos`
   - opcional: clique em `Reprocessar Mes` para apagar e baixar novamente o mes completo
   - clique em `Exportar Mensal (Sheets)` para gerar o CSV final
     (este botao agora reprocessa o periodo selecionado antes de exportar)
     e interpreta datas como periodo mensal (inicio no dia 01 e fim no ultimo dia/hoje)

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
- Credenciais ficam em variaveis de ambiente (`user_token_env`, `user_secret_key_env` e/ou `token_env`).
- Mantenha o `.env` local fora de repositorio.
