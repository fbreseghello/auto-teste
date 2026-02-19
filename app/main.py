from __future__ import annotations

import argparse
import sys
from datetime import datetime, timedelta

from app import __version__
from app.config import load_clients_config
from app.database import connect, init_db, upsert_client
from app.services import export_monthly_sheet_csv, export_order_skus_csv, export_orders_csv, sync_yampi_orders
from app.updater import apply_update_from_github, check_for_updates


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Coletor local de dados para analistas.")
    parser.add_argument("--db-path", default="data/local.db", help="Caminho do banco SQLite.")

    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("init-db", help="Cria as tabelas no banco local.")
    subparsers.add_parser("list-clients", help="Lista clientes/unidades configurados.")
    subparsers.add_parser("list-tree", help="Mostra a arvore plataforma -> clientes.")
    subparsers.add_parser("menu", help="Abre menu interativo para escolher plataforma e cliente.")

    sync = subparsers.add_parser("sync-yampi", help="Sincroniza pedidos da Yampi.")
    sync.add_argument("--client", required=True, help="ID do cliente (config/clients.json).")
    sync.add_argument("--start-date", default="", help="Data inicio (dd/mm/aaaa ou aaaa-mm-dd).")
    sync.add_argument("--end-date", default="", help="Data fim (dd/mm/aaaa ou aaaa-mm-dd).")

    export = subparsers.add_parser("export-orders", help="Exporta pedidos para CSV.")
    export.add_argument("--client", required=True, help="ID do cliente (config/clients.json).")
    export.add_argument("--output", required=True, help="Caminho do CSV de saida.")

    export_skus = subparsers.add_parser("export-skus", help="Exporta itens/SKUs dos pedidos para CSV.")
    export_skus.add_argument("--client", required=True, help="ID do cliente (config/clients.json).")
    export_skus.add_argument("--output", required=True, help="Caminho do CSV de saida.")
    export_skus.add_argument("--start-date", default="", help="Data inicio (dd/mm/aaaa ou aaaa-mm-dd).")
    export_skus.add_argument("--end-date", default="", help="Data fim (dd/mm/aaaa ou aaaa-mm-dd).")
    export_skus.add_argument("--order-number", default="", help="Numero do pedido (prioridade sobre periodo).")

    export_monthly = subparsers.add_parser("export-monthly", help="Exporta agregado mensal (1:1 Sheets).")
    export_monthly.add_argument("--client", required=True, help="ID do cliente (config/clients.json).")
    export_monthly.add_argument("--output", required=True, help="Caminho do CSV de saida.")
    export_monthly.add_argument("--start-date", default="", help="Data inicio (dd/mm/aaaa ou aaaa-mm-dd).")
    export_monthly.add_argument("--end-date", default="", help="Data fim (dd/mm/aaaa ou aaaa-mm-dd).")

    update_app = subparsers.add_parser("update-app", help="Atualiza o app pela ultima release do GitHub.")
    update_app.add_argument("--repo", default="", help="Repositorio GitHub (dono/repositorio).")
    update_app.add_argument("--force", action="store_true", help="Forca atualizacao mesmo sem nova versao.")
    update_app.add_argument(
        "--check-only",
        action="store_true",
        help="Somente verifica se existe versao nova, sem baixar arquivos.",
    )

    return parser.parse_args()


def _require_client(clients, client_id):
    if client_id not in clients:
        available = ", ".join(sorted(clients.keys())) or "(nenhum)"
        raise ValueError(f"Cliente '{client_id}' nao encontrado. Disponiveis: {available}")
    return clients[client_id]


def _group_clients_by_platform(clients):
    grouped = {}
    for client in clients.values():
        grouped.setdefault(client.platform, []).append(client)
    for platform in grouped:
        grouped[platform].sort(key=lambda c: (c.company.lower(), c.branch.lower(), c.id.lower()))
    return grouped


def _group_clients_by_company(platform_clients):
    grouped = {}
    for client in platform_clients:
        grouped.setdefault(client.company, []).append(client)
    for company in grouped:
        grouped[company].sort(key=lambda c: (c.branch.lower(), c.id.lower()))
    return grouped


def _normalize_date(date_str: str) -> str:
    value = date_str.strip()
    if not value:
        return ""
    for pattern in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, pattern).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError(f"Data invalida '{date_str}'. Use dd/mm/aaaa ou aaaa-mm-dd.")


def _resolve_sync_window(start_date: str, end_date: str) -> tuple[str, str]:
    start = _normalize_date(start_date)
    end = _normalize_date(end_date)
    if start and end and start > end:
        raise ValueError("Data inicio nao pode ser maior que data fim.")
    return start, end


def _expand_to_month_bounds(start_date: str, end_date: str) -> tuple[str, str]:
    if not start_date or not end_date:
        return start_date, end_date

    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d")

    start_month = start_dt.replace(day=1)
    if end_dt.month == 12:
        next_month = end_dt.replace(year=end_dt.year + 1, month=1, day=1)
    else:
        next_month = end_dt.replace(month=end_dt.month + 1, day=1)
    end_month = next_month - timedelta(days=1)

    today = datetime.now()
    if end_month.year == today.year and end_month.month == today.month and end_month > today:
        end_month = today

    return start_month.strftime("%Y-%m-%d"), end_month.strftime("%Y-%m-%d")


def _has_yampi_auth(client) -> bool:
    if client.token:
        return True
    return bool(client.user_token and client.user_secret_key)


def _auth_hint(client) -> str:
    hints = []
    if client.token_env:
        hints.append(f"token_env='{client.token_env}'")
    if client.user_token_env:
        hints.append(f"user_token_env='{client.user_token_env}'")
    if client.user_secret_key_env:
        hints.append(f"user_secret_key_env='{client.user_secret_key_env}'")
    return ", ".join(hints) if hints else "credenciais nao configuradas no clients.json"


def _pick_option(title, options, formatter):
    print()
    print(title)
    for idx, option in enumerate(options, start=1):
        print(f"{idx}. {formatter(option)}")
    print("0. Voltar/Sair")

    while True:
        raw = input("Escolha: ").strip()
        if raw == "0":
            return None
        if raw.isdigit():
            pos = int(raw)
            if 1 <= pos <= len(options):
                return options[pos - 1]
        print("Opcao invalida. Tente novamente.")


def _run_interactive_menu(conn, clients) -> int:
    init_db(conn)
    grouped = _group_clients_by_platform(clients)
    platforms = sorted(grouped.keys())
    if not platforms:
        print("Nenhum cliente configurado em config/clients.json.")
        return 0

    while True:
        platform = _pick_option("Plataformas", platforms, lambda p: p)
        if platform is None:
            return 0

        while True:
            platform_clients = grouped[platform]
            companies = sorted(_group_clients_by_company(platform_clients).keys())
            company = _pick_option(f"Empresas de '{platform}'", companies, lambda c: c)
            if company is None:
                break
            company_clients = _group_clients_by_company(platform_clients)[company]

            client = _pick_option(
                f"Alias/Filiais de '{company}'",
                company_clients,
                lambda c: f"{c.branch} [alias={c.alias}] ({c.id})",
            )
            if client is None:
                continue

            if client.platform == "yampi":
                action = _pick_option(
                    f"Acoes para '{client.id}'",
                    ["sync", "export_monthly", "export_orders", "export_skus"],
                    lambda a: (
                        "Sincronizar pedidos"
                        if a == "sync"
                        else (
                            "Exportar agregado mensal (Sheets)"
                            if a == "export_monthly"
                            else ("Exportar pedidos detalhados" if a == "export_orders" else "Exportar itens/SKUs")
                        )
                    ),
                )
                if action is None:
                    continue

                if action == "sync":
                    if not _has_yampi_auth(client):
                        print(f"Credenciais ausentes para '{client.id}'. Defina {_auth_hint(client)}.")
                        continue
                    start_date_input = input("Data inicio (dd/mm/aaaa) [vazio=incremental]: ").strip()
                    end_date_input = input("Data fim (dd/mm/aaaa) [vazio=incremental]: ").strip()
                    try:
                        start_date, end_date = _resolve_sync_window(start_date_input, end_date_input)
                    except ValueError as exc:
                        print(f"Erro: {exc}")
                        continue
                    upsert_client(
                        conn,
                        client_id=client.id,
                        company=client.company,
                        branch=client.branch,
                        alias=client.alias,
                        name=client.name,
                        platform=client.platform,
                    )
                    conn.commit()
                    synced = sync_yampi_orders(
                        conn=conn,
                        client_id=client.id,
                        base_url=client.base_url,
                        alias=client.alias,
                        token=client.token,
                        user_token=client.user_token,
                        user_secret_key=client.user_secret_key,
                        page_size=client.page_size,
                        start_date=start_date,
                        end_date=end_date,
                    )
                    print(f"Sincronizacao concluida para '{client.id}': {synced} pedidos processados.")
                    continue

                if action == "export_orders":
                    default_output = f"exports/{client.id}_orders.csv"
                    output = input(f"Caminho do CSV [{default_output}]: ").strip() or default_output
                    count = export_orders_csv(conn, client.id, output)
                    print(f"Arquivo gerado: {output} ({count} linhas)")
                    continue

                if action == "export_skus":
                    order_number = input("Numero do pedido [vazio=filtrar por periodo]: ").strip()
                    start_date = ""
                    end_date = ""
                    if not order_number:
                        start_date_input = input("Data inicio (dd/mm/aaaa): ").strip()
                        end_date_input = input("Data fim (dd/mm/aaaa): ").strip()
                        try:
                            start_date, end_date = _resolve_sync_window(start_date_input, end_date_input)
                        except ValueError as exc:
                            print(f"Erro: {exc}")
                            continue
                        if not start_date or not end_date:
                            print("Erro: informe data inicio e fim, ou numero do pedido.")
                            continue
                    default_output = f"exports/{client.id}_skus.csv"
                    output = input(f"Caminho do CSV [{default_output}]: ").strip() or default_output
                    count = export_order_skus_csv(
                        conn,
                        client.id,
                        output,
                        start_date=start_date,
                        end_date=end_date,
                        order_number=order_number,
                    )
                    print(f"Arquivo gerado: {output} ({count} linhas)")
                    continue

                start_date_input = input("Data inicio (dd/mm/aaaa) [vazio=tudo]: ").strip()
                end_date_input = input("Data fim (dd/mm/aaaa) [vazio=tudo]: ").strip()
                try:
                    start_date, end_date = _resolve_sync_window(start_date_input, end_date_input)
                except ValueError as exc:
                    print(f"Erro: {exc}")
                    continue
                start_date, end_date = _expand_to_month_bounds(start_date, end_date)
                default_output = f"exports/{client.id}_mensal.csv"
                output = input(f"Caminho do CSV [{default_output}]: ").strip() or default_output
                count = export_monthly_sheet_csv(
                    conn, client.id, output, start_date=start_date, end_date=end_date
                )
                print(f"Arquivo gerado: {output} ({count} linhas)")
                continue

            print(f"Plataforma '{client.platform}' ainda nao implementada para sync/export neste MVP.")


def run() -> int:
    args = parse_args()

    if args.command == "update-app":
        if args.check_only:
            check = check_for_updates(current_version=__version__, repo=args.repo)
            if check.has_update:
                print(
                    f"Atualizacao disponivel: {check.current_version} -> {check.latest_version} "
                    f"({check.release_url})"
                )
            else:
                print(f"Sem atualizacao. Versao atual: {check.current_version}.")
            return 0

        result = apply_update_from_github(
            current_version=__version__,
            repo=args.repo,
            force=args.force,
            project_dir=".",
        )
        print(result.message)
        return 0

    conn = connect(args.db_path)

    if args.command == "init-db":
        init_db(conn)
        print(f"Banco inicializado em: {args.db_path}")
        return 0

    clients = load_clients_config()
    init_db(conn)

    if args.command == "list-clients":
        for client in sorted(clients.values(), key=lambda c: c.id):
            token_status = "OK" if _has_yampi_auth(client) else "SEM_CREDENCIAL"
            auth_type = "bearer_token" if client.token else "user_token_secret"
            print(
                f"{client.id} | {client.company} | {client.branch} | "
                f"{client.platform} | alias={client.alias} | {auth_type} | {token_status}"
            )
        return 0

    if args.command == "list-tree":
        grouped = _group_clients_by_platform(clients)
        for platform in sorted(grouped.keys()):
            print(platform)
            by_company = _group_clients_by_company(grouped[platform])
            for company in sorted(by_company.keys()):
                print(f"  {company}")
                for client in by_company[company]:
                    print(f"    - {client.branch} [alias={client.alias}] ({client.id})")
        return 0

    if args.command == "menu":
        return _run_interactive_menu(conn, clients)

    if args.command == "sync-yampi":
        client = _require_client(clients, args.client)
        if client.platform != "yampi":
            raise ValueError(
                f"Cliente '{client.id}' nao esta com platform='yampi' (atual: {client.platform})."
            )
        if not _has_yampi_auth(client):
            raise ValueError(
                f"Credenciais nao definidas para '{client.id}'. "
                f"Defina {_auth_hint(client)}."
            )
        start_date, end_date = _resolve_sync_window(args.start_date, args.end_date)
        upsert_client(
            conn,
            client_id=client.id,
            company=client.company,
            branch=client.branch,
            alias=client.alias,
            name=client.name,
            platform=client.platform,
        )
        conn.commit()

        synced = sync_yampi_orders(
            conn=conn,
            client_id=client.id,
            base_url=client.base_url,
            alias=client.alias,
            token=client.token,
            user_token=client.user_token,
            user_secret_key=client.user_secret_key,
            page_size=client.page_size,
            start_date=start_date,
            end_date=end_date,
        )
        print(f"Sincronizacao concluida para '{client.id}': {synced} pedidos processados.")
        return 0

    if args.command == "export-orders":
        _require_client(clients, args.client)
        count = export_orders_csv(conn, args.client, args.output)
        print(f"Arquivo gerado: {args.output} ({count} linhas)")
        return 0

    if args.command == "export-skus":
        _require_client(clients, args.client)
        order_number = args.order_number.strip()
        start_date = ""
        end_date = ""
        if not order_number:
            start_date, end_date = _resolve_sync_window(args.start_date, args.end_date)
            if not start_date or not end_date:
                raise ValueError("Informe --order-number ou o periodo completo com --start-date e --end-date.")
        count = export_order_skus_csv(
            conn,
            args.client,
            args.output,
            start_date=start_date,
            end_date=end_date,
            order_number=order_number,
        )
        print(f"Arquivo gerado: {args.output} ({count} linhas)")
        return 0

    if args.command == "export-monthly":
        _require_client(clients, args.client)
        start_date, end_date = _resolve_sync_window(args.start_date, args.end_date)
        start_date, end_date = _expand_to_month_bounds(start_date, end_date)
        count = export_monthly_sheet_csv(
            conn, args.client, args.output, start_date=start_date, end_date=end_date
        )
        print(f"Arquivo gerado: {args.output} ({count} linhas)")
        return 0

    raise ValueError(f"Comando nao suportado: {args.command}")


if __name__ == "__main__":
    try:
        sys.exit(run())
    except Exception as exc:  # noqa: BLE001
        print(f"Erro: {exc}", file=sys.stderr)
        sys.exit(1)
