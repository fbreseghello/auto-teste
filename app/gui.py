from __future__ import annotations

import threading
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app import __version__
from app.config import load_clients_config
from app.connectors.yampi import YampiClient
from app.database import connect, init_db, upsert_client
from app.services import (
    export_monthly_sheet_csv,
    export_orders_csv,
    reprocess_orders_for_period,
    sync_yampi_orders,
)
from app.updater import apply_update_from_github, check_for_updates


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


def _resolve_window(start_date: str, end_date: str) -> tuple[str, str]:
    start = _normalize_date(start_date)
    end = _normalize_date(end_date)
    if start and end and start > end:
        raise ValueError("Data inicio nao pode ser maior que data fim.")
    return start, end


def _expand_to_month_bounds(start_date: str, end_date: str) -> tuple[str, str]:
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


class AppGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Coletor de Dados - Analistas")
        self.root.geometry("860x620")
        self.root.minsize(820, 580)

        self.clients = load_clients_config()
        self.by_platform: dict[str, list] = {}
        for client in self.clients.values():
            self.by_platform.setdefault(client.platform, []).append(client)
        for platform in self.by_platform:
            self.by_platform[platform].sort(key=lambda c: (c.company.lower(), c.branch.lower(), c.id.lower()))

        self.platform_var = tk.StringVar()
        self.company_var = tk.StringVar()
        self.client_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.db_path_var = tk.StringVar(value="data/local.db")
        self.output_var = tk.StringVar()
        self._client_lookup: dict[str, object] = {}
        self._busy = False

        self._build_ui()
        self._load_platforms()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text=f"Extracao Yampi (1:1 Google Sheets)  {__version__}",
            font=("Segoe UI", 14, "bold"),
        )
        title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 14))

        ttk.Label(container, text="Plataforma").grid(row=1, column=0, sticky="w")
        self.platform_combo = ttk.Combobox(container, textvariable=self.platform_var, state="readonly", width=24)
        self.platform_combo.grid(row=1, column=1, sticky="ew", padx=(6, 18))
        self.platform_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_platform_change())

        ttk.Label(container, text="Empresa").grid(row=1, column=2, sticky="w")
        self.company_combo = ttk.Combobox(container, textvariable=self.company_var, state="readonly", width=24)
        self.company_combo.grid(row=1, column=3, sticky="ew", padx=(6, 0))
        self.company_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_company_change())

        ttk.Label(container, text="Alias/Filial").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.client_combo = ttk.Combobox(container, textvariable=self.client_var, state="readonly")
        self.client_combo.grid(row=2, column=1, columnspan=3, sticky="ew", padx=(6, 0), pady=(10, 0))
        self.client_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_client_change())

        ttk.Label(container, text="Data inicio").grid(row=3, column=0, sticky="w", pady=(14, 0))
        ttk.Entry(container, textvariable=self.start_date_var).grid(row=3, column=1, sticky="ew", padx=(6, 18), pady=(14, 0))
        ttk.Label(container, text="Data fim").grid(row=3, column=2, sticky="w", pady=(14, 0))
        ttk.Entry(container, textvariable=self.end_date_var).grid(row=3, column=3, sticky="ew", padx=(6, 0), pady=(14, 0))

        ttk.Label(container, text="Banco local").grid(row=4, column=0, sticky="w", pady=(14, 0))
        ttk.Entry(container, textvariable=self.db_path_var, state="readonly").grid(row=4, column=1, columnspan=2, sticky="ew", padx=(6, 10), pady=(14, 0))
        ttk.Button(container, text="Escolher", command=self._pick_db_path).grid(row=4, column=3, sticky="e", pady=(14, 0))

        ttk.Label(container, text="CSV destino").grid(row=5, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(container, textvariable=self.output_var, state="readonly").grid(row=5, column=1, columnspan=2, sticky="ew", padx=(6, 10), pady=(10, 0))
        ttk.Button(container, text="Salvar como", command=self._pick_output_path).grid(row=5, column=3, sticky="e", pady=(10, 0))

        actions = ttk.Frame(container)
        actions.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(16, 6))
        self.test_button = ttk.Button(actions, text="Testar Conexao API", command=self._test_connection_clicked)
        self.test_button.pack(side="left")
        self.sync_button = ttk.Button(actions, text="Sincronizar Pedidos", command=self._sync_clicked)
        self.sync_button.pack(side="left", padx=8)
        self.export_monthly_button = ttk.Button(
            actions, text="Exportar Mensal (Sheets)", command=self._export_monthly_clicked
        )
        self.export_monthly_button.pack(side="left")
        self.reprocess_button = ttk.Button(actions, text="Reprocessar Mes", command=self._reprocess_month_clicked)
        self.reprocess_button.pack(side="left", padx=8)
        self.export_orders_button = ttk.Button(actions, text="Exportar Pedidos", command=self._export_orders_clicked)
        self.export_orders_button.pack(side="left")
        self.update_button = ttk.Button(actions, text="Atualizar App", command=self._update_app_clicked)
        self.update_button.pack(side="left", padx=8)

        ttk.Label(container, text="Log").grid(row=7, column=0, sticky="w", pady=(10, 4))
        self.log_text = tk.Text(container, height=18, wrap="word")
        self.log_text.grid(row=8, column=0, columnspan=4, sticky="nsew")
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=8, column=4, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        container.columnconfigure(1, weight=1)
        container.columnconfigure(3, weight=1)
        container.rowconfigure(8, weight=1)

    def _log(self, message: str) -> None:
        stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.log_text.insert("end", f"[{stamp}] {message}\n")
        self.log_text.see("end")

    def _load_platforms(self) -> None:
        platforms = sorted(self.by_platform.keys())
        self.platform_combo["values"] = platforms
        if platforms:
            self.platform_var.set(platforms[0])
            self._on_platform_change()

    def _on_platform_change(self) -> None:
        platform = self.platform_var.get().strip()
        clients = self.by_platform.get(platform, [])
        companies = sorted({c.company for c in clients})
        self.company_combo["values"] = companies
        self.company_var.set(companies[0] if companies else "")
        self._on_company_change()

    def _on_company_change(self) -> None:
        platform = self.platform_var.get().strip()
        company = self.company_var.get().strip()
        clients = [c for c in self.by_platform.get(platform, []) if c.company == company]
        labels = [f"{c.branch} [alias={c.alias}] ({c.id})" for c in clients]
        self._client_lookup = dict(zip(labels, clients))
        self.client_combo["values"] = labels
        if labels:
            self.client_var.set(labels[0])
            self._on_client_change()
        else:
            self.client_var.set("")
            self.output_var.set("")

    def _on_client_change(self) -> None:
        client = self._selected_client()
        if not client:
            return
        self.output_var.set(f"exports/{client.id}_mensal.csv")
        self._apply_default_dates()

    def _apply_default_dates(self) -> None:
        today = datetime.now()
        first_day = today.replace(day=1)
        self.start_date_var.set(first_day.strftime("%d/%m/%Y"))
        self.end_date_var.set(today.strftime("%d/%m/%Y"))

    def _selected_client(self):
        return self._client_lookup.get(self.client_var.get().strip())

    def _pick_db_path(self) -> None:
        chosen = filedialog.asksaveasfilename(
            title="Escolha o banco SQLite",
            defaultextension=".db",
            initialfile=Path(self.db_path_var.get()).name or "local.db",
            filetypes=[("SQLite DB", "*.db"), ("Todos", "*.*")],
        )
        if chosen:
            self.db_path_var.set(chosen)

    def _pick_output_path(self) -> None:
        chosen = filedialog.asksaveasfilename(
            title="Salvar CSV",
            defaultextension=".csv",
            initialfile=Path(self.output_var.get()).name or "dados.csv",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
        )
        if chosen:
            self.output_var.set(chosen)

    def _run_background(self, fn) -> None:
        self._set_busy(True)

        def wrapper():
            try:
                fn()
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda: self._log(f"Erro: {exc}"))
                self.root.after(0, lambda: messagebox.showerror("Erro", str(exc)))
            finally:
                self.root.after(0, lambda: self._set_busy(False))

        threading.Thread(target=wrapper, daemon=True).start()

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.test_button.configure(state=state)
        self.sync_button.configure(state=state)
        self.export_monthly_button.configure(state=state)
        self.reprocess_button.configure(state=state)
        self.export_orders_button.configure(state=state)
        self.update_button.configure(state=state)
        self.platform_combo.configure(state="disabled" if busy else "readonly")
        self.company_combo.configure(state="disabled" if busy else "readonly")
        self.client_combo.configure(state="disabled" if busy else "readonly")

    def _test_connection_clicked(self) -> None:
        client = self._selected_client()
        if not client:
            messagebox.showwarning("Selecao", "Escolha um alias/filial.")
            return
        if client.platform != "yampi":
            messagebox.showwarning("Plataforma", "Somente Yampi implementado neste MVP.")
            return
        if not _has_yampi_auth(client):
            messagebox.showerror("Credencial", "Credenciais ausentes para este alias.")
            return

        def task():
            yampi = YampiClient(
                base_url=client.base_url,
                token=client.token,
                user_token=client.user_token,
                user_secret_key=client.user_secret_key,
            )
            self.root.after(0, lambda: self._log(f"Testando conexao API para {client.id}..."))
            ok, msg = yampi.test_connection(client.alias)
            if ok:
                self.root.after(0, lambda: self._log(f"Conexao OK: {msg}"))
                self.root.after(0, lambda: messagebox.showinfo("API", msg))
            else:
                self.root.after(0, lambda: self._log(f"Falha na conexao: {msg}"))
                self.root.after(0, lambda: messagebox.showerror("API", msg))

        self._run_background(task)

    def _sync_clicked(self) -> None:
        client = self._selected_client()
        if not client:
            messagebox.showwarning("Selecao", "Escolha um alias/filial.")
            return
        if client.platform != "yampi":
            messagebox.showwarning("Plataforma", "Somente Yampi implementado neste MVP.")
            return
        if not _has_yampi_auth(client):
            messagebox.showerror("Credencial", "Credenciais ausentes para este alias.")
            return

        try:
            start_date, end_date = _resolve_window(self.start_date_var.get(), self.end_date_var.get())
        except ValueError as exc:
            messagebox.showerror("Data invalida", str(exc))
            return
        if not start_date or not end_date:
            messagebox.showwarning("Periodo", "Informe data inicio e fim para exportar mensal com consistencia.")
            return
        start_date, end_date = _expand_to_month_bounds(start_date, end_date)

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            upsert_client(
                conn=conn,
                client_id=client.id,
                company=client.company,
                branch=client.branch,
                alias=client.alias,
                name=client.name,
                platform=client.platform,
            )
            conn.commit()

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            self.root.after(0, lambda: self._log(f"Sincronizando {client.id}..."))
            processed = sync_yampi_orders(
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
            conn.close()
            self.root.after(0, lambda: self._log(f"Sincronizacao concluida: {processed} pedidos processados."))
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", f"Sincronizacao concluida: {processed} pedidos."))

        self._run_background(task)

    def _export_monthly_clicked(self) -> None:
        client = self._selected_client()
        if not client:
            messagebox.showwarning("Selecao", "Escolha um alias/filial.")
            return
        if client.platform != "yampi":
            messagebox.showwarning("Plataforma", "Somente Yampi implementado neste MVP.")
            return
        if not _has_yampi_auth(client):
            messagebox.showerror("Credencial", "Credenciais ausentes para este alias.")
            return

        output = self.output_var.get().strip()
        if not output:
            messagebox.showwarning("Arquivo", "Informe o caminho do CSV.")
            return
        if not output.lower().endswith(".csv"):
            messagebox.showwarning("Arquivo", "O arquivo de saida deve terminar com .csv.")
            return

        try:
            start_date, end_date = _resolve_window(self.start_date_var.get(), self.end_date_var.get())
        except ValueError as exc:
            messagebox.showerror("Data invalida", str(exc))
            return
        if not start_date or not end_date:
            messagebox.showwarning("Periodo", "Informe data inicio e fim para exportar mensal com consistencia.")
            return
        start_date, end_date = _expand_to_month_bounds(start_date, end_date)

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            upsert_client(
                conn=conn,
                client_id=client.id,
                company=client.company,
                branch=client.branch,
                alias=client.alias,
                name=client.name,
                platform=client.platform,
            )
            conn.commit()

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            self.root.after(
                0,
                lambda: self._log(
                    f"Reprocessando periodo mensal {start_date} ate {end_date} antes da exportacao..."
                ),
            )
            deleted, synced = reprocess_orders_for_period(
                conn=conn,
                client_id=client.id,
                base_url=client.base_url,
                alias=client.alias,
                start_date=start_date,
                end_date=end_date,
                token=client.token,
                user_token=client.user_token,
                user_secret_key=client.user_secret_key,
                page_size=client.page_size,
            )
            self.root.after(
                0,
                lambda: self._log(f"Reprocessamento concluido. Removidos: {deleted}. Baixados: {synced}."),
            )
            self.root.after(0, lambda: self._log(f"Gerando CSV mensal: {output}"))
            count = export_monthly_sheet_csv(conn, client.id, output, start_date=start_date, end_date=end_date)
            conn.close()
            self.root.after(0, lambda: self._log(f"CSV mensal gerado com {count} linha(s)."))
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", f"CSV mensal gerado: {count} linha(s)."))

        self._run_background(task)

    def _reprocess_month_clicked(self) -> None:
        client = self._selected_client()
        if not client:
            messagebox.showwarning("Selecao", "Escolha um alias/filial.")
            return
        if client.platform != "yampi":
            messagebox.showwarning("Plataforma", "Somente Yampi implementado neste MVP.")
            return
        if not _has_yampi_auth(client):
            messagebox.showerror("Credencial", "Credenciais ausentes para este alias.")
            return

        try:
            start_date, end_date = _resolve_window(self.start_date_var.get(), self.end_date_var.get())
        except ValueError as exc:
            messagebox.showerror("Data invalida", str(exc))
            return

        if not start_date or not end_date:
            messagebox.showwarning("Periodo", "Informe data inicio e fim para reprocessar.")
            return
        if start_date[:7] != end_date[:7]:
            messagebox.showwarning("Periodo", "Reprocessar Mes exige inicio e fim no mesmo mes.")
            return

        month_start = f"{start_date[:7]}-01"
        month_dt = datetime.strptime(month_start, "%Y-%m-%d")
        if month_dt.month == 12:
            next_month = month_dt.replace(year=month_dt.year + 1, month=1, day=1)
        else:
            next_month = month_dt.replace(month=month_dt.month + 1, day=1)
        month_end = (next_month - timedelta(days=1)).strftime("%Y-%m-%d")

        proceed = messagebox.askyesno(
            "Confirmar Reprocessamento",
            f"Apagar dados locais de {client.id} entre {month_start} e {month_end} e baixar novamente?",
        )
        if not proceed:
            return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            upsert_client(
                conn=conn,
                client_id=client.id,
                company=client.company,
                branch=client.branch,
                alias=client.alias,
                name=client.name,
                platform=client.platform,
            )
            conn.commit()

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            self.root.after(0, lambda: self._log(f"Reprocessando mes {month_start[:7]} para {client.id}..."))
            deleted, synced = reprocess_orders_for_period(
                conn=conn,
                client_id=client.id,
                base_url=client.base_url,
                alias=client.alias,
                start_date=month_start,
                end_date=month_end,
                token=client.token,
                user_token=client.user_token,
                user_secret_key=client.user_secret_key,
                page_size=client.page_size,
            )
            conn.close()

            self.root.after(
                0,
                lambda: self._log(
                    f"Reprocessamento concluido. Removidos: {deleted}. Baixados novamente: {synced}."
                ),
            )
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Sucesso",
                    f"Reprocessamento concluido.\nRemovidos: {deleted}\nBaixados: {synced}",
                ),
            )

        self._run_background(task)

    def _export_orders_clicked(self) -> None:
        client = self._selected_client()
        if not client:
            messagebox.showwarning("Selecao", "Escolha um alias/filial.")
            return

        output = self.output_var.get().strip()
        if not output:
            output = f"exports/{client.id}_orders.csv"
        if not output.lower().endswith(".csv"):
            messagebox.showwarning("Arquivo", "O arquivo de saida deve terminar com .csv.")
            return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            self.root.after(0, lambda: self._log(f"Gerando CSV detalhado: {output}"))
            count = export_orders_csv(conn, client.id, output)
            conn.close()
            self.root.after(0, lambda: self._log(f"CSV detalhado gerado com {count} linha(s)."))
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", f"CSV detalhado gerado: {count} linha(s)."))

        self._run_background(task)

    def _update_app_clicked(self) -> None:
        def task():
            self.root.after(0, lambda: self._log("Verificando atualizacao no GitHub..."))
            check = check_for_updates(current_version=__version__)
            if not check.has_update:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Atualizacao",
                        f"Aplicativo atualizado.\nVersao atual: {check.current_version}",
                    ),
                )
                self.root.after(0, lambda: self._log("Sem atualizacao disponivel."))
                return

            self.root.after(0, lambda: self._log(f"Baixando e aplicando {check.latest_version}..."))
            result = apply_update_from_github(
                current_version=__version__,
                repo=check.repo,
                project_dir=".",
            )
            self.root.after(0, lambda: self._log(result.message))
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Atualizacao concluida",
                    f"{result.message}\n\nFeche e abra o app novamente.",
                ),
            )

        self._run_background(task)


def main() -> None:
    root = tk.Tk()
    AppGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
