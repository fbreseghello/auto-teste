from __future__ import annotations

import os
import threading
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app import __version__
from app.config import load_clients_config, resolve_runtime_paths, save_client_credentials
from app.connectors.yampi import YampiClient
from app.database import connect, init_db, upsert_client
from app.google_sheets import (
    resolve_monthly_sheets_settings,
    upload_csv_to_google_sheet,
)
from app.services import (
    export_monthly_sheet_csv,
    export_order_skus_csv,
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
        self.root.geometry("980x640")
        self.root.minsize(900, 580)

        self.clients = load_clients_config()
        self.by_platform: dict[str, list] = {}
        self._rebuild_client_index()

        self.platform_var = tk.StringVar()
        self.company_var = tk.StringVar()
        self.select_all_var = tk.BooleanVar(value=False)
        self.selection_info_var = tk.StringVar(value="Nenhum alias selecionado.")
        self.status_var = tk.StringVar(value="Pronto")
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.order_number_var = tk.StringVar()
        self.db_path_var = tk.StringVar(value="data/local.db")
        self.output_var = tk.StringVar()
        self._company_clients: list = []
        self._client_check_vars: dict[str, tk.BooleanVar] = {}
        self._client_checkbuttons: list[ttk.Checkbutton] = []
        self._client_canvas_window = None
        self._busy = False

        self._configure_styles()
        self._build_ui()
        self._load_platforms()
        self.root.after(0, self._log_runtime_sources)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Muted.TLabel", font=("Segoe UI", 9))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            header,
            text=f"Extracao Yampi  {__version__}",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=1, sticky="e")
        header.columnconfigure(0, weight=1)

        select_frame = ttk.LabelFrame(container, text="Clientes", style="Section.TLabelframe")
        select_frame.grid(row=1, column=0, sticky="nsew")
        left_panel = ttk.Frame(select_frame)
        left_panel.grid(row=0, column=0, sticky="nw", padx=(0, 12))
        right_panel = ttk.Frame(select_frame)
        right_panel.grid(row=0, column=1, sticky="nsew")

        ttk.Label(left_panel, text="Plataforma").grid(row=0, column=0, sticky="w")
        self.platform_combo = ttk.Combobox(left_panel, textvariable=self.platform_var, state="readonly", width=28)
        self.platform_combo.grid(row=1, column=0, sticky="ew", pady=(2, 8))
        self.platform_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_platform_change())

        ttk.Label(left_panel, text="Empresa").grid(row=2, column=0, sticky="w")
        self.company_combo = ttk.Combobox(left_panel, textvariable=self.company_var, state="readonly", width=28)
        self.company_combo.grid(row=3, column=0, sticky="ew", pady=(2, 8))
        self.company_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_company_change())

        ttk.Label(left_panel, textvariable=self.selection_info_var, style="Muted.TLabel").grid(
            row=4, column=0, sticky="w", pady=(2, 0)
        )

        ttk.Label(right_panel, text="Alias/Filiais").grid(row=0, column=0, sticky="w")
        self.select_all_check = ttk.Checkbutton(
            right_panel,
            text="Selecionar todos",
            variable=self.select_all_var,
            command=self._toggle_select_all_clients,
        )
        self.select_all_check.grid(row=0, column=1, sticky="e")

        alias_frame = ttk.Frame(right_panel)
        alias_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        self.client_canvas = tk.Canvas(alias_frame, height=130, highlightthickness=0)
        self.client_canvas.grid(row=0, column=0, sticky="ew")
        self.client_scrollbar = ttk.Scrollbar(alias_frame, orient="vertical", command=self.client_canvas.yview)
        self.client_scrollbar.grid(row=0, column=1, sticky="ns")
        self.client_canvas.configure(yscrollcommand=self.client_scrollbar.set)
        self.client_checks_frame = ttk.Frame(self.client_canvas)
        self._client_canvas_window = self.client_canvas.create_window((0, 0), window=self.client_checks_frame, anchor="nw")
        self.client_checks_frame.bind("<Configure>", self._on_client_checks_configure)
        self.client_canvas.bind("<Configure>", self._on_client_canvas_configure)

        left_panel.columnconfigure(0, weight=1)
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)
        select_frame.columnconfigure(1, weight=1)
        select_frame.rowconfigure(0, weight=1)
        alias_frame.columnconfigure(0, weight=1)

        config_frame = ttk.LabelFrame(container, text="Periodo e Arquivos", style="Section.TLabelframe")
        config_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(config_frame, text="Inicio").grid(row=0, column=0, sticky="w")
        self.start_date_entry = ttk.Entry(config_frame, textvariable=self.start_date_var, width=14)
        self.start_date_entry.grid(row=0, column=1, sticky="w", padx=(6, 12))
        ttk.Label(config_frame, text="Fim").grid(row=0, column=2, sticky="w")
        self.end_date_entry = ttk.Entry(config_frame, textvariable=self.end_date_var, width=14)
        self.end_date_entry.grid(row=0, column=3, sticky="w", padx=(6, 12))
        self.current_month_button = ttk.Button(config_frame, text="Mes atual", command=self._set_current_month_dates)
        self.current_month_button.grid(row=0, column=4, sticky="e")
        self.last_30_days_button = ttk.Button(config_frame, text="Ultimos 30 dias", command=self._set_last_30_days_dates)
        self.last_30_days_button.grid(row=0, column=5, sticky="e", padx=(8, 0))

        ttk.Label(config_frame, text="Banco").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(config_frame, textvariable=self.db_path_var, state="readonly").grid(
            row=1, column=1, columnspan=4, sticky="ew", padx=(6, 8), pady=(8, 0)
        )
        self.choose_db_button = ttk.Button(config_frame, text="Escolher", command=self._pick_db_path)
        self.choose_db_button.grid(row=1, column=5, sticky="e", pady=(8, 0))

        ttk.Label(config_frame, text="CSV").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(config_frame, textvariable=self.output_var, state="readonly").grid(
            row=2, column=1, columnspan=4, sticky="ew", padx=(6, 8), pady=(8, 0)
        )
        self.choose_output_button = ttk.Button(config_frame, text="Salvar como", command=self._pick_output_path)
        self.choose_output_button.grid(row=2, column=5, sticky="e", pady=(8, 0))

        ttk.Label(config_frame, text="Numero Pedido").grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.order_number_entry = ttk.Entry(config_frame, textvariable=self.order_number_var, width=20)
        self.order_number_entry.grid(row=3, column=1, sticky="w", padx=(6, 12), pady=(8, 0))

        self.start_date_entry.bind("<FocusOut>", lambda _e: self._refresh_monthly_output_default())
        self.end_date_entry.bind("<FocusOut>", lambda _e: self._refresh_monthly_output_default())

        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(4, weight=1)

        actions_frame = ttk.LabelFrame(container, text="Acoes", style="Section.TLabelframe")
        actions_frame.grid(row=3, column=0, sticky="ew", pady=(8, 0))

        self.credentials_button = ttk.Button(actions_frame, text="Credenciais", command=self._configure_credentials_clicked)
        self.credentials_button.grid(row=0, column=0, sticky="ew")
        self.test_button = ttk.Button(actions_frame, text="Testar API", command=self._test_connection_clicked)
        self.test_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.update_button = ttk.Button(actions_frame, text="Atualizar App", command=self._update_app_clicked)
        self.update_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

        self.sync_button = ttk.Button(actions_frame, text="Sincronizar Pedidos", command=self._sync_clicked)
        self.sync_button.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.export_monthly_button = ttk.Button(
            actions_frame,
            text="Exportar Mensal",
            command=self._export_monthly_clicked,
        )
        self.export_monthly_button.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.reprocess_button = ttk.Button(actions_frame, text="Reprocessar Mes", command=self._reprocess_month_clicked)
        self.reprocess_button.grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.export_orders_button = ttk.Button(actions_frame, text="Exportar Pedidos", command=self._export_orders_clicked)
        self.export_orders_button.grid(row=1, column=3, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.export_skus_button = ttk.Button(actions_frame, text="Exportar SKUs", command=self._export_skus_clicked)
        self.export_skus_button.grid(row=1, column=4, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.export_monthly_sheets_button = ttk.Button(
            actions_frame,
            text="Exportar + Enviar Sheets",
            command=self._export_monthly_sheets_clicked,
        )
        self.export_monthly_sheets_button.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        for col in range(5):
            actions_frame.columnconfigure(col, weight=1)

        log_frame = ttk.LabelFrame(container, text="Log", style="Section.TLabelframe")
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(8, 0))
        self.log_text = tk.Text(log_frame, height=14, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        container.rowconfigure(4, weight=2)

    def _log(self, message: str) -> None:
        stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.log_text.insert("end", f"[{stamp}] {message}\n")
        self.log_text.see("end")

    def _log_runtime_sources(self) -> None:
        config_path, env_path = resolve_runtime_paths()
        self._log(f"Config em uso: {config_path}")
        self._log(f"Credenciais em uso: {env_path}")

    def _rebuild_client_index(self) -> None:
        self.by_platform = {}
        for client in self.clients.values():
            self.by_platform.setdefault(client.platform, []).append(client)
        for platform in self.by_platform:
            self.by_platform[platform].sort(key=lambda c: (c.company.lower(), c.branch.lower(), c.id.lower()))

    def _reload_clients(self, preferred_client_id: str = "") -> None:
        self.clients = load_clients_config()
        self._rebuild_client_index()
        target_client = self.clients.get(preferred_client_id) if preferred_client_id else None
        if target_client:
            self._load_platforms()
            if target_client.platform in self.by_platform:
                self.platform_var.set(target_client.platform)
                self._on_platform_change(preferred_client_id=target_client.id)
            available_companies = set(self.company_combo["values"])
            if target_client.company in available_companies:
                self.company_var.set(target_client.company)
                self._on_company_change(preferred_client_id=target_client.id)
            return
        self._load_platforms()

    def _load_platforms(self) -> None:
        platforms = sorted(self.by_platform.keys())
        self.platform_combo["values"] = platforms
        if platforms:
            self.platform_var.set(platforms[0])
            self._on_platform_change()
            return
        self.platform_var.set("")
        self.company_combo["values"] = []
        self.company_var.set("")
        self._render_client_checkboxes([])

    def _on_platform_change(self, preferred_client_id: str = "") -> None:
        platform = self.platform_var.get().strip()
        clients = self.by_platform.get(platform, [])
        companies = sorted({c.company for c in clients})
        self.company_combo["values"] = companies
        self.company_var.set(companies[0] if companies else "")
        self._on_company_change(preferred_client_id=preferred_client_id)

    def _on_company_change(self, preferred_client_id: str = "") -> None:
        platform = self.platform_var.get().strip()
        company = self.company_var.get().strip()
        clients = [c for c in self.by_platform.get(platform, []) if c.company == company]
        self._render_client_checkboxes(clients, preferred_client_id=preferred_client_id)

    def _on_client_checks_configure(self, _event=None) -> None:
        self.client_canvas.configure(scrollregion=self.client_canvas.bbox("all"))

    def _on_client_canvas_configure(self, event) -> None:
        if self._client_canvas_window is not None:
            self.client_canvas.itemconfigure(self._client_canvas_window, width=event.width)

    def _render_client_checkboxes(self, clients: list, preferred_client_id: str = "") -> None:
        previous_selection = {client.id for client in self._selected_clients()}
        self._company_clients = clients

        selected_ids: set[str] = set()
        if preferred_client_id:
            selected_ids = {preferred_client_id}
        elif previous_selection:
            selected_ids = {client.id for client in clients if client.id in previous_selection}
        if not selected_ids and clients:
            selected_ids = {clients[0].id}

        next_vars: dict[str, tk.BooleanVar] = {}
        for client in clients:
            existing = self._client_check_vars.get(client.id)
            checked = client.id in selected_ids
            if existing is not None:
                existing.set(checked)
                next_vars[client.id] = existing
            else:
                next_vars[client.id] = tk.BooleanVar(value=checked)
        self._client_check_vars = next_vars

        for child in self.client_checks_frame.winfo_children():
            child.destroy()
        self._client_checkbuttons = []

        for client in self._company_clients:
            var = self._client_check_vars.get(client.id)
            if var is None:
                var = tk.BooleanVar(value=False)
                self._client_check_vars[client.id] = var
            label = (client.alias or "").strip() or client.branch
            check = ttk.Checkbutton(
                self.client_checks_frame,
                text=label,
                variable=var,
                command=self._on_client_selection_changed,
            )
            check.pack(anchor="w", fill="x")
            self._client_checkbuttons.append(check)

        if not self._company_clients:
            ttk.Label(self.client_checks_frame, text="Nenhum alias disponivel.", style="Muted.TLabel").pack(anchor="w")

        if self._busy:
            for check in self._client_checkbuttons:
                check.configure(state="disabled")
        self._on_client_checks_configure()
        self._on_client_selection_changed()

    def _update_selection_summary(self) -> None:
        selected_total = len(self._selected_clients())
        total = len(self._company_clients)

        if total == 0:
            self.selection_info_var.set("Nenhum alias para a empresa selecionada.")
            return
        self.selection_info_var.set(f"{selected_total} de {total} alias selecionados.")

    def _toggle_select_all_clients(self) -> None:
        target = self.select_all_var.get()
        for client in self._company_clients:
            var = self._client_check_vars.get(client.id)
            if var is None:
                continue
            var.set(target)
        self._on_client_selection_changed()

    def _on_client_selection_changed(self) -> None:
        client = self._selected_client()
        all_selected = bool(self._company_clients) and all(
            bool(self._client_check_vars.get(current.id) and self._client_check_vars[current.id].get())
            for current in self._company_clients
        )
        self.select_all_var.set(all_selected)
        self._update_selection_summary()
        if not client:
            self.output_var.set("")
            return
        self._apply_default_dates()
        current_output = self.output_var.get().strip().lower()
        if not current_output or current_output.endswith("_mensal.csv"):
            self.output_var.set(self._default_monthly_output(client))

    def _apply_default_dates(self) -> None:
        if self.start_date_var.get().strip() and self.end_date_var.get().strip():
            return
        self._set_current_month_dates()

    def _set_current_month_dates(self) -> None:
        today = datetime.now()
        first_day = today.replace(day=1)
        self.start_date_var.set(first_day.strftime("%d/%m/%Y"))
        self.end_date_var.set(today.strftime("%d/%m/%Y"))
        self._refresh_monthly_output_default()

    def _set_last_30_days_dates(self) -> None:
        today = datetime.now()
        start = today - timedelta(days=29)
        self.start_date_var.set(start.strftime("%d/%m/%Y"))
        self.end_date_var.set(today.strftime("%d/%m/%Y"))
        self._refresh_monthly_output_default()

    def _selected_clients(self) -> list:
        selected = []
        for client in self._company_clients:
            checked = self._client_check_vars.get(client.id)
            if checked and checked.get():
                selected.append(client)
        return selected

    def _selected_client(self):
        selected = self._selected_clients()
        if selected:
            return selected[0]
        return None

    def _require_single_selected_client(self):
        selected = self._selected_clients()
        if not selected:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return None
        if len(selected) > 1:
            messagebox.showwarning("Selecao", "Escolha apenas um alias/filial para esta acao.")
            return None
        return selected[0]

    def _validate_clients_for_yampi(self, clients: list) -> bool:
        invalid_platform = [client.id for client in clients if client.platform != "yampi"]
        if invalid_platform:
            joined = ", ".join(invalid_platform)
            messagebox.showwarning("Plataforma", f"Somente Yampi implementado neste MVP. Invalidos: {joined}")
            return False

        missing_auth = [client.id for client in clients if not _has_yampi_auth(client)]
        if missing_auth:
            joined = ", ".join(missing_auth)
            messagebox.showerror(
                "Credencial",
                f"Credenciais ausentes para: {joined}. Clique em '1) Credenciais'.",
            )
            return False
        return True

    def _output_dir_from_field(self) -> Path:
        raw = self.output_var.get().strip()
        if not raw:
            return self._downloads_dir()
        path = Path(raw).expanduser()
        if path.suffix.lower() == ".csv":
            return path.parent
        return path

    def _downloads_dir(self) -> Path:
        downloads = Path.home() / "Downloads"
        if downloads.exists():
            return downloads
        return Path.cwd()

    def _period_suffix(self) -> str:
        try:
            start = _normalize_date(self.start_date_var.get())
            if start:
                return start[:7].replace("-", "_")
        except ValueError:
            pass
        return datetime.now().strftime("%Y_%m")

    def _default_monthly_output(self, client) -> str:
        suffix = self._period_suffix()
        return str(self._downloads_dir() / f"{client.id}_{suffix}_mensal.csv")

    def _default_orders_output(self, client) -> str:
        stamp = datetime.now().strftime("%Y_%m_%d")
        return str(self._downloads_dir() / f"{client.id}_{stamp}_pedidos.csv")

    def _default_skus_output(self, client) -> str:
        stamp = datetime.now().strftime("%Y_%m_%d")
        return str(self._downloads_dir() / f"{client.id}_{stamp}_skus.csv")

    def _refresh_monthly_output_default(self) -> None:
        client = self._selected_client()
        if not client:
            return
        current = self.output_var.get().strip()
        if not current or current.lower().endswith("_mensal.csv"):
            self.output_var.set(self._default_monthly_output(client))

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

    def _open_output_folder(self, output_path: str) -> None:
        try:
            os.startfile(str(Path(output_path).resolve().parent))  # type: ignore[attr-defined]
        except Exception:
            pass

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
        self.status_var.set("Processando... aguarde." if busy else "Pronto")
        self.credentials_button.configure(state=state)
        self.test_button.configure(state=state)
        self.sync_button.configure(state=state)
        self.export_monthly_button.configure(state=state)
        self.export_monthly_sheets_button.configure(state=state)
        self.reprocess_button.configure(state=state)
        self.export_orders_button.configure(state=state)
        self.export_skus_button.configure(state=state)
        self.update_button.configure(state=state)
        self.start_date_entry.configure(state=state)
        self.end_date_entry.configure(state=state)
        self.order_number_entry.configure(state=state)
        self.current_month_button.configure(state=state)
        self.last_30_days_button.configure(state=state)
        self.choose_db_button.configure(state=state)
        self.choose_output_button.configure(state=state)
        self.platform_combo.configure(state="disabled" if busy else "readonly")
        self.company_combo.configure(state="disabled" if busy else "readonly")
        self.select_all_check.configure(state=state)
        for check in self._client_checkbuttons:
            check.configure(state=state)

    def _configure_credentials_clicked(self) -> None:
        client = self._require_single_selected_client()
        if not client:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Credenciais - {client.id}")
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")

        token_var = tk.StringVar(value=client.token or "")
        user_token_var = tk.StringVar(value=client.user_token or "")
        secret_var = tk.StringVar(value=client.user_secret_key or "")

        row = 0
        if client.token_env:
            ttk.Label(frame, text=f"Token ({client.token_env})").grid(row=row, column=0, sticky="w")
            ttk.Entry(frame, textvariable=token_var, width=62).grid(row=row + 1, column=0, sticky="ew", pady=(0, 8))
            row += 2
        if client.user_token_env:
            ttk.Label(frame, text=f"User Token ({client.user_token_env})").grid(row=row, column=0, sticky="w")
            ttk.Entry(frame, textvariable=user_token_var, width=62).grid(row=row + 1, column=0, sticky="ew", pady=(0, 8))
            row += 2
        if client.user_secret_key_env:
            ttk.Label(frame, text=f"Secret Key ({client.user_secret_key_env})").grid(row=row, column=0, sticky="w")
            ttk.Entry(frame, textvariable=secret_var, width=62, show="*").grid(row=row + 1, column=0, sticky="ew", pady=(0, 8))
            row += 2

        def save_clicked() -> None:
            try:
                save_client_credentials(
                    client=client,
                    token=token_var.get(),
                    user_token=user_token_var.get(),
                    user_secret_key=secret_var.get(),
                )
                self._reload_clients(preferred_client_id=client.id)
                self._log(f"Credenciais salvas para {client.id}.")
                messagebox.showinfo("Credenciais", "Credenciais salvas com sucesso no .env.")
                dialog.destroy()
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Erro", str(exc))

        buttons = ttk.Frame(frame)
        buttons.grid(row=row, column=0, sticky="e", pady=(4, 0))
        ttk.Button(buttons, text="Cancelar", command=dialog.destroy).pack(side="left")
        ttk.Button(buttons, text="Salvar", command=save_clicked).pack(side="left", padx=(8, 0))
        frame.columnconfigure(0, weight=1)

    def _test_connection_clicked(self) -> None:
        clients = self._selected_clients()
        if not clients:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return
        if not self._validate_clients_for_yampi(clients):
            return

        def task():
            results: list[tuple[str, bool, str]] = []
            for client in clients:
                yampi = YampiClient(
                    base_url=client.base_url,
                    token=client.token,
                    user_token=client.user_token,
                    user_secret_key=client.user_secret_key,
                )
                self.root.after(0, lambda client_id=client.id: self._log(f"Testando conexao API para {client_id}..."))
                ok, msg = yampi.test_connection(client.alias)
                results.append((client.id, ok, msg))
                if ok:
                    self.root.after(0, lambda client_id=client.id: self._log(f"Conexao OK para {client_id}."))
                else:
                    self.root.after(0, lambda client_id=client.id, detail=msg: self._log(f"Falha para {client_id}: {detail}"))

            failed = [entry for entry in results if not entry[1]]
            summary = "\n".join(
                f"{client_id}: {'OK' if ok else 'FALHA'} - {msg}" for client_id, ok, msg in results
            )
            if failed:
                self.root.after(
                    0,
                    lambda text=summary: messagebox.showwarning(
                        "API",
                        f"Teste finalizado com falhas.\n\n{text}",
                    ),
                )
                return
            self.root.after(
                0,
                lambda text=summary: messagebox.showinfo(
                    "API",
                    f"Teste finalizado com sucesso para {len(results)} cliente(s).\n\n{text}",
                ),
            )

        self._run_background(task)

    def _sync_clicked(self) -> None:
        clients = self._selected_clients()
        if not clients:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return
        if not self._validate_clients_for_yampi(clients):
            return

        try:
            start_date, end_date = _resolve_window(self.start_date_var.get(), self.end_date_var.get())
        except ValueError as exc:
            messagebox.showerror("Data invalida", str(exc))
            return
        if not start_date or not end_date:
            messagebox.showwarning("Periodo", "Informe data inicio e fim para sincronizar pedidos.")
            return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            results: list[tuple[str, int]] = []
            errors: list[tuple[str, str]] = []

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            for client in clients:
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

                try:
                    self.root.after(0, lambda client_id=client.id: self._log(f"Sincronizando {client_id}..."))
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
                    results.append((client.id, processed))
                    self.root.after(
                        0,
                        lambda client_id=client.id, count=processed: self._log(
                            f"Sincronizacao concluida para {client_id}: {count} pedidos."
                        ),
                    )
                except Exception as exc:  # noqa: BLE001
                    errors.append((client.id, str(exc)))
                    self.root.after(
                        0,
                        lambda client_id=client.id, detail=str(exc): self._log(
                            f"Erro na sincronizacao de {client_id}: {detail}"
                        ),
                    )
            conn.close()

            total = sum(count for _, count in results)
            if results:
                self.root.after(0, lambda count=total: self._log(f"Total sincronizado: {count} pedidos."))
            if errors:
                error_text = "\n".join(f"{client_id}: {detail}" for client_id, detail in errors)
                if results:
                    self.root.after(
                        0,
                        lambda text=error_text, count=total: messagebox.showwarning(
                            "Parcial",
                            f"Sincronizacao parcial.\nPedidos sincronizados: {count}\n\nFalhas:\n{text}",
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda text=error_text: messagebox.showerror(
                            "Erro",
                            f"Nenhum cliente foi sincronizado.\n\nFalhas:\n{text}",
                        ),
                    )
                return
            self.root.after(
                0,
                lambda count=total: messagebox.showinfo(
                    "Sucesso",
                    f"Sincronizacao concluida para {len(results)} cliente(s): {count} pedidos.",
                ),
            )

        self._run_background(task)

    def _export_monthly_clicked(self) -> None:
        self._export_monthly_common(upload_to_sheets=False)

    def _export_monthly_sheets_clicked(self) -> None:
        self._export_monthly_common(upload_to_sheets=True)

    def _export_monthly_common(self, upload_to_sheets: bool) -> None:
        clients = self._selected_clients()
        if not clients:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return
        if not self._validate_clients_for_yampi(clients):
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
        monthly_suffix = self._period_suffix()

        sheets_settings = None
        if upload_to_sheets:
            try:
                sheets_settings = resolve_monthly_sheets_settings()
            except ValueError as exc:
                messagebox.showerror("Google Sheets", str(exc))
                return

        single_output = ""
        if len(clients) == 1:
            single_output = self.output_var.get().strip() or self._default_monthly_output(clients[0])
            self.output_var.set(single_output)
            if not single_output.lower().endswith(".csv"):
                messagebox.showwarning("Arquivo", "O arquivo de saida deve terminar com .csv.")
                return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            output_dir = self._output_dir_from_field()
            output_dir.mkdir(parents=True, exist_ok=True)
            generated_files: list[str] = []
            errors: list[tuple[str, str]] = []
            uploaded_targets: list[tuple[str, str, int]] = []

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            for client in clients:
                output = single_output or str(output_dir / f"{client.id}_{monthly_suffix}_mensal.csv")
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

                try:
                    self.root.after(
                        0,
                        lambda client_id=client.id: self._log(
                            f"Reprocessando periodo mensal para {client_id}: {start_date} ate {end_date}..."
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
                        lambda client_id=client.id, removed=deleted, loaded=synced: self._log(
                            f"Reprocessamento de {client_id} concluido. Removidos: {removed}. Baixados: {loaded}."
                        ),
                    )
                    self.root.after(0, lambda output_path=output: self._log(f"Gerando CSV mensal: {output_path}"))
                    count = export_monthly_sheet_csv(conn, client.id, output, start_date=start_date, end_date=end_date)
                    generated_files.append(output)
                    self.root.after(
                        0,
                        lambda client_id=client.id, lines=count: self._log(
                            f"CSV mensal de {client_id} gerado com {lines} linha(s)."
                        ),
                    )
                    if upload_to_sheets and sheets_settings:
                        sheet_name = client.sheets_worksheet.strip() or client.company.strip() or client.id
                        self.root.after(
                            0,
                            lambda client_id=client.id, tab=sheet_name: self._log(
                                f"Enviando {client_id} para Google Sheets (aba '{tab}')..."
                            ),
                        )
                        uploaded_rows = upload_csv_to_google_sheet(
                            csv_path=output,
                            spreadsheet_id=sheets_settings.spreadsheet_id,
                            worksheet_name=sheet_name,
                            credentials_json_path=sheets_settings.credentials_json_path,
                            clear_before_upload=sheets_settings.clear_before_upload,
                            create_worksheet_if_missing=sheets_settings.create_worksheet_if_missing,
                        )
                        uploaded_targets.append((client.id, sheet_name, uploaded_rows))
                        self.root.after(
                            0,
                            lambda client_id=client.id, tab=sheet_name, rows=uploaded_rows: self._log(
                                f"Google Sheets atualizado para {client_id} (aba '{tab}', {rows} linha(s))."
                            ),
                        )
                except Exception as exc:  # noqa: BLE001
                    errors.append((client.id, str(exc)))
                    self.root.after(
                        0,
                        lambda client_id=client.id, detail=str(exc): self._log(
                            f"Erro na exportacao mensal de {client_id}: {detail}"
                        ),
                    )
            conn.close()

            if errors:
                error_text = "\n".join(f"{client_id}: {detail}" for client_id, detail in errors)
                if generated_files:
                    self.root.after(
                        0,
                        lambda total=len(generated_files), text=error_text: messagebox.showwarning(
                            "Parcial",
                            f"Exportacao mensal parcial.\nArquivos gerados: {total}\n\nFalhas:\n{text}",
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda text=error_text: messagebox.showerror(
                            "Erro",
                            f"Nenhum CSV mensal foi gerado.\n\nFalhas:\n{text}",
                        ),
                    )
                if generated_files:
                    self.root.after(0, lambda first_file=generated_files[0]: self._open_output_folder(first_file))
                return

            if upload_to_sheets:
                self.root.after(
                    0,
                    lambda total=len(generated_files), uploaded=len(uploaded_targets): messagebox.showinfo(
                        "Sucesso",
                        (
                            f"CSV mensal gerado para {total} cliente(s).\n"
                            f"Google Sheets atualizado para {uploaded} cliente(s)."
                        ),
                    ),
                )
            else:
                self.root.after(
                    0,
                    lambda total=len(generated_files): messagebox.showinfo(
                        "Sucesso",
                        f"CSV mensal gerado para {total} cliente(s).",
                    ),
                )
            if generated_files:
                self.root.after(0, lambda first_file=generated_files[0]: self._open_output_folder(first_file))

        self._run_background(task)

    def _reprocess_month_clicked(self) -> None:
        clients = self._selected_clients()
        if not clients:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return
        if not self._validate_clients_for_yampi(clients):
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
            f"Apagar dados locais de {len(clients)} cliente(s) entre {month_start} e {month_end} e baixar novamente?",
        )
        if not proceed:
            return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            totals_deleted = 0
            totals_synced = 0
            errors: list[tuple[str, str]] = []

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            for client in clients:
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

                try:
                    self.root.after(
                        0,
                        lambda client_id=client.id: self._log(
                            f"Reprocessando mes {month_start[:7]} para {client_id}..."
                        ),
                    )
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
                    totals_deleted += deleted
                    totals_synced += synced
                    self.root.after(
                        0,
                        lambda client_id=client.id, removed=deleted, loaded=synced: self._log(
                            f"Reprocessamento de {client_id} concluido. Removidos: {removed}. Baixados: {loaded}."
                        ),
                    )
                except Exception as exc:  # noqa: BLE001
                    errors.append((client.id, str(exc)))
                    self.root.after(
                        0,
                        lambda client_id=client.id, detail=str(exc): self._log(
                            f"Erro no reprocessamento de {client_id}: {detail}"
                        ),
                    )
            conn.close()

            if errors:
                error_text = "\n".join(f"{client_id}: {detail}" for client_id, detail in errors)
                if totals_deleted or totals_synced:
                    self.root.after(
                        0,
                        lambda removed=totals_deleted, loaded=totals_synced, text=error_text: messagebox.showwarning(
                            "Parcial",
                            f"Reprocessamento parcial.\nRemovidos: {removed}\nBaixados: {loaded}\n\nFalhas:\n{text}",
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda text=error_text: messagebox.showerror(
                            "Erro",
                            f"Nenhum cliente foi reprocessado.\n\nFalhas:\n{text}",
                        ),
                    )
                return

            self.root.after(
                0,
                lambda removed=totals_deleted, loaded=totals_synced: self._log(
                    f"Reprocessamento concluido. Removidos: {removed}. Baixados novamente: {loaded}."
                ),
            )
            self.root.after(
                0,
                lambda removed=totals_deleted, loaded=totals_synced: messagebox.showinfo(
                    "Sucesso",
                    f"Reprocessamento concluido.\nRemovidos: {removed}\nBaixados: {loaded}",
                ),
            )

        self._run_background(task)

    def _export_orders_clicked(self) -> None:
        clients = self._selected_clients()
        if not clients:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return

        single_output = ""
        if len(clients) == 1:
            single_output = self.output_var.get().strip()
            if not single_output or single_output.lower().endswith("_mensal.csv"):
                single_output = self._default_orders_output(clients[0])
                self.output_var.set(single_output)
            if not single_output.lower().endswith(".csv"):
                messagebox.showwarning("Arquivo", "O arquivo de saida deve terminar com .csv.")
                return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            output_dir = self._output_dir_from_field()
            output_dir.mkdir(parents=True, exist_ok=True)
            generated_files: list[str] = []
            errors: list[tuple[str, str]] = []

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            for client in clients:
                output = single_output or str(output_dir / Path(self._default_orders_output(client)).name)
                try:
                    self.root.after(0, lambda client_id=client.id, path=output: self._log(f"Gerando CSV detalhado de {client_id}: {path}"))
                    count = export_orders_csv(conn, client.id, output)
                    generated_files.append(output)
                    self.root.after(
                        0,
                        lambda client_id=client.id, lines=count: self._log(
                            f"CSV detalhado de {client_id} gerado com {lines} linha(s)."
                        ),
                    )
                except Exception as exc:  # noqa: BLE001
                    errors.append((client.id, str(exc)))
                    self.root.after(
                        0,
                        lambda client_id=client.id, detail=str(exc): self._log(
                            f"Erro na exportacao detalhada de {client_id}: {detail}"
                        ),
                    )
            conn.close()

            if errors:
                error_text = "\n".join(f"{client_id}: {detail}" for client_id, detail in errors)
                if generated_files:
                    self.root.after(
                        0,
                        lambda total=len(generated_files), text=error_text: messagebox.showwarning(
                            "Parcial",
                            f"Exportacao detalhada parcial.\nArquivos gerados: {total}\n\nFalhas:\n{text}",
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda text=error_text: messagebox.showerror(
                            "Erro",
                            f"Nenhum CSV detalhado foi gerado.\n\nFalhas:\n{text}",
                        ),
                    )
                if generated_files:
                    self.root.after(0, lambda first_file=generated_files[0]: self._open_output_folder(first_file))
                return

            self.root.after(
                0,
                lambda total=len(generated_files): messagebox.showinfo(
                    "Sucesso",
                    f"CSV detalhado gerado para {total} cliente(s).",
                ),
            )
            if generated_files:
                self.root.after(0, lambda first_file=generated_files[0]: self._open_output_folder(first_file))

        self._run_background(task)

    def _export_skus_clicked(self) -> None:
        clients = self._selected_clients()
        if not clients:
            messagebox.showwarning("Selecao", "Escolha pelo menos um alias/filial.")
            return

        order_number = self.order_number_var.get().strip()
        start_date = ""
        end_date = ""
        if not order_number:
            try:
                start_date, end_date = _resolve_window(self.start_date_var.get(), self.end_date_var.get())
            except ValueError as exc:
                messagebox.showerror("Data invalida", str(exc))
                return
            if not start_date or not end_date:
                messagebox.showwarning(
                    "Filtro",
                    "Informe Numero Pedido ou data inicio e fim.",
                )
                return

        single_output = ""
        if len(clients) == 1:
            single_output = self.output_var.get().strip()
            if (
                not single_output
                or single_output.lower().endswith("_mensal.csv")
                or single_output.lower().endswith("_pedidos.csv")
            ):
                single_output = self._default_skus_output(clients[0])
                self.output_var.set(single_output)
            if not single_output.lower().endswith(".csv"):
                messagebox.showwarning("Arquivo", "O arquivo de saida deve terminar com .csv.")
                return

        def task():
            db_path = self.db_path_var.get().strip() or "data/local.db"
            conn = connect(db_path)
            init_db(conn)
            output_dir = self._output_dir_from_field()
            output_dir.mkdir(parents=True, exist_ok=True)
            generated_files: list[str] = []
            errors: list[tuple[str, str]] = []

            self.root.after(0, lambda: self._log(f"Banco em uso: {db_path}"))
            for client in clients:
                output = single_output or str(output_dir / Path(self._default_skus_output(client)).name)
                try:
                    filtro = f"pedido={order_number}" if order_number else f"periodo={start_date} ate {end_date}"
                    self.root.after(
                        0,
                        lambda client_id=client.id, path=output, filtro_texto=filtro: self._log(
                            f"Gerando CSV SKUs de {client_id} ({filtro_texto}): {path}"
                        ),
                    )
                    count = export_order_skus_csv(
                        conn,
                        client.id,
                        output,
                        start_date=start_date,
                        end_date=end_date,
                        order_number=order_number,
                    )
                    generated_files.append(output)
                    self.root.after(
                        0,
                        lambda client_id=client.id, lines=count: self._log(
                            f"CSV SKUs de {client_id} gerado com {lines} linha(s)."
                        ),
                    )
                except PermissionError:
                    detail = (
                        f"Arquivo em uso: {output}. "
                        "Feche o CSV no Excel/planilha (ou escolha outro nome) e tente novamente."
                    )
                    errors.append((client.id, detail))
                    self.root.after(
                        0,
                        lambda client_id=client.id, msg=detail: self._log(
                            f"Erro na exportacao de SKUs de {client_id}: {msg}"
                        ),
                    )
                except Exception as exc:  # noqa: BLE001
                    errors.append((client.id, str(exc)))
                    self.root.after(
                        0,
                        lambda client_id=client.id, detail=str(exc): self._log(
                            f"Erro na exportacao de SKUs de {client_id}: {detail}"
                        ),
                    )
            conn.close()

            if errors:
                error_text = "\n".join(f"{client_id}: {detail}" for client_id, detail in errors)
                if generated_files:
                    self.root.after(
                        0,
                        lambda total=len(generated_files), text=error_text: messagebox.showwarning(
                            "Parcial",
                            f"Exportacao de SKUs parcial.\nArquivos gerados: {total}\n\nFalhas:\n{text}",
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda text=error_text: messagebox.showerror(
                            "Erro",
                            f"Nenhum CSV de SKUs foi gerado.\n\nFalhas:\n{text}",
                        ),
                    )
                if generated_files:
                    self.root.after(0, lambda first_file=generated_files[0]: self._open_output_folder(first_file))
                return

            self.root.after(
                0,
                lambda total=len(generated_files): messagebox.showinfo(
                    "Sucesso",
                    f"CSV de SKUs gerado para {total} cliente(s).",
                ),
            )
            if generated_files:
                self.root.after(0, lambda first_file=generated_files[0]: self._open_output_folder(first_file))

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
