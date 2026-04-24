# -*- coding: utf-8 -*-
# Patch para duplo clique: garante diretório e import do core
import os, sys
APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)
sys.path.insert(0, APP_DIR)

import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import shutil

import pandas as pd
import pythoncom
import importlib

# Importa o core
try:
    core = importlib.import_module("faturas_b3")
except Exception as e:
    raise SystemExit(
        "Não foi possível importar 'faturas_b3.py'. "
        "Coloque este arquivo (ui_faturas_b3.py) na mesma pasta do seu script e tente novamente.\n\n"
        f"Erro: {e}"
    )

# Tenta importar tkinterdnd2 para drag-and-drop
try:
    from tkinterdnd2 import DND_FILES, DND_TEXT, DND_ALL, Tk as DND_Tk
    HAS_DRAG_DROP = True
except ImportError:
    HAS_DRAG_DROP = False
    DND_Tk = tk.Tk

# Cores do tema
COLOR_WHITE = "#FFFFFF"
COLOR_DARK_GRAY = "#2B2B2B"
COLOR_LIGHT_GRAY = "#F5F5F5"
COLOR_MED_GRAY = "#CCCCCC"
COLOR_ACCENT = "#0078D4"

class FaturasB3GUI(DND_Tk if HAS_DRAG_DROP else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Processador de Faturas B3")
        self.geometry("900x650")
        self.minsize(800, 600)
        self.configure(bg=COLOR_WHITE)
        
        # Estilo customizado com cores branco/cinza
        style = ttk.Style()
        style.theme_use("clam")
        
        # Configurar cores do ttk
        style.configure("TFrame", background=COLOR_WHITE)
        style.configure("TLabel", background=COLOR_WHITE, foreground=COLOR_DARK_GRAY)
        style.configure("TButton", background=COLOR_LIGHT_GRAY, foreground=COLOR_DARK_GRAY)
        style.configure("TLabelFrame", background=COLOR_WHITE, foreground=COLOR_DARK_GRAY)
        style.configure("TLabelFrame.Label", background=COLOR_WHITE, foreground=COLOR_DARK_GRAY)
        style.map("TButton", background=[("active", COLOR_MED_GRAY)])

        self._build_vars()
        self._build_layout()

        self.worker_thread = None
        self.is_running = False
        
        # Configurar drag-and-drop
        self._setup_drag_drop()

    # ---------------------------
    # Vars
    # ---------------------------
    def _build_vars(self):
        # Defaults puxados do core
        root_default = getattr(core, "ROOT", os.path.expanduser("~"))
        data_default = getattr(core, "DATA", "022026")
        reply_all_default = getattr(core, "REPLY_ALL", True)
        cc_emails_default = getattr(core, "CC_EMAILS", "")
        my_email_default = getattr(core, "MY_EMAIL", "")
        open_modal_default = getattr(core, "OPEN_REPLY_MODAL", False)
        stop_after_open_default = getattr(core, "STOP_AFTER_OPENING_REPLY", True)
        preview_default = getattr(core, "PREVIEW_EXCEL", True)
        pause_default = getattr(core, "PAUSE_BEFORE_REPLY", False)
        preview_readonly_default = getattr(core, "PREVIEW_READONLY", True)

        # Data atual para Competência e Envio - AUTOMÁTICAS
        today = datetime.now()
        data_competencia_default = f"{today.month:02d}{today.year}"  # MMYYYY
        data_envio_default = today.strftime("%d/%m/%Y")

        self.var_root = tk.StringVar(value=root_default)
        self.var_data = tk.StringVar(value=data_competencia_default)  # Data automática
        self.var_date_envio = tk.StringVar(value=data_envio_default)  # Data automática
        self.var_reply_all = tk.BooleanVar(value=reply_all_default)
        self.var_cc = tk.StringVar(value=cc_emails_default)
        self.var_my_email = tk.StringVar(value=my_email_default)
        self.var_open_modal = tk.BooleanVar(value=open_modal_default)
        self.var_stop_after_open = tk.BooleanVar(value=stop_after_open_default)
        self.var_preview = tk.BooleanVar(value=preview_default)
        self.var_pause_before = tk.BooleanVar(value=pause_default)
        self.var_preview_readonly = tk.BooleanVar(value=preview_readonly_default)

        # Preview do "MÊS" (01/MM/YYYY)
        self.var_mes_preview = tk.StringVar(value=self._mes_preview(self.var_data.get()))

        # Status
        self.var_status = tk.StringVar(value="Pronto.")

    # ---------------------------
    # UI - Design Simples e Limpo
    # ---------------------------
    def _build_layout(self):
        # ==== CABEÇALHO COM TÍTULO E ENGRENAGEM ====
        frm_header = ttk.Frame(self)
        frm_header.pack(fill="x", padx=20, pady=15)
        
        lbl_title = tk.Label(frm_header, text="Processador de Faturas B3", 
                            font=("Arial", 18, "bold"), bg=COLOR_WHITE, fg=COLOR_DARK_GRAY)
        lbl_title.pack(side="left", expand=True, anchor="w")
        
        # Botão de Engrenagem (Configurações)
        btn_settings = tk.Button(frm_header, text="⚙️", font=("Arial", 20), 
                                bg=COLOR_WHITE, fg=COLOR_DARK_GRAY, border=0, 
                                command=self._open_settings_window, cursor="hand2")
        btn_settings.pack(side="right")
        
        # ==== SEÇÃO PRINCIPAL: INFORMAÇÕES BÁSICAS ====
        frm_main = ttk.LabelFrame(self, text="Informações Básicas", padding=15)
        frm_main.pack(fill="both", expand=False, padx=20, pady=10)
        
        # Pasta ROOT
        row1 = ttk.Frame(frm_main)
        row1.pack(fill="x", pady=8)
        tk.Label(row1, text="Pasta:", font=("Arial", 10, "bold"), 
                fg=COLOR_DARK_GRAY, width=12).pack(side="left")
        self.ent_root = ttk.Entry(row1, textvariable=self.var_root, width=50)
        self.ent_root.pack(side="left", fill="x", expand=True, padx=10)
        ttk.Button(row1, text="Procurar", command=self._choose_root, width=10).pack(side="left")
        
        # Data Competência (automática)
        row2 = ttk.Frame(frm_main)
        row2.pack(fill="x", pady=8)
        tk.Label(row2, text="Competência:", font=("Arial", 10, "bold"), 
                fg=COLOR_DARK_GRAY, width=12).pack(side="left")
        lbl_data = tk.Label(row2, textvariable=self.var_mes_preview, 
                           fg=COLOR_ACCENT, font=("Arial", 10), bg=COLOR_WHITE)
        lbl_data.pack(side="left", padx=10)
        
        # Data Envio (automática)
        row3 = ttk.Frame(frm_main)
        row3.pack(fill="x", pady=8)
        tk.Label(row3, text="Data Envio:", font=("Arial", 10, "bold"), 
                fg=COLOR_DARK_GRAY, width=12).pack(side="left")
        lbl_date_envio = tk.Label(row3, textvariable=self.var_date_envio, 
                                 fg=COLOR_ACCENT, font=("Arial", 10), bg=COLOR_WHITE)
        lbl_date_envio.pack(side="left", padx=10)
        
        # ==== BOTÃO PRINCIPAL: PROCESSAR ====
        frm_action = ttk.Frame(self)
        frm_action.pack(fill="x", padx=20, pady=20)
        
        self.btn_process = tk.Button(frm_action, text="▶  PROCESSAR E-MAILS", 
                                    font=("Arial", 12, "bold"), bg=COLOR_ACCENT, 
                                    fg=COLOR_WHITE, command=self._on_process, 
                                    padx=30, pady=10, cursor="hand2", border=0)
        self.btn_process.pack(side="left", padx=5)
        
        tk.Label(frm_action, text="Processa e-mails da pasta, gera Excel e responde", 
                fg=COLOR_DARK_GRAY, font=("Arial", 9), bg=COLOR_WHITE).pack(side="left", padx=15)
        
        # ==== AÇÕES SECUNDÁRIAS ====
        frm_secondary = ttk.LabelFrame(self, text="Ações Secundárias", padding=10)
        frm_secondary.pack(fill="x", padx=20, pady=10)
        
        row_secondary = ttk.Frame(frm_secondary)
        row_secondary.pack(fill="x")
        
        btn_s1 = tk.Button(row_secondary, text="📂 Abrir Input", font=("Arial", 9), 
                          bg=COLOR_LIGHT_GRAY, fg=COLOR_DARK_GRAY, border=1, 
                          command=self._open_inbox_dir, cursor="hand2")
        btn_s1.pack(side="left", padx=5, pady=5)
        
        btn_s2 = tk.Button(row_secondary, text="✓ Testar Base3", font=("Arial", 9), 
                          bg=COLOR_LIGHT_GRAY, fg=COLOR_DARK_GRAY, border=1, 
                          command=self._test_base, cursor="hand2")
        btn_s2.pack(side="left", padx=5, pady=5)
        
        btn_s3 = tk.Button(row_secondary, text="📊 Ver Excel", font=("Arial", 9), 
                          bg=COLOR_LIGHT_GRAY, fg=COLOR_DARK_GRAY, border=1, 
                          command=self._open_excel_only, cursor="hand2")
        btn_s3.pack(side="left", padx=5, pady=5)
        
        btn_s4 = tk.Button(row_secondary, text="📋 Ver Log", font=("Arial", 9), 
                          bg=COLOR_LIGHT_GRAY, fg=COLOR_DARK_GRAY, border=1, 
                          command=self._log_excel_only, cursor="hand2")
        btn_s4.pack(side="left", padx=5, pady=5)
        
        # ==== LOG ====
        frm_log = ttk.LabelFrame(self, text="Log de Execução", padding=8)
        frm_log.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.txt_log = ScrolledText(frm_log, height=10, font=("Courier", 9), 
                                   bg=COLOR_LIGHT_GRAY, fg=COLOR_DARK_GRAY)
        self.txt_log.pack(fill="both", expand=True, padx=5, pady=5)
        self._log("✓ Sistema pronto. Configure os dados e clique em 'PROCESSAR E-MAILS'.")
        
        # ==== RODAPÉ ====
        frm_footer = ttk.Frame(self)
        frm_footer.pack(fill="x", padx=20, pady=10)
        
        self.progress = ttk.Progressbar(frm_footer, mode="determinate", length=200)
        self.progress.pack(side="left")
        
        tk.Label(frm_footer, textvariable=self.var_status, font=("Arial", 10, "bold"), 
                bg=COLOR_WHITE, fg=COLOR_DARK_GRAY).pack(side="left", padx=15)

    def _open_settings_window(self):
        """Abre janela modal com configurações avançadas."""
        settings_win = tk.Toplevel(self)
        settings_win.title("⚙️ Configurações Avançadas")
        settings_win.geometry("600x500")
        settings_win.resizable(False, False)
        settings_win.configure(bg=COLOR_WHITE)
        settings_win.transient(self)
        settings_win.grab_set()
        
        # Centralizar na janela pai
        settings_win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 300
        y = self.winfo_y() + (self.winfo_height() // 2) - 250
        settings_win.geometry(f"+{x}+{y}")
        
        # ========== SEÇÃO: TIPO DE RESPOSTA ==========
        frm_reply = ttk.LabelFrame(settings_win, text="Tipo de Resposta", padding=10)
        frm_reply.pack(fill="x", padx=15, pady=10)
        
        r1 = ttk.Radiobutton(frm_reply, text="Responder a TODOS (ReplyAll)", variable=self.var_reply_all, value=True)
        r1.pack(anchor="w", pady=5)
        
        r2 = ttk.Radiobutton(frm_reply, text="Somente ao REMETENTE (Reply)", variable=self.var_reply_all, value=False)
        r2.pack(anchor="w", pady=5)
        
        # ========== SEÇÃO: E-MAILS ==========
        frm_emails = ttk.LabelFrame(settings_win, text="Configuração de E-mails", padding=10)
        frm_emails.pack(fill="x", padx=15, pady=10)
        
        tk.Label(frm_emails, text="Cc adicionais (separar com ;):", fg=COLOR_DARK_GRAY, 
                font=("Arial", 9, "bold")).pack(anchor="w", pady=5)
        ent_cc = ttk.Entry(frm_emails, textvariable=self.var_cc, width=50)
        ent_cc.pack(fill="x", pady=5)
        
        tk.Label(frm_emails, text="Seu e-mail (opcional, para evitar loops):", fg=COLOR_DARK_GRAY, 
                font=("Arial", 9, "bold")).pack(anchor="w", pady=5)
        ent_my_email = ttk.Entry(frm_emails, textvariable=self.var_my_email, width=50)
        ent_my_email.pack(fill="x", pady=5)
        
        # ========== SEÇÃO: PROCESSAMENTO EXCEL ==========
        frm_excel = ttk.LabelFrame(settings_win, text="Processamento do Excel", padding=10)
        frm_excel.pack(fill="x", padx=15, pady=10)
        
        chk_preview = ttk.Checkbutton(frm_excel, text="Pré-visualizar Excel após gerar", variable=self.var_preview)
        chk_preview.pack(anchor="w", pady=5)
        
        chk_readonly = ttk.Checkbutton(frm_excel, text="Abrir como SOMENTE LEITURA", variable=self.var_preview_readonly)
        chk_readonly.pack(anchor="w", padx=20, pady=5)
        
        chk_pause = ttk.Checkbutton(frm_excel, text="Pausar antes de responder o e-mail", variable=self.var_pause_before)
        chk_pause.pack(anchor="w", pady=5)
        
        # ========== SEÇÃO: COMPORTAMENTO ==========
        frm_behavior = ttk.LabelFrame(settings_win, text="Comportamento", padding=10)
        frm_behavior.pack(fill="x", padx=15, pady=10)
        
        chk_modal = ttk.Checkbutton(frm_behavior, text="Janela de resposta MODAL", variable=self.var_open_modal)
        chk_modal.pack(anchor="w", pady=5)
        
        chk_stop = ttk.Checkbutton(frm_behavior, text="Encerrar após abrir primeira resposta", variable=self.var_stop_after_open)
        chk_stop.pack(anchor="w", pady=5)
        
        # ========== BOTÕES ==========
        frm_buttons = ttk.Frame(settings_win)
        frm_buttons.pack(fill="x", padx=15, pady=15)
        
        btn_close = tk.Button(frm_buttons, text="Fechar", font=("Arial", 10), 
                             bg=COLOR_ACCENT, fg=COLOR_WHITE, command=settings_win.destroy, 
                             padx=20, cursor="hand2", border=0)
        btn_close.pack(side="right")

    # ---------------------------
    # Drag-and-Drop
    # ---------------------------
    def _setup_drag_drop(self):
        """Configura drag-and-drop para aceitar arquivos .msg."""
        if not HAS_DRAG_DROP:
            self._log("[INFO] tkinterdnd2 não está instalado. Drag-and-drop desativado.")
            return
        
        try:
            # Registra a janela para aceitar drops
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
            self._log("[OK] Drag-and-drop ativado. Você pode arrastar arquivos .msg para a interface.")
        except Exception as e:
            self._log(f"[AVISO] Não foi possível ativar drag-and-drop: {e}")

    def _on_drop(self, event):
        """Handler para quando um arquivo é arrastado para a janela."""
        # event.data pode conter múltiplos arquivos separados por espaço ou entre chaves
        files = self._parse_drop_data(event.data)
        
        if not files:
            messagebox.showwarning("Atenção", "Nenhum arquivo detectado.")
            return
        
        # Filtrar apenas arquivos .msg
        msg_files = [f for f in files if f.lower().endswith(".msg")]
        
        if not msg_files:
            messagebox.showwarning("Atenção", "Por favor, arraste apenas arquivos .msg (e-mails).")
            return
        
        if len(msg_files) > 1:
            messagebox.showwarning("Atenção", f"Apenas 1 arquivo de cada vez. Você arrastar {len(msg_files)}.")
            return
        
        # Processar o arquivo
        self._process_dropped_email(msg_files[0])

    def _parse_drop_data(self, data: str) -> list:
        """Converte dados de drop em lista de caminhos de arquivo."""
        # Tkinterdnd2 pode passar arquivos entre chaves ou com espaços
        files = []
        import re
        
        # Remover chaves externas se presentes
        data = data.strip()
        if data.startswith("{") and data.endswith("}"):
            data = data[1:-1]
        
        # Tentar separar por chaves individuais
        pattern = r"\{([^}]+)\}|([^\s]+)"
        matches = re.findall(pattern, data)
        
        for match in matches:
            file_path = match[0] or match[1]
            if file_path:
                # Remover aspas se presentes
                file_path = file_path.strip('"\'')
                if os.path.exists(file_path):
                    files.append(file_path)
        
        return files

    def _process_dropped_email(self, msg_path: str):
        """Processa o e-mail arrastado: limpa input e adiciona o novo."""
        root = self.var_root.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showwarning("Atenção", "Defina a pasta ROOT primeiro.")
            return
        
        input_dir = os.path.join(root, "input")
        os.makedirs(input_dir, exist_ok=True)
        
        try:
            # Apagar arquivos .msg antigos na pasta input
            for file in os.listdir(input_dir):
                if file.lower().endswith(".msg"):
                    file_path = os.path.join(input_dir, file)
                    try:
                        os.remove(file_path)
                        self._log(f"[INFO] Arquivo removido: {file}")
                    except Exception as e:
                        self._log(f"[AVISO] Não foi possível remover {file}: {e}")
            
            # Copiar novo arquivo para input
            file_name = os.path.basename(msg_path)
            dest_path = os.path.join(input_dir, file_name)
            
            shutil.copy2(msg_path, dest_path)
            self._log(f"[OK] E-mail adicionado: {file_name}")
            
            messagebox.showinfo("Sucesso", f"E-mail '{file_name}' adicionado à pasta input.\n\n"
                                         f"Agora clique em 'PROCESSAR E-MAILS' para continuar.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n\n{e}")
            self._log(f"[ERRO] {e}")

    # ---------------------------
    # Helpers
    # ---------------------------
    @staticmethod
    def _mes_preview(data_str: str) -> str:
        # Espera MMYYYY
        if len(data_str) == 6 and data_str.isdigit():
            mm, yyyy = data_str[:2], data_str[2:]
            return f"MÊS (coluna): 01/{mm}/{yyyy}"
        return "MÊS (coluna): —"

    def _log(self, msg: str):
        self.txt_log.insert("end", msg.rstrip() + "\n")
        self.txt_log.see("end")

    def _choose_root(self):
        choice = filedialog.askdirectory(title="Escolha a pasta ROOT")
        if choice:
            self.var_root.set(choice)

    def _open_inbox_dir(self):
        root = self.var_root.get().strip()
        if not root:
            messagebox.showwarning("Atenção", "Defina a pasta ROOT.")
            return
        inbox = os.path.join(root, "input")
        try:
            os.makedirs(inbox, exist_ok=True)
            os.startfile(inbox)  # Windows
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta 'input'.\n\n{e}")

    def _test_base(self):
        # Apenas tenta resolver o Base3.xlsx e informa resultado.
        try:
            old_root = getattr(core, "ROOT", "")
            core.ROOT = self.var_root.get().strip() or old_root
            path = core.resolve_base_xlsx()
            messagebox.showinfo("OK", f"Base encontrada:\n{path}")
            self._log(f"[OK] Base encontrada: {path}")
        except Exception as e:
            messagebox.showerror("Falha", str(e))
            self._log(f"[ERRO] {e}")

    def _latest_output_xlsx(self):
        """Retorna o caminho do .xlsx mais recente na pasta output do ROOT atual."""
        root = self.var_root.get().strip()
        if not root or not os.path.isdir(root):
            return None
        output_dir = os.path.join(root, "output")
        if not os.path.isdir(output_dir):
            return None
        files = [
            os.path.join(output_dir, f)
            for f in os.listdir(output_dir)
            if f.lower().endswith(".xlsx")
        ]
        if not files:
            return None
        files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return files[0]

    def _open_excel_async(self, xlsx_path: str):
        """Abre o Excel em uma thread separada (COM inicializado), com fallback para os.startfile."""
        def runner():
            try:
                pythoncom.CoInitialize()
                try:
                    core.preview_excel(xlsx_path)  # respeita PREVIEW_READONLY e AUTOFIT
                finally:
                    pythoncom.CoUninitialize()
            except Exception:
                try:
                    os.startfile(xlsx_path)
                except Exception as e2:
                    self._log(f"[ERRO] Não foi possível abrir o Excel: {e2}")
        threading.Thread(target=runner, daemon=True).start()

    # ---------------------------
    # NOVOS BOTÕES: ações
    # ---------------------------
    def _log_excel_only(self):
        """Mostra no log todo o conteúdo do Excel mais recente (todas as planilhas)."""
        xlsx_path = self._latest_output_xlsx()
        if not xlsx_path:
            messagebox.showinfo("Sem arquivo", "Nenhum arquivo .xlsx encontrado na pasta 'output'.")
            return

        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(xlsx_path)).strftime("%d/%m/%Y %H:%M:%S")
            self._log("============================================")
            self._log(f"[INFO] Excel mais recente: {xlsx_path}")
            self._log(f"[INFO] Modificado em: {mtime}")

            # Lê TODAS as planilhas
            xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
            sheet_names = xls.sheet_names
            self._log(f"[INFO] Planilhas encontradas: {', '.join(sheet_names)}")

            for sname in sheet_names:
                df = pd.read_excel(xls, sheet_name=sname, engine="openpyxl")
                rows, cols = df.shape
                self._log(f"\n--- [{sname}] {rows} linha(s) x {cols} coluna(s) ---")
                self._log(f"Colunas: {list(df.columns)}")

                # Se for muito grande, confirma antes de despejar tudo no log
                PRINT_LIMIT = 3000  # limite de segurança para não travar a UI
                if rows > PRINT_LIMIT:
                    go = messagebox.askyesno(
                        "Arquivo grande",
                        f"A planilha '{sname}' tem {rows} linha(s). Deseja imprimir TUDO no log?\n"
                        f"Isso pode deixar a interface lenta."
                    )
                    if not go:
                        head = df.head(30).to_string(index=False)
                        tail = df.tail(30).to_string(index=False)
                        self._log("[PARCIAL] Primeiras 30 linhas:")
                        self._log(head)
                        self._log("[PARCIAL] Últimas 30 linhas:")
                        self._log(tail)
                        continue

                with pd.option_context("display.max_rows", None, "display.max_columns", None, "display.width", 0):
                    text = df.to_string(index=False)
                self._log(text)

            self._log("============================================")
        except Exception:
            err = traceback.format_exc()
            self._log("❌ ERRO ao ler o Excel:")
            self._log(err)

    def _open_excel_only(self):
        """Abre o Excel mais recente na pasta 'output'."""
        xlsx_path = self._latest_output_xlsx()
        if not xlsx_path:
            messagebox.showinfo("Sem arquivo", "Nenhum arquivo .xlsx encontrado na pasta 'output'.")
            return
        self._open_excel_async(xlsx_path)

    def _reply_with_existing(self):
        """Escolhe um .msg e um .xlsx existente e abre a resposta com esse arquivo, sem reprocessar PDFs."""
        root = self.var_root.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showwarning("Atenção", "Selecione uma pasta ROOT válida.")
            return

        # Escolhe o .msg
        initial_msg_dir = os.path.join(root, "input") if os.path.isdir(os.path.join(root, "input")) else root
        msg_path = filedialog.askopenfilename(
            title="Selecione o e-mail (.msg) a responder",
            initialdir=initial_msg_dir,
            filetypes=[("Arquivos MSG", "*.msg"), ("Todos os arquivos", "*.*")]
        )
        if not msg_path:
            return

        # Sugere o .xlsx mais recente ou permite escolher outro
        xlsx_path = self._latest_output_xlsx()
        use_latest = False
        if xlsx_path:
            use_latest = messagebox.askyesno(
                "Usar Excel mais recente?",
                f"Deseja usar este arquivo?\n\n{xlsx_path}\n\nSim = Usar este | Não = Escolher outro"
            )
        if not use_latest:
            initial_xlsx_dir = os.path.join(root, "output") if os.path.isdir(os.path.join(root, "output")) else root
            xlsx_path = filedialog.askopenfilename(
                title="Selecione o Excel (.xlsx) já editado",
                initialdir=initial_xlsx_dir,
                filetypes=[("Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
            )
            if not xlsx_path:
                return

        # Roda em thread (COM inicializado)
        def runner():
            pythoncom.CoInitialize()
            try:
                importlib.reload(core)
                # Atualiza configurações do core
                core.ROOT = root
                core.INBOX_DIR     = os.path.join(core.ROOT, "input")
                core.PROCESSED_DIR = os.path.join(core.ROOT, "processed")
                core.ERROR_DIR     = os.path.join(core.ROOT, "error")
                core.TEMP_DIR      = os.path.join(core.ROOT, "temp")
                core.OUTPUT_DIR    = os.path.join(core.ROOT, "output")

                core.DATA = self.var_data.get().strip()
                core.REPLY_ALL = bool(self.var_reply_all.get())
                core.CC_EMAILS = self.var_cc.get().strip()
                core.MY_EMAIL = self.var_my_email.get().strip()
                core.OPEN_REPLY_MODAL = bool(self.var_open_modal.get())
                
                # Data de envio salvo para uso se necessário
                date_envio = self.var_date_envio.get().strip()

                self._log("============================================")
                self._log(f"[INFO] Reply com Excel existente:\nMSG = {msg_path}\nXLSX = {xlsx_path}")
                ok = core.reply_with_existing(xlsx_path, msg_path)
                if ok:
                    self._log("[OK] Janela de resposta aberta (Excel existente).")
                else:
                    self._log("❌ Falha ao abrir resposta com Excel existente.")
            except Exception:
                err = traceback.format_exc()
                self._log("❌ ERRO no fluxo 'Responder com Excel EXISTENTE':")
                self._log(err)
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=runner, daemon=True).start()

    # ---------------------------
    # Execução
    # ---------------------------
    def _validate(self) -> bool:
        root = self.var_root.get().strip()
        data = self.var_data.get().strip()

        if not root or not os.path.isdir(root):
            messagebox.showwarning("Atenção", "Selecione uma pasta ROOT válida.")
            return False

        if len(data) != 6 or not data.isdigit() or not (1 <= int(data[:2]) <= 12):
            messagebox.showwarning("Atenção", "Competência deve ser MMYYYY (ex.: 022026).")
            return False
        return True

    def _on_process(self):
        if self.is_running:
            return
        if not self._validate():
            return

        self.is_running = True
        self.btn_process.config(state="disabled")
        self.progress.config(mode="indeterminate")
        self.progress.start(12)
        self.var_status.set("Processando…")

        # Captura configurações
        config = {
            "root": self.var_root.get().strip(),
            "data": self.var_data.get().strip(),
            "date_envio": self.var_date_envio.get().strip(),
            "reply_all": bool(self.var_reply_all.get()),
            "cc": self.var_cc.get().strip(),
            "my_email": self.var_my_email.get().strip(),
            "open_modal": bool(self.var_open_modal.get()),
            "stop_after_open": bool(self.var_stop_after_open.get()),
            "preview_excel": bool(self.var_preview.get()),
            "pause_before_reply": bool(self.var_pause_before.get()),
            "preview_readonly": bool(self.var_preview_readonly.get()),
        }

        self._log("============================================")
        self._log("Iniciando processamento com as seguintes opções:")
        self._log(f"ROOT: {config['root']}")
        self._log(f"DATA (MMYYYY): {config['data']}")
        self._log(f"Data de Envio: {config['date_envio']}")
        self._log(f"ReplyAll: {config['reply_all']} | CC adicionais: {config['cc'] or '-'}")
        self._log(f"Meu e-mail: {config['my_email'] or '-'}")
        self._log(f"Janela MODAL: {config['open_modal']} | Encerrar após 1ª resposta: {config['stop_after_open']}")
        self._log(f"Pré-visualizar Excel: {config['preview_excel']} | Pausar antes da resposta: {config['pause_before_reply']}")
        self._log(f"Preview SOMENTE LEITURA: {config['preview_readonly']}")

        # Thread de trabalho
        self.worker_thread = threading.Thread(target=self._worker, args=(config,), daemon=True)
        self.worker_thread.start()

        # Checar a thread periodicamente
        self.after(200, self._check_thread)

    def _worker(self, config: dict):
        """
        Roda o main() do core, ajustando as variáveis globais antes.
        COM precisa ser inicializado nesta thread.
        """
        pythoncom.CoInitialize()
        try:
            # (Re)importa o core para garantir que estamos com a versão atual
            importlib.reload(core)

            # Aplicar configurações ao core
            core.ROOT = config["root"]
            core.INBOX_DIR     = os.path.join(core.ROOT, "input")
            core.PROCESSED_DIR = os.path.join(core.ROOT, "processed")
            core.ERROR_DIR     = os.path.join(core.ROOT, "error")
            core.TEMP_DIR      = os.path.join(core.ROOT, "temp")
            core.OUTPUT_DIR    = os.path.join(core.ROOT, "output")

            core.DATA = config["data"]
            core.REPLY_ALL = config["reply_all"]
            core.CC_EMAILS = config["cc"]
            core.MY_EMAIL = config["my_email"]
            core.OPEN_REPLY_MODAL = config["open_modal"]
            core.STOP_AFTER_OPENING_REPLY = config["stop_after_open"]

            # Novas flags
            core.PREVIEW_EXCEL = config["preview_excel"]
            core.PAUSE_BEFORE_REPLY = config["pause_before_reply"]
            core.PREVIEW_READONLY = config["preview_readonly"]
            
            # Se vai visualizar o Excel, automaticamente ativa a pausa para o usuário clicar em "Continuar"
            if core.PREVIEW_EXCEL and not core.PAUSE_BEFORE_REPLY:
                core.PAUSE_BEFORE_REPLY = True
            
            # Data de envio
            core.DATE_ENVIO = config["date_envio"]

            # Log informativo
            self._log(f"[INFO] Pastas: INBOX={core.INBOX_DIR} | PROCESSED={core.PROCESSED_DIR} | OUTPUT={core.OUTPUT_DIR}")
            core.ensure_dirs()

            # Executa
            core.main()
            self._log("[OK] Execução finalizada.")

        except Exception:
            err = traceback.format_exc()
            self._log("❌ ERRO durante o processamento:")
            self._log(err)
        finally:
            pythoncom.CoUninitialize()

    def _check_thread(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.after(250, self._check_thread)
            return

        # Finalizou
        self.is_running = False
        self.btn_process.config(state="normal")
        self.progress.stop()
        self.progress.config(mode="determinate", value=0)
        self.var_status.set("Pronto.")
        self._log("============================================")


if __name__ == "__main__":
    app = FaturasB3GUI()
    app.mainloop()