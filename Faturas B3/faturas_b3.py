# -*- coding: utf-8 -*-
r"""
AUTOMAÇÃO COMPLETA FATURAS B3
-----------------------------
✔ Lê .msg via Outlook (COM) usando OpenSharedItem (MAPI)
✔ Extrai anexos PDF
✔ Processa com pdfplumber
✔ Gera Excel
✔ (NOVO) Pré-visualiza Excel (COM Excel/app padrão), com opção "Somente leitura"
✔ (NOVO) Pode PAUSAR para você revisar/editar/salvar antes de abrir a resposta
✔ (NOVO) Espera e RE-LE a versão salva para anexar exatamente o arquivo editado
✔ (NOVO) Fallback: se Reply/ReplyAll falhar, cria novo e-mail com TO/CC/Assunto/Anexo
✔ Responde o e-mail original, preserva CC
✔ Move .msg para processed/error com nome único
✔ Localiza automaticamente o Base3.xlsx (ou falha com mensagem clara)
"""

import os
import re
import sys
import time
import traceback
import shutil
from datetime import datetime
from typing import List, Tuple
import pandas as pd
import numpy as np
import pdfplumber

# ========= Dependência para Outlook =========
try:
    import win32com.client as win32
except Exception:
    print("Erro: pywin32 não está instalado. Execute:  python -m pip install pywin32")
    raise

# ============================================================
# LOCALIZAÇÃO PADRÃO (funciona em .py e em .exe do PyInstaller)
# ============================================================
def _app_dir():
    # Quando empacotado (PyInstaller onefile)
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.dirname(sys.executable)
    # Em desenvolvimento
    return os.path.dirname(os.path.abspath(__file__))

# ============================================================
# CONFIGURAÇÕES (podem ser ajustadas pela UI)
# ============================================================
ROOT = r"C:\Users\matheus.oliveira\OneDrive - BGC Partners, O365 Tenant\BGCG, O365 Tenant\DMA - Documentos\Accounts Payable\Classificação Fatura B3\Teste"
INBOX_DIR     = os.path.join(ROOT, "input")
PROCESSED_DIR = os.path.join(ROOT, "processed")
ERROR_DIR     = os.path.join(ROOT, "error")
TEMP_DIR      = os.path.join(ROOT, "temp")
OUTPUT_DIR    = os.path.join(ROOT, "output")

BASE_XLSX = os.path.join(ROOT, "Base3.xlsx")  # o script também tenta achar em outros lugares

# Competência (MMYYYY) -> "01/MM/YYYY" será escrito na coluna MÊS
DATA = "022026"

# Responder para todos (ReplyAll) = True; só para o remetente (Reply) = False
REPLY_ALL = True

# CC adicionais (além dos CC do e-mail original)
CC_EMAILS = ""  # ex.: "leticia.silva@empresa.com; financeiro@empresa.com"

# Opcional: remover seu próprio e-mail das listas (evita loops)
MY_EMAIL = ""  # ex.: "matheus.oliveira@empresa.com"

# —— Comportamento da janela e do fluxo —— :
OPEN_REPLY_MODAL = False               # False = janela NÃO modal (recomendado)
STOP_AFTER_OPENING_REPLY = True        # True = encerra o script após abrir a resposta

# —— NOVO: Pré-visualização do Excel —— :
PREVIEW_EXCEL = True             # True = abre o arquivo Excel gerado para conferência
PREVIEW_READONLY = True          # True = abre como SOMENTE LEITURA (desmarque para editar e salvar)
PREVIEW_AUTOFIT = True           # Ajusta largura das colunas automaticamente (COM Excel)
PAUSE_BEFORE_REPLY = False       # True = pausa e pede confirmação antes de abrir a resposta

# ============================================================
# UTILITÁRIOS
# ============================================================
def ensure_dirs():
    for d in [INBOX_DIR, PROCESSED_DIR, ERROR_DIR, TEMP_DIR, OUTPUT_DIR]:
        os.makedirs(d, exist_ok=True)

def safe_move(src_path: str, dst_dir: str) -> str:
    """Move com nome único (evita erro de arquivo já existente)."""
    os.makedirs(dst_dir, exist_ok=True)
    base = os.path.basename(src_path)
    name, ext = os.path.splitext(base)
    target = os.path.join(dst_dir, base)
    i = 1
    while os.path.exists(target):
        target = os.path.join(dst_dir, f"{name} ({i}){ext}")
        i += 1
    shutil.move(src_path, target)
    return target

def mes_competencia_str(data_str: str) -> str:
    mm = data_str[:2]
    yyyy = data_str[2:6]
    return f"01/{mm}/{yyyy}"

def to_float(valor):
    try:
        if isinstance(valor, str):
            return float(valor.replace('.', '').replace(',', '.'))
        return float(valor)
    except Exception:
        return np.nan

def resolve_base_xlsx() -> str:
    """
    Tenta localizar o Base3.xlsx:
    1) BASE_XLSX configurado
    2) mesma pasta do script
    3) ROOT\\Base3.xlsx
    4) busca recursiva dentro de ROOT
    """
    candidates = []
    if BASE_XLSX:
        candidates.append(BASE_XLSX)

    # mesma pasta do script/app
    try:
        script_dir = _app_dir()
        candidates.append(os.path.join(script_dir, "Base3.xlsx"))
    except Exception:
        pass

    # raiz
    candidates.append(os.path.join(ROOT, "Base3.xlsx"))

    for c in candidates:
        if c and os.path.isfile(c):
            print(f"[INFO] Base encontrada: {c}")
            return c

    # busca recursiva
    for dirpath, _, filenames in os.walk(ROOT):
        if "Base3.xlsx" in filenames:
            found = os.path.join(dirpath, "Base3.xlsx")
            print(f"[INFO] Base encontrada (busca): {found}")
            return found

    # falha com instrução clara
    msg = (
        "Base3.xlsx não encontrada.\n"
        f"Tente uma destas opções:\n"
        f"  1) Coloque o arquivo em: {os.path.join(ROOT, 'Base3.xlsx')}\n"
        f"  2) Coloque o arquivo na mesma pasta do faturas_b3.py\n"
        f"  3) Ajuste a variável BASE_XLSX no topo do script para o caminho exato do seu Base3.xlsx\n"
    )
    raise FileNotFoundError(msg)

def build_resumo_html(discrepancias: pd.DataFrame) -> str:
    if discrepancias is None or discrepancias.empty:
        return ""
    items = []
    for _, row in discrepancias.iterrows():
        try:
            desc = str(row.get("DESCRIÇÃO", ""))
            val = float(row.get("VALOR TOTAL", 0) or 0)
            vmax = float(row.get("Valor maximo", 0) or 0)
            ndoc = str(row.get("Num Doc", ""))
            val_br = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            vmax_br = f"{vmax:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            items.append(f"<li>{desc} | Fatura: R$ {val_br} | Máx: R$ {vmax_br} | Doc: {ndoc}</li>")
        except Exception:
            continue
    return "<p><b>Resumo:</b> Foram encontradas as seguintes discrepâncias:</p><ul>" + "\n".join(items) + "</ul>"

def split_recipients(s: str) -> List[str]:
    """Divide por ';' ou ',' e remove vazios/espacos."""
    if not s:
        return []
    parts = []
    for chunk in re.split(r"[;,]", s):
        c = chunk.strip()
        if c:
            parts.append(c)
    return parts

def unique_semicolon(emails: List[str]) -> str:
    """Deduplica preservando ordem e retorna string separada por '; '."""
    seen = set()
    result = []
    for e in emails:
        e_low = e.lower()
        if e_low not in seen:
            seen.add(e_low)
            result.append(e)
    return "; ".join(result)

def get_my_smtp() -> str:
    """Tenta obter o SMTP do usuário logado no Outlook, para remover de TO/CC em ReplyAll (opcional)."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        addr_entry = session.CurrentUser.AddressEntry
        if addr_entry is None:
            return ""
        try:
            ex_user = addr_entry.GetExchangeUser()
            if ex_user:
                return ex_user.PrimarySmtpAddress or ""
        except Exception:
            # POP/IMAP ou outro tipo
            return addr_entry.Address or ""
    except Exception:
        return ""

def open_msg_mapi(msg_path: str):
    """Abre o .msg vinculado à MAPI para permitir Reply/ReplyAll."""
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.Session  # GetNamespace("MAPI")
    return session.OpenSharedItem(msg_path)

def extract_attachments_from_msg(msg_path: str, out_dir: str) -> Tuple[str, str, List[str]]:
    mail = open_msg_mapi(msg_path)

    # tentar SMTP real do remetente
    sender_email = ""
    try:
        sender = mail.Sender
        if sender is not None:
            try:
                ex_user = sender.GetExchangeUser()
                if ex_user:
                    sender_email = ex_user.PrimarySmtpAddress
            except Exception:
                pass
        if not sender_email:
            sender_email = getattr(mail, "SenderEmailAddress", "") or ""
    except Exception:
        pass

    subject = getattr(mail, "Subject", "") or ""

    os.makedirs(out_dir, exist_ok=True)
    attachments_saved = []
    atts = getattr(mail, "Attachments", None)
    if atts and atts.Count > 0:
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fname = att.FileName or f"anexo_{i}.bin"
            fname = re.sub(r'[\\/:*?"<>|]', '_', fname)
            save_path = os.path.join(out_dir, fname)
            att.SaveAsFile(save_path)
            if save_path.lower().endswith(".pdf"):
                attachments_saved.append(save_path)

    print(f"[INFO] Remetente: {sender_email or 'desconhecido'} | Assunto: {subject}")
    print(f"[INFO] PDFs extraídos: {[os.path.basename(p) for p in attachments_saved]}")
    return sender_email, subject, attachments_saved

def extract_tables_from_pdf(pdf_path: str) -> List[pd.DataFrame]:
    tabelas = []
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            for tabela in pagina.extract_tables():
                df = pd.DataFrame(tabela)
                if df.shape[1] >= 3 and df.shape[0] > 3:
                    tabelas.append(df)
    return tabelas

def process_pdf(pdf_path: str, base_df: pd.DataFrame, data_str: str) -> pd.DataFrame:
    tabelas = extract_tables_from_pdf(pdf_path)
    if not tabelas:
        print(f"[WARN] Nenhuma tabela encontrada em {os.path.basename(pdf_path)}")
        return pd.DataFrame()

    fatura = tabelas[0].copy()

    # extrai Nº do documento
    num_doc = "N/A"
    try:
        for _, row in fatura.iterrows():
            for cell in row:
                if cell and "Nº do documento" in str(cell):
                    num_doc = (
                        str(cell)
                        .replace("Nº do documento", "")
                        .replace("\r", "")
                        .replace("\n", "")
                        .strip()
                    )
                    raise StopIteration
    except StopIteration:
        pass

    # limpeza e cabeçalho
    fatura = fatura.dropna(how='all')
    if fatura.shape[0] >= 4:
        fatura.columns = fatura.iloc[3]
    else:
        fatura.columns = fatura.iloc[0]
    target_cols = ['DESCRIÇÃO', 'QTDE', 'VALOR TOTAL']
    fatura = fatura.rename(columns={str(c).strip(): str(c).strip() for c in fatura.columns})
    fatura = fatura[[c for c in target_cols if c in fatura.columns]]
    fatura = fatura.drop([0, 1, 2, 3], errors='ignore')
    fatura = fatura.replace('', np.nan)
    fatura = fatura.dropna(how='any')

    # segunda tabela (se houver)
    if len(tabelas) > 1:
        f2 = tabelas[1].copy()
        try:
            rename_map = {}
            if f2.shape[1] > 1:
                rename_map[f2.columns[1]] = 'DESCRIÇÃO'
            if f2.shape[1] > 5:
                rename_map[f2.columns[5]] = 'QTDE'
            if f2.shape[1] > 6:
                rename_map[f2.columns[6]] = 'VALOR TOTAL'
            f2 = f2.rename(columns=rename_map)
            f2 = f2[['DESCRIÇÃO', 'QTDE', 'VALOR TOTAL']]
            f2 = f2.replace({'': np.nan, '-': np.nan})
            f2 = f2.dropna(how='any')
            fatura = pd.concat([fatura, f2], ignore_index=True)
        except Exception as e:
            print(f"[WARN] 2ª tabela ignorada ({os.path.basename(pdf_path)}): {e}")

    # limpeza de texto
    for col in ["DESCRIÇÃO", "QTDE", "VALOR TOTAL"]:
        if col in fatura.columns:
            fatura[col] = fatura[col].astype(str).str.replace("\r|\n", " ", regex=True)

    # merge com base
    fatura = pd.merge(fatura, base_df, how="left")
    fatura["Num Doc"] = num_doc
    fatura["MÊS"] = mes_competencia_str(data_str)
    # ordena colunas se existirem
    col_order = ["MÊS", "DESCRIÇÃO", "QTDE", "VALOR TOTAL", "Grupo", "Quem aprova", "Descrição do Serviço", "Observação", "Num Doc"]
    cols = [c for c in col_order if c in fatura.columns] + [c for c in fatura.columns if c not in col_order]
    fatura = fatura[cols]

    return fatura

# —— NOVO: Preview do Excel ——
def preview_excel(xlsx_path: str):
    """
    Abre o arquivo Excel para pré-visualização.
    Tenta via COM (Excel) e, se falhar, usa 'os.startfile' como fallback.
    - Respeita PREVIEW_READONLY e PREVIEW_AUTOFIT.
    - NÃO fecha o Excel: usuário fecha quando quiser.
    """
    try:
        import win32com.client as win32_local
    except Exception:
        # Fallback: abre com o aplicativo padrão (geralmente Excel)
        try:
            os.startfile(xlsx_path)
            print("[INFO] Pré-visualização do Excel aberta via aplicativo padrão.")
        except Exception as e:
            print(f"[WARN] Não foi possível abrir o Excel: {e}")
        return

    try:
        excel = win32_local.DispatchEx("Excel.Application")  # nova instância
        excel.Visible = True
        # ReadOnly=PREVIEW_READONLY (True=Só leitura; False=Edição)
        wb = excel.Workbooks.Open(xlsx_path, ReadOnly=bool(PREVIEW_READONLY))
        # Ajustes de visualização
        try:
            for ws in wb.Worksheets:
                ws.Activate()
                ws.Cells(1, 1).Select()
                if PREVIEW_AUTOFIT:
                    used = ws.UsedRange
                    if used is not None:
                        used.Columns.AutoFit()
        except Exception:
            pass

        print("[INFO] Excel aberto para pré-visualização.")
        # Retorna o objeto Workbook (COM) para que o chamador possa monitorar salvamento/fechamento
        return wb
    except Exception as e:
        print(f"[WARN] Falha ao abrir Excel via COM: {e}")
        # Tenta fallback
        try:
            os.startfile(xlsx_path)
            print("[INFO] Pré-visualização do Excel aberta via aplicativo padrão.")
        except Exception as e2:
            print(f"[ERRO] Não foi possível abrir o Excel: {e2}")
        return None
# —— NOVO: Diálogo de confirmação (Windows nativo) ——
def confirm_dialog(title: str, text: str) -> bool:
    """
    Mostra uma caixa de diálogo nativa do Windows com OK/Cancelar,
    sempre em primeiro plano.
    """
    try:
        import ctypes
        MB_OKCANCEL = 0x00000001
        MB_ICONQUESTION = 0x00000020
        MB_TOPMOST = 0x00040000
        MB_SETFOREGROUND = 0x00010000
        flags = MB_OKCANCEL | MB_ICONQUESTION | MB_TOPMOST | MB_SETFOREGROUND
        res = ctypes.windll.user32.MessageBoxW(0, text, title, flags)
        return res == 1  # IDOK
    except Exception:
        return True

# ============================================================
# FLUXO PRINCIPAL
# ============================================================
def process_one_email(msg_path: str) -> bool:
    """
    Processa 1 e-mail (.msg). Retorna True se abriu a janela de resposta com sucesso
    (para o main encerrar, se configurado).
    """
    print(f"\n=== Processando: {os.path.basename(msg_path)} ===")
    temp_dir = os.path.join(TEMP_DIR, os.path.splitext(os.path.basename(msg_path))[0])
    os.makedirs(temp_dir, exist_ok=True)

    try:
        sender, subject, pdfs = extract_attachments_from_msg(msg_path, temp_dir)
        if not pdfs:
            raise RuntimeError("Nenhum PDF encontrado neste e-mail.")

        base_path = resolve_base_xlsx()
        base_df = pd.read_excel(base_path, engine="openpyxl")

        all_rows = []
        for pdf in pdfs:
            print(f"[INFO] Processando PDF: {os.path.basename(pdf)}")
            df = process_pdf(pdf, base_df, DATA)
            if not df.empty:
                all_rows.append(df)

        if not all_rows:
            raise RuntimeError("Nenhuma linha válida foi extraída dos PDFs.")

        series = pd.concat(all_rows, ignore_index=True)

        # Conversões e discrepâncias
        if "VALOR TOTAL" in series.columns:
            series["VALOR TOTAL"] = series["VALOR TOTAL"].apply(to_float)
        if "Valor maximo" in series.columns:
            series["Valor maximo"] = series["Valor maximo"].apply(to_float)
            series["Excede Limite"] = series["VALOR TOTAL"] > series["Valor maximo"]
            discrep = series[series.get("Excede Limite", False) == True].copy()
        else:
            discrep = pd.DataFrame()

        # Exporta removendo colunas auxiliares
        export_df = series.drop(columns=["Valor maximo", "Excede Limite"], errors="ignore")

        mm = DATA[:2]; yyyy = DATA[2:6]
        mes_ano = f"{mm}/{yyyy}"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"Fatura_B3_{mm}{yyyy}_{ts}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        export_df.to_excel(out_path, index=False, engine="openpyxl")
        print(f"[INFO] Excel gerado: {out_path}")

        created_mtime = os.path.getmtime(out_path)

        # —— PRÉ-VISUALIZAÇÃO + PAUSA (se habilitado) —— #
        if PREVIEW_EXCEL:
            # Abre o Excel e, quando possível, retorna o objeto Workbook (COM).
            wb = preview_excel(out_path)

            # Se o fluxo pedir pausa explícita, pergunta ao usuário (manteve comportamento anterior).
            if PAUSE_BEFORE_REPLY:
                ok = confirm_dialog(
                    "Pré-visualização do Excel",
                    "Revise e, se necessário, edite o Excel gerado.\n"
                    "Salve suas alterações no MESMO arquivo (Ctrl+S).\n\n"
                    "Clique OK para continuar e abrir a resposta,\n"
                    "ou Cancelar para abortar este e-mail."
                )
                if not ok:
                    try:
                        dst = safe_move(msg_path, ERROR_DIR)
                        print(f"[INFO] Operação cancelada pelo usuário. E-mail movido para error: {dst}")
                    except Exception as e_move:
                        print(f"[WARN] Falha ao mover para error após cancelamento: {e_move}")
                    return False

            # Se o arquivo foi aberto via COM e está editável, aguarda o usuário salvar/fechar.
            if wb is not None and not PREVIEW_READONLY:
                # Espera até timeout o arquivo ser salvo (mtime > criado) ou o Workbook reportar Saved=True
                def _mtime(p):
                    try:
                        return os.path.getmtime(p)
                    except Exception:
                        return created_mtime

                deadline = time.time() + 120
                last_err = None
                while time.time() < deadline:
                    try:
                        # Se o workbook foi salvo por Ctrl+S, wb.Saved será True only after save; check file mtime too.
                        saved_flag = False
                        try:
                            saved_flag = bool(getattr(wb, 'Saved', False))
                        except Exception:
                            # acessando wb pode falhar se o usuário fechou o Excel
                            saved_flag = False

                        if _mtime(out_path) > created_mtime or saved_flag:
                            # Tenta reler brevemente para atualizar dados
                            try:
                                wb_dict = pd.read_excel(out_path, sheet_name=None, engine="openpyxl")
                                export_reloaded = pd.concat(wb_dict.values(), ignore_index=True)
                                merged = pd.merge(export_reloaded, base_df, how="left")
                                if "VALOR TOTAL" in merged.columns and "Valor maximo" in merged.columns:
                                    merged["VALOR TOTAL"] = merged["VALOR TOTAL"].apply(to_float)
                                    merged["Valor maximo"] = merged["Valor maximo"].apply(to_float)
                                    merged["Excede Limite"] = merged["VALOR TOTAL"] > merged["Valor maximo"]
                                    discrep = merged[merged["Excede Limite"] == True].copy()
                                else:
                                    print("[WARN] Não foi possível recalcular discrepâncias a partir do Excel salvo (colunas ausentes).")
                                print("[INFO] Excel salvo pelo usuário foi relido com sucesso; resumo atualizado.")
                                last_err = None
                                break
                            except Exception as e_reload:
                                last_err = e_reload
                                # Se não conseguiu reler imediatamente, espera um pouco e tenta de novo
                                time.sleep(1.0)
                                continue

                    except Exception as e_loop:
                        last_err = e_loop
                    time.sleep(0.5)

                if last_err:
                    print(f"[WARN] Não foi possível reler o Excel salvo após esperar: {last_err}")

        # ==== MONTA RESPOSTA (com fallback) ====
        try:
            original = open_msg_mapi(msg_path)
        except Exception as e_open:
            print(f"[WARN] Falha ao reabrir o .msg via MAPI para responder: {e_open}")
            original = None

        reply = None
        try:
            if original is not None:
                reply = original.ReplyAll() if REPLY_ALL else original.Reply()
        except Exception as e_rep:
            print(f"[WARN] Reply/ReplyAll falhou: {e_rep}")
            reply = None

        # Extrai destinatários originais (para o fallback também)
        try:
            subject = getattr(original, "Subject", "") if original else ""
            orig_to_list = split_recipients(getattr(original, "To", "") if original else "")
            orig_cc_list = split_recipients(getattr(original, "CC", "") if original else "")
        except Exception:
            subject, orig_to_list, orig_cc_list = "", [], []

        my_email = (MY_EMAIL or "").strip() or get_my_smtp().strip()
        if my_email:
            orig_to_list = [e for e in orig_to_list if e.lower() != my_email.lower()]
            orig_cc_list = [e for e in orig_cc_list if e.lower() != my_email.lower()]

        intro_html = f"<p>Olá,</p><p>Segue o Excel consolidado das faturas B3 referentes a {mes_ano}.</p>"
        resumo_html = build_resumo_html(discrep)

        # Se Reply/ReplyAll falhar, cria um novo e-mail zerado com TO/CC/Subject/Body/Anexo
        if reply is None:
            try:
                outlook = win32.Dispatch("Outlook.Application")
                reply = outlook.CreateItem(0)  # 0 = olMailItem
                if subject:
                    reply.Subject = f"Re: {subject}"
                reply.To = unique_semicolon(orig_to_list)
                extra_cc_list = split_recipients(CC_EMAILS)
                reply.CC = unique_semicolon(orig_cc_list + extra_cc_list)
                reply.HTMLBody = intro_html + resumo_html
                reply.Attachments.Add(out_path)
                print("[INFO] Fallback acionado: criado novo e-mail (sem Reply/ReplyAll).")
            except Exception as e_new:
                print(f"❌ ERRO ao criar e-mail (fallback): {e_new}")
                raise
        else:
            # Ajusta TO/CC para o caso de Reply que venha vazio e insere corpo
            try:
                reply_cc_existing = split_recipients(getattr(reply, "CC", ""))
                extra_cc_list = split_recipients(CC_EMAILS)
                final_cc = reply_cc_existing + orig_cc_list + extra_cc_list
                reply.CC = unique_semicolon(final_cc)

                reply_to_existing = split_recipients(getattr(reply, "To", ""))
                if not reply_to_existing and orig_to_list:
                    reply.To = unique_semicolon(orig_to_list)

                reply.HTMLBody = intro_html + resumo_html + reply.HTMLBody
                reply.Attachments.Add(out_path)
            except Exception as e_set:
                print(f"[WARN] Ajustes no e-mail de resposta falharam, tentando fallback: {e_set}")
                # fallback “duro” — cria outro do zero
                outlook = win32.Dispatch("Outlook.Application")
                newmail = outlook.CreateItem(0)
                if subject:
                    newmail.Subject = f"Re: {subject}"
                newmail.To = unique_semicolon(orig_to_list)
                extra_cc_list = split_recipients(CC_EMAILS)
                newmail.CC = unique_semicolon(orig_cc_list + extra_cc_list)
                newmail.HTMLBody = intro_html + resumo_html
                newmail.Attachments.Add(out_path)
                reply = newmail

        # Exibe a janela e garante foco/primeiro plano
        try:
            reply.Display(OPEN_REPLY_MODAL)  # False = NÃO modal
            try:
                insp = reply.GetInspector
                if insp is not None:
                    insp.Activate()  # traz para frente
            except Exception:
                pass
            print("[OK] Janela de resposta aberta (envio manual).")
        except Exception as e_disp:
            print(f"❌ ERRO ao exibir a janela de e-mail: {e_disp}")
            raise

        # Mover o .msg para processed (nome único)
        dst = safe_move(msg_path, PROCESSED_DIR)
        print(f"[OK] Arquivado em: {dst}")

        return True

    except Exception as e:
        print("❌ ERRO:", e)
        traceback.print_exc()
        try:
            dst = safe_move(msg_path, ERROR_DIR)
            print(f"[INFO] Movido para error: {dst}")
        except Exception as e2:
            print(f"[WARN] Falha ao mover para error: {e2}")
        return False
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    ensure_dirs()
    if not os.path.isdir(INBOX_DIR):
        print("Pasta INBOX não encontrada:", INBOX_DIR)
        return

    files = [f for f in os.listdir(INBOX_DIR) if f.lower().endswith(".msg")]
    if not files:
        print("Nenhum e-mail (.msg) encontrado em:", INBOX_DIR)
        return

    print(f"[INFO] Encontrados {len(files)} arquivo(s) .msg para processar.")
    for i, f in enumerate(files, start=1):
        print(f"[INFO] ({i}/{len(files)}) {f}")
        opened = process_one_email(os.path.join(INBOX_DIR, f))
        # Encerrar após abrir a primeira resposta (se configurado)
        if opened and STOP_AFTER_OPENING_REPLY:
            print("✔ Resposta aberta. Encerrando o script para liberar o Outlook.")
            return

# ===== NOVO: Responder com Excel EXISTENTE (sem reprocessar PDFs) =====
def reply_with_existing(xlsx_path: str, msg_path: str) -> bool:
    """
    Abre uma resposta (ou cria novo e-mail, fallback) para o .msg informado,
    anexando exatamente o arquivo Excel existente (xlsx_path).
    Recalcula o resumo (discrepâncias) a partir do conteúdo do Excel + Base3.
    """
    try:
        ensure_dirs()
        if not os.path.isfile(xlsx_path):
            raise FileNotFoundError(f"Excel não encontrado: {xlsx_path}")
        if not os.path.isfile(msg_path):
            raise FileNotFoundError(f".msg não encontrado: {msg_path}")

        base_path = resolve_base_xlsx()
        base_df = pd.read_excel(base_path, engine="openpyxl")

        # Lê o Excel existente (todas as planilhas)
        wb_dict = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
        df_all = pd.concat(wb_dict.values(), ignore_index=True)

        # Se habilitado preview, abre o Excel para possível edição e aguarda salvamento quando aplicável
        if PREVIEW_EXCEL:
            wb = preview_excel(xlsx_path)
            if PAUSE_BEFORE_REPLY:
                ok = confirm_dialog(
                    "Pré-visualização do Excel",
                    "Revise e, se necessário, edite o Excel selecionado.\n"
                    "Salve suas alterações no MESMO arquivo (Ctrl+S).\n\n"
                    "Clique OK para continuar e abrir a resposta,\n"
                    "ou Cancelar para abortar."
                )
                if not ok:
                    print("[INFO] Operação cancelada pelo usuário (reply_with_existing).")
                    return False

            if wb is not None and not PREVIEW_READONLY:
                created_mtime = os.path.getmtime(xlsx_path)
                def _mtime(p):
                    try:
                        return os.path.getmtime(p)
                    except Exception:
                        return created_mtime

                deadline = time.time() + 120
                last_err = None
                while time.time() < deadline:
                    try:
                        saved_flag = False
                        try:
                            saved_flag = bool(getattr(wb, 'Saved', False))
                        except Exception:
                            saved_flag = False
                        if _mtime(xlsx_path) > created_mtime or saved_flag:
                            try:
                                wb_dict = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
                                df_all = pd.concat(wb_dict.values(), ignore_index=True)
                                # recompute discrep after edit
                                merged = pd.merge(df_all, base_df, how="left")
                                if "VALOR TOTAL" in merged.columns and "Valor maximo" in merged.columns:
                                    merged["VALOR TOTAL"] = merged["VALOR TOTAL"].apply(to_float)
                                    merged["Valor maximo"] = merged["Valor maximo"].apply(to_float)
                                    merged["Excede Limite"] = merged["VALOR TOTAL"] > merged["Valor maximo"]
                                    discrep = merged[merged["Excede Limite"] == True].copy()
                                print("[INFO] Excel editado foi relido com sucesso; resumo atualizado.")
                                last_err = None
                                break
                            except Exception as e_reload:
                                last_err = e_reload
                                time.sleep(1.0)
                                continue
                    except Exception as e_loop:
                        last_err = e_loop
                    time.sleep(0.5)
                if last_err:
                    print(f"[WARN] Não foi possível reler o Excel salvo após esperar: {last_err}")

        # Tenta recompor discrepâncias com base
        discrep = pd.DataFrame()
        try:
            merged = pd.merge(df_all, base_df, how="left")
            if "VALOR TOTAL" in merged.columns and "Valor maximo" in merged.columns:
                merged["VALOR TOTAL"] = merged["VALOR TOTAL"].apply(to_float)
                merged["Valor maximo"] = merged["Valor maximo"].apply(to_float)
                merged["Excede Limite"] = merged["VALOR TOTAL"] > merged["Valor maximo"]
                discrep = merged[merged["Excede Limite"] == True].copy()
        except Exception as e:
            print(f"[WARN] Falha ao calcular discrepâncias a partir do Excel existente: {e}")

        mm = DATA[:2]; yyyy = DATA[2:6]
        mes_ano = f"{mm}/{yyyy}"

        # ==== MONTA RESPOSTA (reuso do bloco com fallback) ====
        try:
            original = open_msg_mapi(msg_path)
        except Exception as e_open:
            print(f"[WARN] Falha ao abrir o .msg via MAPI: {e_open}")
            original = None

        reply = None
        try:
            if original is not None:
                reply = original.ReplyAll() if REPLY_ALL else original.Reply()
        except Exception as e_rep:
            print(f"[WARN] Reply/ReplyAll falhou: {e_rep}")
            reply = None

        # Extrai destinatários originais
        try:
            subject = getattr(original, "Subject", "") if original else ""
            orig_to_list = split_recipients(getattr(original, "To", "") if original else "")
            orig_cc_list = split_recipients(getattr(original, "CC", "") if original else "")
        except Exception:
            subject, orig_to_list, orig_cc_list = "", [], []

        my_email = (MY_EMAIL or "").strip() or get_my_smtp().strip()
        if my_email:
            orig_to_list = [e for e in orig_to_list if e.lower() != my_email.lower()]
            orig_cc_list = [e for e in orig_cc_list if e.lower() != my_email.lower()]

        intro_html = f"<p>Olá,</p><p>Segue o Excel consolidado das faturas B3 referentes a {mes_ano}.</p>"
        resumo_html = build_resumo_html(discrep)

        if reply is None:
            # Fallback: cria novo e-mail
            outlook = win32.Dispatch("Outlook.Application")
            reply = outlook.CreateItem(0)
            if subject:
                reply.Subject = f"Re: {subject}"
            reply.To = unique_semicolon(orig_to_list)
            extra_cc_list = split_recipients(CC_EMAILS)
            reply.CC = unique_semicolon(orig_cc_list + extra_cc_list)
            reply.HTMLBody = intro_html + resumo_html
            reply.Attachments.Add(xlsx_path)
        else:
            # Ajusta TO/CC se necessário e insere corpo
            reply_cc_existing = split_recipients(getattr(reply, "CC", ""))
            extra_cc_list = split_recipients(CC_EMAILS)
            final_cc = reply_cc_existing + orig_cc_list + extra_cc_list
            reply.CC = unique_semicolon(final_cc)
            reply_to_existing = split_recipients(getattr(reply, "To", ""))
            if not reply_to_existing and orig_to_list:
                reply.To = unique_semicolon(orig_to_list)
            reply.HTMLBody = intro_html + resumo_html + reply.HTMLBody
            reply.Attachments.Add(xlsx_path)

        # Mostra janela
        reply.Display(OPEN_REPLY_MODAL)
        try:
            insp = reply.GetInspector
            if insp is not None:
                insp.Activate()
        except Exception:
            pass

        print("[OK] Janela de resposta aberta (com Excel EXISTENTE).")
        return True

    except Exception as e:
        print("❌ ERRO no reply_with_existing:", e)
        traceback.print_exc()
        return False


if __name__ == "__main__":
    main()