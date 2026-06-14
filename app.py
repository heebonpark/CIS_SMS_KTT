# =====================================================================
# Copyright (c) 2026 세은아빠 (bough38@naver.com)
# All rights reserved.
# 본 프로그램의 저작권은 '세은아빠'에게 있으며, 개발자의 허락 없는 
# 무단 전재, 변형, 상업적 이용 및 재배포를 엄격히 금지합니다.
# =====================================================================
#
# ┌─────────────────────────────────────────────────────┐
# │  EXE 빌드:                                          │
# │  pip install pyinstaller pandas pyautogui            │
# │  pip install pyperclip xlwings openpyxl              │
# │  cd "D:\util\CIS단체, 개별 선택 문자보내기"           │
# │  pyinstaller --onefile --noconsole app.py            │
# └─────────────────────────────────────────────────────┘

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinter import font as tkfont
import pandas as pd
import pyautogui
import pyperclip
import time
import threading
import xlwings as xw
import os, sys, json, base64, socket, hashlib, platform, subprocess
from datetime import datetime, date

pyautogui.FAILSAFE = True

# =========================================================
#  OS 구분 및 단축키 설정
# =========================================================
IS_MAC = platform.system() == "Darwin"
MODIFIER = "command" if IS_MAC else "ctrl"

# =========================================================
#  디자인 시스템 (Pretendard & Modern Slate Theme)
# =========================================================
MAIN_FONT = "Arial"  # 기본값, init_fonts에서 변경됨

# Colors (Slate / Navy Premium Palette)
BG_MAIN = "#f1f5f9"        # Slate 100
BG_CARD = "#ffffff"        # Card 배경
COLOR_NAVY = "#0f172a"     # Dark Navy
COLOR_PRIMARY = "#2563eb"  # Royal Blue
COLOR_PRIMARY_HOVER = "#1d4ed8"
COLOR_SUCCESS = "#10b981"  # Emerald
COLOR_SUCCESS_HOVER = "#059669"
COLOR_INFO = "#06b6d4"     # Cyan
COLOR_INFO_HOVER = "#0891b2"
COLOR_MUTED = "#64748b"    # Slate 500
COLOR_MUTED_HOVER = "#475569"
COLOR_DANGER = "#ef4444"   # Red 500
COLOR_DANGER_HOVER = "#dc2626"
COLOR_BORDER = "#cbd5e1"   # Slate 300

# Fonts placeholders (init_fonts에서 초기화)
FONT_TITLE = (MAIN_FONT, 16, "bold")
FONT_SUBTITLE = (MAIN_FONT, 11, "bold")
FONT_BOLD_11 = (MAIN_FONT, 11, "bold")
FONT_BOLD_10 = (MAIN_FONT, 10, "bold")
FONT_REG_10 = (MAIN_FONT, 10)
FONT_BOLD_9 = (MAIN_FONT, 9, "bold")
FONT_REG_9 = (MAIN_FONT, 9)

def init_fonts():
    global MAIN_FONT, FONT_TITLE, FONT_SUBTITLE, FONT_BOLD_11, FONT_BOLD_10, FONT_REG_10, FONT_BOLD_9, FONT_REG_9
    try:
        available_fonts = tkfont.families()
    except Exception:
        available_fonts = []
    
    if "Pretendard" in available_fonts:
        MAIN_FONT = "Pretendard"
    elif "Apple SD Gothic Neo" in available_fonts:
        MAIN_FONT = "Apple SD Gothic Neo"
    elif "Segoe UI" in available_fonts:
        MAIN_FONT = "Segoe UI"
    else:
        MAIN_FONT = "Arial"
        
    FONT_TITLE = (MAIN_FONT, 15, "bold")
    FONT_SUBTITLE = (MAIN_FONT, 11, "bold")
    FONT_BOLD_11 = (MAIN_FONT, 11, "bold")
    FONT_BOLD_10 = (MAIN_FONT, 10, "bold")
    FONT_REG_10 = (MAIN_FONT, 10)
    FONT_BOLD_9 = (MAIN_FONT, 9, "bold")
    FONT_REG_9 = (MAIN_FONT, 9)

# =========================================================
#  경로 설정
# =========================================================
def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = get_app_dir()
LOG_FILE        = os.path.join(APP_DIR, "발송로그.txt")
USAGE_LOG_FILE  = os.path.join(APP_DIR, "사용이력.txt")
ADMIN_CONFIG    = os.path.join(APP_DIR, "admin_config.dat")
SESSION_FILE    = os.path.join(APP_DIR, "session.json")       # ✅ 중단점 이어하기
TEMPLATE_FILE   = os.path.join(APP_DIR, "templates.json")     # ✅ 문구 템플릿

ADMIN_PW_HASH = hashlib.sha256("admin0303".encode()).hexdigest()

# ── 전역 변수 ──
registered_indices = set()
failed_indices = set()
is_paused = False
is_running = False
coord_input = coord_msg = coord_phone = coord_send = coord_confirm = None
data = pd.DataFrame(columns=["전화번호", "발송문자"])
start_index = 0
loaded_file_path = ""   # 엑셀 결과 기록용
PASTE_DELAY = 0.3
REG_DELAY_INPUT = 0.05
REG_DELAY_TAB = 0.1

MACRO_CONFIG_FILE = os.path.join(APP_DIR, "macro_config.json")

def load_macro_config():
    defaults = {
        "번호등록": [
            {"type": "click_type", "name": "받는사람 입력칸", "coord": None, "value_source": "phone"}
        ],
        "문자발송": [
            {"type": "click_type", "name": "① 메시지 입력칸", "coord": None, "value_source": "message"},
            {"type": "click_type", "name": "② 받는사람 입력칸", "coord": None, "value_source": "phone"},
            {"type": "click_only", "name": "③ SMS 발송 버튼", "coord": None},
            {"type": "click_only", "name": "④ 확인 버튼", "coord": None}
        ]
    }
    if not os.path.exists(MACRO_CONFIG_FILE):
        save_macro_config(defaults)
        return defaults
    try:
        with open(MACRO_CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for mode in ["번호등록", "문자발송"]:
            if mode not in cfg or not isinstance(cfg[mode], list):
                cfg[mode] = defaults[mode]
        return cfg
    except:
        return defaults

def save_macro_config(cfg):
    try:
        with open(MACRO_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except:
        pass

macro_config = load_macro_config()

# =========================================================
#  관리자 설정 (base64)
# =========================================================
def load_admin_config():
    defaults = {"user_password":"0303","expire_date":"","expire_message":"사용 기간이 만료되었습니다.\n관리자에게 문의하세요.","created":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"discord_webhook":"","remote_password_url":""}
    if not os.path.exists(ADMIN_CONFIG):
        save_admin_config(defaults); return defaults
    try:
        with open(ADMIN_CONFIG,"r",encoding="utf-8") as f: raw=f.read().strip()
        cfg=json.loads(base64.b64decode(raw).decode("utf-8"))
        for k,v in defaults.items():
            if k not in cfg: cfg[k]=v
        return cfg
    except: return defaults

def save_admin_config(cfg):
    try:
        encoded=base64.b64encode(json.dumps(cfg,ensure_ascii=False,indent=2).encode("utf-8")).decode("utf-8")
        with open(ADMIN_CONFIG,"w",encoding="utf-8") as f: f.write(encoded)
    except: pass

# =========================================================
#  사용 이력
# =========================================================
def get_user_info():
    try: pc=socket.gethostname()
    except: pc="UNKNOWN"
    try: user=os.getlogin()
    except: user=os.environ.get("USERNAME",os.environ.get("USER","UNKNOWN"))
    return pc,user

def send_discord_message(webhook_url, payload_or_content):
    if not webhook_url:
        return
    def worker():
        try:
            import urllib.request
            import json
            if isinstance(payload_or_content, dict):
                data = payload_or_content
            else:
                data = {"content": payload_or_content}
            req = urllib.request.Request(
                webhook_url,
                data=json.dumps(data).encode("utf-8"),
                headers={"Content-Type": "application/json", "User-Agent": "urllib-discord-bot"}
            )
            with urllib.request.urlopen(req, timeout=10) as response:
                pass
        except Exception:
            pass
    threading.Thread(target=worker, daemon=True).start()

def write_usage_log(action,detail=""):
    try:
        pc,user=get_user_info()
        now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{now}] PC:{pc} | 사용자:{user} | {action}" + (f" | {detail}" if detail else "")
        with open(USAGE_LOG_FILE,"a",encoding="utf-8") as f:
            f.write(log_line + "\n")
            
        cfg = load_admin_config()
        webhook_url = cfg.get("discord_webhook", "").strip()
        if webhook_url:
            def discord_worker():
                try:
                    import urllib.request
                    import json
                    import platform
                    
                    # 1. OS Info
                    os_info = f"{platform.system()} {platform.release()}"
                    
                    # 2. Public IP & Location Info (ip-api.com)
                    ip_addr = "Unknown"
                    location_str = "Unknown"
                    try:
                        req_ip = urllib.request.Request(
                            "http://ip-api.com/json/",
                            headers={"User-Agent": "Mozilla/5.0"}
                        )
                        with urllib.request.urlopen(req_ip, timeout=5) as resp:
                            data = json.loads(resp.read().decode("utf-8"))
                            ip_addr = data.get("query", "Unknown")
                            city = data.get("city", "")
                            country = data.get("country", "")
                            org = data.get("org", "Unknown")
                            location_str = f"{city}, {country} ({org})"
                    except Exception:
                        pass
                    
                    # 3. Build Embed Payload
                    payload = {
                        "embeds": [{
                            "title": "🚀 CIS Execution Detected",
                            "color": 0x3498db,  # Blue
                            "fields": [
                                {"name": "작업 (Action)", "value": action, "inline": True},
                                {"name": "상세 (Detail)", "value": detail if detail else "N/A", "inline": True},
                                {"name": "PC / 사용자", "value": f"{pc} / {user}", "inline": False},
                                {"name": "OS", "value": os_info, "inline": True},
                                {"name": "공인 IP (Public IP)", "value": ip_addr, "inline": True},
                                {"name": "위치 (Location)", "value": location_str, "inline": False},
                            ],
                            "footer": {"text": "Real-time Monitoring System"},
                            "timestamp": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
                        }]
                    }
                    send_discord_message(webhook_url, payload)
                except Exception:
                    pass
            threading.Thread(target=discord_worker, daemon=True).start()
    except: pass

def check_expiration(cfg):
    exp_str=cfg.get("expire_date","").strip()
    if not exp_str: return True
    try: exp_date=datetime.strptime(exp_str,"%Y-%m-%d").date()
    except: return True
    remaining=(exp_date-date.today()).days
    if remaining<0:
        write_usage_log("접속차단",f"만료({exp_str})")
        messagebox.showerror("사용 기간 만료",f"{cfg.get('expire_message','')}\n\n만료일: {exp_str}")
        return False
    if remaining<=7:
        messagebox.showwarning("기간 안내",f"사용 기간 {remaining}일 남음\n만료일: {exp_str}")
    return True

# =========================================================
#  ✅ 1. 세션 저장/복원 (중단점 이어하기)
# =========================================================
def save_session():
    """현재 진행 상태를 파일로 저장"""
    try:
        session = {
            "file_path": loaded_file_path,
            "start_index": start_index,
            "registered": sorted(list(registered_indices)),
            "failed": sorted(list(failed_indices)),
            "work_mode": work_mode.get() if work_mode else "번호등록",
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total": len(data),
        }
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(session, f, ensure_ascii=False, indent=2)
    except: pass

def load_session():
    """저장된 세션이 있으면 복원 제안"""
    if not os.path.exists(SESSION_FILE):
        return False
    try:
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            session = json.load(f)
        saved_at = session.get("saved_at", "?")
        total = session.get("total", 0)
        done = len(session.get("registered", []))
        mode = session.get("work_mode", "?")
        fpath = session.get("file_path", "")

        answer = messagebox.askyesno(
            "📂 이전 작업 이어하기",
            f"이전에 중단된 작업이 있습니다.\n\n"
            f"  저장 시점: {saved_at}\n"
            f"  작업 모드: {mode}\n"
            f"  진행 상황: {done} / {total} 건 완료\n"
            f"  파일: {os.path.basename(fpath) if fpath else '활성 엑셀'}\n\n"
            f"이어서 진행하시겠습니까?"
        )
        if not answer:
            clear_session()
            return False

        # 파일 다시 로드
        if fpath and os.path.exists(fpath):
            df_file = pd.read_excel(fpath, engine='openpyxl')
            if df_file.shape[1] >= 3:
                df_file = df_file.iloc[:, [1, 2]]; df_file.columns = ["전화번호", "발송문자"]
            elif df_file.shape[1] >= 2:
                df_file = df_file.iloc[:, [1]]; df_file.columns = ["전화번호"]
            else:
                df_file.columns = ["전화번호"]
            process_and_set_data(df_file, silent=True)
            filename_label.config(text=f"📂 이어하기: {os.path.basename(fpath)}")
        else:
            messagebox.showinfo("안내", "원본 파일을 찾을 수 없습니다.\n데이터를 다시 불러오세요.")
            clear_session()
            return False

        # 상태 복원
        global start_index, registered_indices, failed_indices
        start_index = session.get("start_index", 0)
        registered_indices = set(session.get("registered", []))
        failed_indices = set(session.get("failed", []))
        work_mode.set(session.get("work_mode", "번호등록"))
        toggle_work_mode()
        update_tables()
        update_progress()
        write_usage_log("세션 복원", f"이어하기 {done}/{total}")
        return True
    except:
        clear_session()
        return False

def clear_session():
    try:
        if os.path.exists(SESSION_FILE): os.remove(SESSION_FILE)
    except: pass

# =========================================================
#  ✅ 2. 문구 템플릿 저장/불러오기
# =========================================================
def load_templates():
    if not os.path.exists(TEMPLATE_FILE): return []
    try:
        with open(TEMPLATE_FILE,"r",encoding="utf-8") as f: return json.load(f)
    except: return []

def save_templates(templates):
    try:
        with open(TEMPLATE_FILE,"w",encoding="utf-8") as f:
            json.dump(templates,f,ensure_ascii=False,indent=2)
    except: pass

def add_template():
    """현재 문구를 템플릿으로 저장"""
    msg = fixed_msg_text.get("1.0", tk.END).strip()
    if not msg:
        messagebox.showwarning("빈 문구", "저장할 문구를 입력하세요.")
        return
    name = simpledialog.askstring("템플릿 저장", "템플릿 이름을 입력하세요:")
    if not name: return
    templates = load_templates()
    templates.append({"name": name, "message": msg})
    save_templates(templates)
    refresh_template_dropdown()
    messagebox.showinfo("저장 완료", f"'{name}' 템플릿이 저장되었습니다.")

def apply_template(*args):
    """선택된 템플릿을 문구란에 적용"""
    sel = tpl_var.get()
    if not sel or sel == "-- 템플릿 선택 --": return
    templates = load_templates()
    for t in templates:
        if t["name"] == sel:
            fixed_msg_text.config(state=tk.NORMAL)
            fixed_msg_text.delete("1.0", tk.END)
            fixed_msg_text.insert(tk.END, t["message"])
            if send_mode.get() != "일괄":
                fixed_msg_text.config(state=tk.DISABLED, bg="#e9ecef")
            break

def delete_template():
    """선택된 템플릿 삭제"""
    sel = tpl_var.get()
    if not sel or sel == "-- 템플릿 선택 --": return
    templates = [t for t in load_templates() if t["name"] != sel]
    save_templates(templates)
    refresh_template_dropdown()
    messagebox.showinfo("삭제", f"'{sel}' 템플릿이 삭제되었습니다.")

def refresh_template_dropdown():
    templates = load_templates()
    names = ["-- 템플릿 선택 --"] + [t["name"] for t in templates]
    tpl_dropdown["values"] = names
    tpl_var.set(names[0])

# =========================================================
#  ✅ 3. 엑셀 결과 기록
# =========================================================
def write_result_to_excel(row_idx, status_text):
    """활성 엑셀 시트 D열에 발송 결과 기록 (체크박스 활성 시)"""
    try:
        if not excel_writeback_var.get(): return
        app = xw.apps.active
        if not app: return
        sheet = app.books.active.sheets.active
        excel_row = row_idx + 2  # 1행=헤더, 0-indexed → +2
        now = datetime.now().strftime("%H:%M:%S")
        sheet.range(f'D{excel_row}').value = f"{status_text} ({now})"
    except: pass

# =========================================================
#  로그 / 유틸
# =========================================================
def write_log(status, number, message=""):
    try:
        with open(LOG_FILE,"a",encoding="utf-8") as f:
            pc,user=get_user_info()
            now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{now}] {pc}/{user} | {status} | {number} | {(message or '')[:40]}\n")
    except: pass

def open_path(path):
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.call(["open", path])
        else:
            subprocess.call(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("오류", f"경로를 열 수 없습니다: {e}")

def open_log_file():
    if os.path.exists(LOG_FILE): open_path(LOG_FILE)
    else: messagebox.showinfo("로그","아직 발송 로그가 없습니다.")

def open_log_folder(): open_path(APP_DIR)

def safe_input(text):
    pyperclip.copy(text); pyautogui.hotkey(MODIFIER,"v"); time.sleep(PASTE_DELAY)

# ── 현대식 버튼 헬퍼 (마우스 호버 지원 & OS 호환성 완벽 해결) ──
class LabelButton(tk.Label):
    def __init__(self, parent, text, command, bg, fg="white", hover_bg=None, font=None, width=None, state=tk.NORMAL, **kwargs):
        self.command = command
        self.bg = bg
        self.hover_bg = hover_bg or bg
        self.fg = fg
        super().__init__(
            parent,
            text=text,
            bg=bg,
            fg=fg,
            font=font,
            bd=0,
            relief="flat",
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            cursor="hand2" if state == tk.NORMAL else "arrow",
            state=state,
            padx=12,
            pady=6,
            **kwargs
        )
        if width:
            self.config(width=width)
            
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_click(self, event):
        if self.cget("state") == tk.NORMAL and self.command:
            self.command()

    def _on_enter(self, event):
        if self.cget("state") == tk.NORMAL:
            self.config(bg=self.hover_bg)

    def _on_leave(self, event):
        if self.cget("state") == tk.NORMAL:
            self.config(bg=self.bg)

    def disable(self):
        self.config(state=tk.DISABLED, bg="#cbd5e1", fg="#94a3b8", cursor="arrow")

    def enable(self):
        self.config(state=tk.NORMAL, bg=self.bg, fg=self.fg, cursor="hand2")

def create_btn(parent, text, command, bg_color, hover_color, fg_color="white", font=None, width=None, state=tk.NORMAL):
    if font is None:
        font = FONT_BOLD_10
    return LabelButton(parent, text, command, bg_color, fg_color, hover_color, font, width, state)

def set_btn_state(btn, state, normal_bg, hover_bg, fg="white"):
    if state == tk.DISABLED:
        btn.disable()
    else:
        btn.bg = normal_bg
        btn.hover_bg = hover_bg
        btn.fg = fg
        btn.enable()

# =========================================================
#  좌표 캡처 팝업
# =========================================================
def capture_coord_popup(step_title, step_desc, step_no="", countdown_sec=5):
    messagebox.showinfo(f"📍 좌표 설정 {step_no}".strip(),
        f"{step_title}\n\n{step_desc}\n\n[확인] 후 {countdown_sec}초 카운트다운 시작")
    ov=tk.Toplevel(root); ov.overrideredirect(True); ov.attributes("-topmost",True); ov.attributes("-alpha",0.93); ov.configure(bg="#1a1a2e")
    ow,oh=520,250; sx=(root.winfo_screenwidth()-ow)//2; sy=(root.winfo_screenheight()-oh)//2
    ov.geometry(f"{ow}x{oh}+{sx}+{sy}")
    tk.Label(ov,text=f"📍 {step_no}" if step_no else "📍 좌표 캡처",font=("Arial",12,"bold"),fg="#7f8fa6",bg="#1a1a2e").pack(pady=(20,0))
    tk.Label(ov,text=step_title,font=("Arial",24,"bold"),fg="#fff",bg="#1a1a2e").pack(pady=(4,8))
    cd=tk.Label(ov,text=str(countdown_sec),font=("Arial",52,"bold"),fg="#feca57",bg="#1a1a2e"); cd.pack(pady=(0,5))
    tk.Label(ov,text="마우스를 목표 위치에 올려놓으세요",font=("Arial",11),fg="#a4b0be",bg="#1a1a2e").pack()
    ov.update()
    for s in range(countdown_sec,0,-1):
        cd.config(text=str(s)); 
        if s<=2: cd.config(fg="#ee5a24")
        ov.update(); time.sleep(1)
    pos=pyautogui.position()
    cd.config(text="✓",fg="#2ed573",font=("Arial",48,"bold")); ov.update(); time.sleep(0.5); ov.destroy()
    return pos

# =========================================================
#  인증
# =========================================================
def verify_password():
    cfg = load_admin_config()
    if not check_expiration(cfg): return False
    
    root_auth = tk.Tk()
    root_auth.withdraw()
    init_fonts()
    
    pw = simpledialog.askstring("인증", "프로그램 실행 비밀번호를 입력하세요:", show="*")
    if pw is None:
        root_auth.destroy()
        return False
        
    pw = pw.strip()
    remote_url = cfg.get("remote_password_url", "").strip()
    
    is_valid = False
    auth_method = "로컬"
    
    if remote_url and remote_url.startswith("http"):
        try:
            import urllib.request
            import json
            req = urllib.request.Request(
                remote_url,
                headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req, timeout=5) as response:
                content = response.read().decode("utf-8")
                allowed_pws = [line.strip() for line in content.splitlines() if line.strip()]
                if pw in allowed_pws:
                    is_valid = True
                    auth_method = "원격 웹"
        except Exception as e:
            write_usage_log("원격 인증 에러", str(e))
            
    if not is_valid:
        correct_pw = cfg.get("user_password", "0303").strip()
        if pw == correct_pw:
            is_valid = True
            
    if is_valid:
        root_auth.destroy()
        write_usage_log("프로그램 접속", f"인증성공 ({auth_method})")
        return True
    else:
        messagebox.showerror("인증 실패", "비밀번호가 일치하지 않습니다.")
        write_usage_log("인증 실패")
        root_auth.destroy()
        return False

# =========================================================
#  데이터 로드
# =========================================================
def process_and_set_data(df_raw, silent=False):
    global data, registered_indices, failed_indices, start_index
    try:
        df_raw.columns=[str(c).strip() for c in df_raw.columns]
        if "전화번호" not in df_raw.columns:
            messagebox.showerror("오류","'전화번호' 열이 없습니다."); return
        if "발송문자" not in df_raw.columns: df_raw["발송문자"]=""
        dc=df_raw.copy()
        dc['전화번호']=dc['전화번호'].astype(str).str.replace(".0","",regex=False).str.strip()
        dc=dc[dc['전화번호'].str.match(r'^\d+$',na=False)]
        dc['전화번호']=dc['전화번호'].str.zfill(11)
        dc['발송문자']=dc['발송문자'].astype(str).str.strip()
        data=dc[["전화번호","발송문자"]].reset_index(drop=True)
        registered_indices.clear(); failed_indices.clear(); start_index=0
        update_tables(); update_progress()
        if not silent:
            write_usage_log("데이터 로드",f"{len(data)}건")
            messagebox.showinfo("성공",f"총 {len(data)}개 로드!")
    except Exception as e: messagebox.showerror("오류",str(e))

def load_from_active_excel():
    global loaded_file_path
    try:
        app=xw.apps.active
        if not app: messagebox.showwarning("실패","엑셀이 열려있지 않습니다."); return
        sheet=app.books.active.sheets.active
        rb=sheet.range('B1').expand('down').value; rc=sheet.range('C1').expand('down').value
        ml=max(len(rb) if isinstance(rb,list) else 1, len(rc) if isinstance(rc,list) else 1)
        lb=rb if isinstance(rb,list) else [rb]; lc=rc if isinstance(rc,list) else [rc]
        lb+=[""]*(ml-len(lb)); lc+=[""]*(ml-len(lc))
        df=pd.DataFrame({"전화번호":lb,"발송문자":lc})
        process_and_set_data(df); loaded_file_path=""
        filename_label.config(text="📊 연동: 활성 엑셀 시트")
    except Exception as e: messagebox.showerror("오류",str(e))

def load_from_file_path():
    global loaded_file_path
    fp=filedialog.askopenfilename(title="엑셀 파일 선택",filetypes=[("Excel","*.xlsx *.xls")])
    if not fp: return
    try:
        df=pd.read_excel(fp,engine='openpyxl')
        if df.shape[1]>=3: df=df.iloc[:,[1,2]]; df.columns=["전화번호","발송문자"]
        elif df.shape[1]>=2: df=df.iloc[:,[1]]; df.columns=["전화번호"]
        else: df.columns=["전화번호"]
        process_and_set_data(df); loaded_file_path=fp
        filename_label.config(text=f"📂 파일: {os.path.basename(fp)}")
    except Exception as e: messagebox.showerror("오류",str(e))

# =========================================================
#  GUI 헬퍼
# =========================================================
def get_step_delay():
    try: return float(delay_var.get())
    except: return 3.0
def get_batch_size():
    try: return max(1,int(batch_var.get()))
    except: return 100
def get_retry_count():
    try: return max(0,int(retry_var.get()))
    except: return 1
def update_status(t): root.after(0,lambda:status_var.set(t))
def update_progress():
    total=len(data); done=len(registered_indices); fail=len(failed_indices)
    left=total-done-fail; pct=int((done+fail)/total*100) if total else 0
    root.after(0,lambda:(progress_var.set(pct),lbl_progress.config(text=f"  완료 {done} | 실패 {fail} | 남은 {left}  ({pct}%)")))
def update_tables():
    left_table.delete(*left_table.get_children())
    for i,r in data.iterrows():
        if i not in registered_indices and i not in failed_indices:
            left_table.insert("","end",values=(i+1,r["전화번호"],r["발송문자"][:30]))
    right_table.delete(*right_table.get_children())
    for i in sorted(registered_indices):
        right_table.insert("","end",values=(i+1,data.loc[i,"전화번호"],"✅ 완료"))
    for i in sorted(failed_indices):
        right_table.insert("","end",values=(i+1,data.loc[i,"전화번호"],"❌ 실패"))
def get_current_mode(): return work_mode.get()

# =========================================================
#  모드 A: 번호 등록
# =========================================================
def setup_coord_dynamic(mode):
    global macro_config
    cfg = load_macro_config()
    steps = cfg.get(mode, [])
    
    coord_steps = [s for s in steps if s.get("type") in ["click_type", "click_only"]]
    if not coord_steps:
        messagebox.showinfo("안내", f"[{mode}] 모드에는 마우스 클릭 좌표를 설정할 단계가 없습니다.")
        return True
        
    messagebox.showinfo("좌표 설정", f"[{mode}] 모드의 마우스 좌표 {len(coord_steps)}개를 순서대로 설정합니다.\n\n각 안내 팝업을 확인하고 5초 이내에 마우스를 위치시켜 주세요.")
    
    for i, s in enumerate(coord_steps):
        title = s.get("name", f"단계 {i+1}")
        desc = f"[{mode}] {title} 위에 마우스를 올려놓으세요."
        no = f"({i+1}/{len(coord_steps)})"
        coord = capture_coord_popup(title, desc, no, 5)
        if not coord:
            messagebox.showwarning("취소", "좌표 설정이 취소되었습니다.")
            return False
        s["coord"] = coord
        
    save_macro_config(cfg)
    macro_config = cfg
    
    try:
        btn_coord_reset.config(text="🔄 좌표 재설정 (설정됨)")
    except:
        pass
        
    summary = "\n".join([f"• {s['name']}: {s['coord']}" for s in coord_steps])
    messagebox.showinfo("🎯 완료", f"[{mode}] 좌표 설정이 완료되었습니다!\n\n{summary}\n\n이제 실행 버튼을 누르세요.")
    return True

def setup_coord_registration():
    return setup_coord_dynamic("번호등록")

def run_macro_for_row(mode, phone_val, msg_val, row_idx):
    cfg_macro = load_macro_config()
    steps = cfg_macro.get(mode, [])
    delay = get_step_delay()
    prefix = "즉시발송" if row_idx == -2 else f"{row_idx+1}"
    
    try:
        for i, s in enumerate(steps):
            stype = s.get("type")
            name = s.get("name", f"단계 {i+1}")
            coord = s.get("coord")
            
            update_status(f"[{prefix}] {name}")
            
            if stype == "click_type":
                if not coord: continue
                pyautogui.click(*coord); time.sleep(0.3)
                pyautogui.hotkey(MODIFIER, "a"); time.sleep(0.15)
                pyautogui.press("backspace"); time.sleep(0.3)
                
                vsrc = s.get("value_source", "static")
                if vsrc == "phone":
                    val = phone_val
                elif vsrc == "message":
                    val = msg_val
                else:
                    val = s.get("static_value", "")
                    
                safe_input(val); time.sleep(delay)
                
            elif stype == "click_only":
                if not coord: continue
                pyautogui.click(*coord); time.sleep(0.15)
                if s.get("press_enter", True):
                    pyautogui.press("enter")
                time.sleep(delay)
                
            elif stype == "key":
                key_val = s.get("key_value", "enter")
                pyautogui.press(key_val); time.sleep(delay)
                
            elif stype == "delay":
                d_val = float(s.get("delay_value", 1.0))
                time.sleep(d_val)
                
        return True
    except Exception as e:
        write_log("오류", f"매크로 실행 실패: {e}")
        return False

def registration_worker():
    global start_index, is_running
    batch=get_batch_size(); end_idx=start_index+batch
    if start_index>=len(data): messagebox.showinfo("완료","등록할 번호가 없습니다."); return
    is_running=True; set_buttons_running(True)
    cnt=0
    for idx in range(start_index,min(end_idx,len(data))):
        if not is_running: break
        while is_paused: time.sleep(0.3)
        num=data.iloc[idx]['전화번호']
        try:
            ok = run_macro_for_row("번호등록", num, "", idx)
            if ok:
                registered_indices.add(idx); write_log("등록성공",num)
                write_result_to_excel(idx, "등록완료")
                cnt+=1
            else:
                failed_indices.add(idx); write_log("등록실패",num)
                write_result_to_excel(idx, "등록실패")
        except:
            failed_indices.add(idx); write_log("등록실패",num)
            write_result_to_excel(idx, "등록실패")
        root.after(0,update_tables); root.after(0,update_progress)
    start_index=min(end_idx,len(data))
    save_session()  # ✅ 자동 저장
    is_running=False; set_buttons_running(False)
    write_usage_log("번호등록",f"{cnt}건")
    update_status(f"등록 완료: {cnt}개")
    messagebox.showinfo("완료",f"{cnt}개 등록! (누적 {len(registered_indices)}개)")

def setup_coord_sms():
    return setup_coord_dynamic("문자발송")

def sms_send_one(idx, phone, message):
    """1건 발송 시도. 성공=True, 실패=False"""
    return run_macro_for_row("문자발송", phone, message, idx)

def sms_worker():
    global is_running
    is_running=True; set_buttons_running(True)
    mode=send_mode.get(); global_msg=fixed_msg_text.get("1.0",tk.END).strip()
    retry_max=get_retry_count()
    retry_queue=[]  # ✅ 재시도 대기열

    # ── 1차 발송 ──
    for idx in range(len(data)):
        if not is_running: break
        while is_paused: update_status(f"⏸ 일시정지 ({idx+1}번)"); time.sleep(0.3)
        if idx in registered_indices or idx in failed_indices: continue
        row=data.iloc[idx]; phone=row["전화번호"]
        msg=global_msg if mode=="일괄" else row["발송문자"]
        ok=sms_send_one(idx,phone,msg)
        if ok:
            registered_indices.add(idx); write_log("발송성공",phone,msg)
            write_result_to_excel(idx,"발송성공")
        else:
            retry_queue.append((idx,phone,msg))  # 실패 → 재시도 대기열
        root.after(0,update_tables); root.after(0,update_progress)
        save_session()  # ✅ 건별 자동 저장

    # ── ✅ 실패 자동 재시도 ──
    for attempt in range(1, retry_max+1):
        if not retry_queue or not is_running: break
        update_status(f"🔄 재시도 {attempt}/{retry_max} ({len(retry_queue)}건)")
        time.sleep(2)
        still_failed=[]
        for idx,phone,msg in retry_queue:
            if not is_running: break
            update_status(f"🔄 재시도 [{idx+1}] {phone} (시도 {attempt})")
            ok=sms_send_one(idx,phone,msg)
            if ok:
                registered_indices.add(idx); write_log(f"재시도{attempt}성공",phone,msg)
                write_result_to_excel(idx,f"재시도{attempt}성공")
            else:
                still_failed.append((idx,phone,msg))
            root.after(0,update_tables); root.after(0,update_progress)
        retry_queue=still_failed

    # 최종 실패 처리
    for idx,phone,msg in retry_queue:
        failed_indices.add(idx); write_log("최종실패",phone,msg)
        write_result_to_excel(idx,"발송실패")
    root.after(0,update_tables); root.after(0,update_progress)

    save_session()
    is_running=False; set_buttons_running(False)
    retry_cnt=retry_max if retry_queue else 0
    write_usage_log("문자발송",f"성공{len(registered_indices)} 실패{len(failed_indices)}")
    update_status("전체 완료" if len(registered_indices)+len(failed_indices)>=len(data) else "중지됨")
    messagebox.showinfo("완료",f"성공 {len(registered_indices)}건 / 실패 {len(failed_indices)}건"
                        + (f"\n(재시도 {retry_max}회 적용됨)" if retry_max else ""))

# =========================================================
#  통합 실행
# =========================================================
def start_worker():
    mode=get_current_mode()
    if data.empty: messagebox.showwarning("데이터 없음","데이터를 불러오세요."); return
    cfg_macro = load_macro_config()
    coord_steps = [s for s in cfg_macro.get(mode, []) if s.get("type") in ["click_type", "click_only"]]
    if any(s.get("coord") is None for s in coord_steps):
        if not setup_coord_dynamic(mode):
            return
    if mode=="번호등록":
        threading.Thread(target=registration_worker,daemon=True).start()
    else:
        threading.Thread(target=sms_worker,daemon=True).start()

def set_buttons_running(r):
    def update_ui():
        set_btn_state(btn_start, tk.DISABLED if r else tk.NORMAL, COLOR_SUCCESS, COLOR_SUCCESS_HOVER)
        set_btn_state(btn_pause, tk.NORMAL if r else tk.DISABLED, "#f59e0b", "#d97706")
        set_btn_state(btn_resume, tk.NORMAL if r else tk.DISABLED, "#06b6d4", "#0891b2")
        set_btn_state(btn_stop, tk.NORMAL if r else tk.DISABLED, COLOR_DANGER, COLOR_DANGER_HOVER)
    root.after(0, update_ui)

def reset_coordinates():
    global macro_config
    cfg = load_macro_config()
    for mode in ["번호등록", "문자발송"]:
        for s in cfg[mode]:
            if "coord" in s:
                s["coord"] = None
    save_macro_config(cfg)
    macro_config = cfg
    try:
        btn_coord_reset.config(text="🔄 좌표 재설정 (미설정)")
    except:
        pass
    messagebox.showinfo("초기화 완료", "모든 매크로 좌표가 초기화되었습니다.")

def reset_all():
    global is_running,start_index
    is_running=False; start_index=0
    registered_indices.clear(); failed_indices.clear()
    clear_session(); update_tables(); update_progress(); update_status("초기화 완료")

def toggle_send_mode():
    if send_mode.get()=="일괄": fixed_msg_text.config(state=tk.NORMAL,bg="#ffffff")
    else: fixed_msg_text.config(state=tk.DISABLED,bg="#e9ecef")

def toggle_work_mode():
    m=get_current_mode()
    if m=="번호등록":
        sms_opts_frame.pack_forget(); reg_opts_frame.pack(fill=tk.X,pady=(0,5),after=mode_frame)
        btn_start.config(text="▶ 번호등록 실행")
    else:
        reg_opts_frame.pack_forget(); sms_opts_frame.pack(fill=tk.X,pady=(0,5),after=mode_frame)
        btn_start.config(text="▶ 문자발송 실행")

# =========================================================
#  매크로 단계 커스텀 빌더
# =========================================================
def open_step_builder_panel():
    win = tk.Toplevel(root)
    win.title("🛠️ 매크로 단계 커스텀 빌더")
    win.geometry("750x640")
    win.configure(bg=COLOR_NAVY)
    win.attributes("-topmost", True)
    win.resizable(False, False)
    
    tk.Label(win, text="🛠️ 매크로 단계 커스텀 설정", font=FONT_TITLE, bg=COLOR_NAVY, fg="white", pady=15).pack(fill=tk.X)
    
    bd = tk.Frame(win, bg="#f8fafc", padx=15, pady=15)
    bd.pack(fill=tk.BOTH, expand=True)
    
    mode_frame = tk.Frame(bd, bg="#f8fafc")
    mode_frame.pack(fill=tk.X, pady=(0, 10))
    tk.Label(mode_frame, text="⚙️ 작업 모드 선택:", font=FONT_BOLD_10, bg="#f8fafc", fg="#1e293b").pack(side=tk.LEFT, padx=5)
    
    selected_mode = tk.StringVar(value="문자발송")
    
    tree_frame = tk.Frame(bd, bg="#f8fafc")
    tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    scrollbar = tk.Scrollbar(tree_frame, orient="vertical")
    tree = ttk.Treeview(tree_frame, columns=("no", "name", "type", "coord", "val"), show="headings", yscrollcommand=scrollbar.set, style="Treeview")
    scrollbar.config(command=tree.yview)
    
    tree.heading("no", text="No")
    tree.heading("name", text="단계 이름")
    tree.heading("type", text="동작 유형")
    tree.heading("coord", text="마우스 좌표")
    tree.heading("val", text="값/설정")
    
    tree.column("no", width=40, anchor="center")
    tree.column("name", width=180)
    tree.column("type", width=100, anchor="center")
    tree.column("coord", width=120, anchor="center")
    tree.column("val", width=180)
    
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    cfg_working = load_macro_config()
    
    def refresh_tree():
        for item in tree.get_children():
            tree.delete(item)
        
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        for idx, s in enumerate(steps):
            stype = s.get("type")
            stype_kr = {
                "click_type": "클릭 후 입력",
                "click_only": "단독 클릭",
                "key": "키 입력",
                "delay": "대기 시간"
            }.get(stype, stype)
            
            coord_str = f"({s['coord'][0]}, {s['coord'][1]})" if s.get("coord") else "미설정"
            
            if stype == "click_type":
                vsrc = s.get("value_source", "static")
                vsrc_kr = {"phone": "엑셀 전화번호", "message": "엑셀 발송문구", "static": "고정 텍스트"}.get(vsrc, vsrc)
                if vsrc == "static":
                    val_str = f"고정: '{s.get('static_value', '')}'"
                else:
                    val_str = vsrc_kr
            elif stype == "click_only":
                val_str = "Enter 입력" if s.get("press_enter", True) else "클릭만"
            elif stype == "key":
                val_str = f"키: '{s.get('key_value', 'enter')}'"
            elif stype == "delay":
                val_str = f"{s.get('delay_value', 1.0)}초 대기"
            else:
                val_str = ""
                
            tree.insert("", tk.END, values=(idx + 1, s.get("name", ""), stype_kr, coord_str, val_str))
            
    btn_side_frame = tk.Frame(bd, bg="#f8fafc")
    btn_side_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0), pady=5)
    
    prop_frame = tk.LabelFrame(bd, text=" 📝 선택한 단계 속성 편집 ", font=FONT_BOLD_10, bg="#f8fafc", fg=COLOR_NAVY, bd=1, relief="solid", padx=10, pady=8)
    prop_frame.pack(fill=tk.X, pady=10)
    
    tk.Label(prop_frame, text="단계 이름:", font=FONT_REG_10, bg="#f8fafc", fg="#475569").grid(row=0, column=0, sticky="w", pady=2)
    name_var = tk.StringVar()
    name_entry = tk.Entry(prop_frame, textvariable=name_var, font=FONT_REG_10, width=20, bd=1, relief="solid")
    name_entry.grid(row=0, column=1, sticky="w", pady=2, padx=5)
    
    tk.Label(prop_frame, text="동작 유형:", font=FONT_REG_10, bg="#f8fafc", fg="#475569").grid(row=0, column=2, sticky="w", pady=2, padx=(15, 0))
    type_var = tk.StringVar(value="click_type")
    type_combo = ttk.Combobox(prop_frame, textvariable=type_var, values=["클릭 후 입력", "단독 클릭", "키 입력", "대기 시간"], state="readonly", width=12, style="TCombobox")
    type_combo.grid(row=0, column=3, sticky="w", pady=2, padx=5)
    
    opts_sub_frame = tk.Frame(prop_frame, bg="#f8fafc")
    opts_sub_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=5)
    
    lbl_vsrc = tk.Label(opts_sub_frame, text="입력값 종류:", font=FONT_REG_10, bg="#f8fafc", fg="#475569")
    vsrc_var = tk.StringVar(value="message")
    combo_vsrc = ttk.Combobox(opts_sub_frame, textvariable=vsrc_var, values=["엑셀 발송문구", "엑셀 전화번호", "직접 지정"], state="readonly", width=15, style="TCombobox")
    
    lbl_static = tk.Label(opts_sub_frame, text="고정 텍스트:", font=FONT_REG_10, bg="#f8fafc", fg="#475569")
    static_var = tk.StringVar()
    entry_static = tk.Entry(opts_sub_frame, textvariable=static_var, font=FONT_REG_10, width=25, bd=1, relief="solid")
    
    enter_var = tk.BooleanVar(value=True)
    chk_enter = tk.Checkbutton(opts_sub_frame, text="클릭 후 Enter 키 입력 (팝업 확인용)", variable=enter_var, font=FONT_REG_10, bg="#f8fafc", activebackground="#f8fafc", selectcolor="#f8fafc")
    
    lbl_key = tk.Label(opts_sub_frame, text="키 종류:", font=FONT_REG_10, bg="#f8fafc", fg="#475569")
    key_var = tk.StringVar(value="enter")
    combo_key = ttk.Combobox(opts_sub_frame, textvariable=key_var, values=["enter", "tab", "backspace", "escape", "space", "up", "down"], width=10, style="TCombobox")
    
    lbl_delay = tk.Label(opts_sub_frame, text="대기 시간 (초):", font=FONT_REG_10, bg="#f8fafc", fg="#475569")
    delay_val_var = tk.StringVar(value="1.0")
    spin_delay = tk.Spinbox(opts_sub_frame, from_=0.1, to=10.0, increment=0.1, textvariable=delay_val_var, width=6, font=FONT_BOLD_10, justify="center", bd=1, relief="solid")

    def hide_all_options():
        for widget in opts_sub_frame.winfo_children():
            widget.pack_forget()
            widget.grid_forget()
            
    def show_options_for_type(*args):
        hide_all_options()
        t = type_var.get()
        if t == "클릭 후 입력":
            lbl_vsrc.grid(row=0, column=0, sticky="w", pady=2)
            combo_vsrc.grid(row=0, column=1, sticky="w", pady=2, padx=5)
            show_static_field()
        elif t == "단독 클릭":
            chk_enter.grid(row=0, column=0, sticky="w", pady=2)
        elif t == "키 입력":
            lbl_key.grid(row=0, column=0, sticky="w", pady=2)
            combo_key.grid(row=0, column=1, sticky="w", pady=2, padx=5)
        elif t == "대기 시간":
            lbl_delay.grid(row=0, column=0, sticky="w", pady=2)
            spin_delay.grid(row=0, column=1, sticky="w", pady=2, padx=5)
            
    def show_static_field(*args):
        v = vsrc_var.get()
        if v == "직접 지정":
            lbl_static.grid(row=0, column=2, sticky="w", pady=2, padx=(15, 0))
            entry_static.grid(row=0, column=3, sticky="w", pady=2, padx=5)
        else:
            lbl_static.grid_forget()
            entry_static.grid_forget()
            
    type_combo.bind("<<ComboboxSelected>>", show_options_for_type)
    combo_vsrc.bind("<<ComboboxSelected>>", show_static_field)
    
    def on_tree_select(event):
        sel = tree.selection()
        if not sel: return
        idx = tree.index(sel[0])
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        if idx >= len(steps): return
        s = steps[idx]
        
        name_var.set(s.get("name", ""))
        stype = s.get("type", "click_type")
        stype_kr = {"click_type": "클릭 후 입력", "click_only": "단독 클릭", "key": "키 입력", "delay": "대기 시간"}.get(stype, "클릭 후 입력")
        type_var.set(stype_kr)
        
        show_options_for_type()
        
        if stype == "click_type":
            vsrc = s.get("value_source", "message")
            vsrc_kr = {"message": "엑셀 발송문구", "phone": "엑셀 전화번호", "static": "직접 지정"}.get(vsrc, "엑셀 발송문구")
            vsrc_var.set(vsrc_kr)
            static_var.set(s.get("static_value", ""))
            show_static_field()
        elif stype == "click_only":
            enter_var.set(s.get("press_enter", True))
        elif stype == "key":
            key_var.set(s.get("key_value", "enter"))
        elif stype == "delay":
            delay_val_var.set(str(s.get("delay_value", 1.0)))
            
    tree.bind("<<TreeviewSelect>>", on_tree_select)
    
    def apply_properties():
        sel = tree.selection()
        if not sel: return
        idx = tree.index(sel[0])
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        if idx >= len(steps): return
        s = steps[idx]
        
        s["name"] = name_var.get().strip() or f"단계 {idx+1}"
        tk_type = type_var.get()
        stype = {"클릭 후 입력": "click_type", "단독 클릭": "click_only", "키 입력": "key", "대기 시간": "delay"}.get(tk_type, "click_type")
        s["type"] = stype
        
        if stype == "click_type":
            tk_vsrc = vsrc_var.get()
            vsrc = {"엑셀 발송문구": "message", "엑셀 전화번호": "phone", "직접 지정": "static"}.get(tk_vsrc, "message")
            s["value_source"] = vsrc
            s["static_value"] = static_var.get().strip()
            if "coord" not in s: s["coord"] = None
        elif stype == "click_only":
            s["press_enter"] = enter_var.get()
            if "coord" not in s: s["coord"] = None
        elif stype == "key":
            s["key_value"] = combo_key.get().strip() or "enter"
            if "coord" in s: del s["coord"]
        elif stype == "delay":
            try: s["delay_value"] = float(delay_val_var.get())
            except: s["delay_value"] = 1.0
            if "coord" in s: del s["coord"]
            
        refresh_tree()
        items = tree.get_children()
        if idx < len(items):
            tree.selection_set(items[idx])
            
    create_btn(prop_frame, "✔️ 설정 변경 적용", apply_properties, COLOR_SUCCESS, COLOR_SUCCESS_HOVER, font=FONT_BOLD_9).grid(row=0, column=4, rowspan=2, sticky="ns", padx=(10, 0))
    
    def move_step(direction):
        sel = tree.selection()
        if not sel: return
        idx = tree.index(sel[0])
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(steps): return
        
        steps[idx], steps[new_idx] = steps[new_idx], steps[idx]
        refresh_tree()
        tree.selection_set(tree.get_children()[new_idx])
        
    def add_step():
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        new_step = {"type": "click_type", "name": f"새 단계 {len(steps)+1}", "coord": None, "value_source": "message"}
        steps.append(new_step)
        refresh_tree()
        items = tree.get_children()
        if items:
            tree.selection_set(items[-1])
            
    def delete_step():
        sel = tree.selection()
        if not sel: return
        idx = tree.index(sel[0])
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        if idx >= len(steps): return
        
        del steps[idx]
        refresh_tree()
        items = tree.get_children()
        if items:
            tree.selection_set(items[min(idx, len(items)-1)])
            
    def capture_coord_selected():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("선택 필요", "좌표를 설정할 단계를 선택하세요.", parent=win)
            return
        idx = tree.index(sel[0])
        mode = selected_mode.get()
        steps = cfg_working.get(mode, [])
        s = steps[idx]
        
        if s.get("type") not in ["click_type", "click_only"]:
            messagebox.showwarning("좌표 불필요", "마우스 클릭이 필요한 단계만 좌표 설정이 가능합니다.", parent=win)
            return
            
        title = s.get("name", "좌표 설정")
        desc = f"[{mode}] {title} 위에 마우스를 올려놓으세요."
        
        win.attributes("-topmost", False)
        coord = capture_coord_popup(title, desc, "(1/1)", 5)
        win.attributes("-topmost", True)
        
        if coord:
            s["coord"] = coord
            refresh_tree()
            tree.selection_set(tree.get_children()[idx])
            messagebox.showinfo("완료", f"'{title}' 좌표가 {coord}로 설정되었습니다.", parent=win)

    create_btn(btn_side_frame, "🔼 위로", lambda: move_step(-1), "#475569", "#334155", width=12).pack(pady=2)
    create_btn(btn_side_frame, "🔽 아래로", lambda: move_step(1), "#475569", "#334155", width=12).pack(pady=2)
    create_btn(btn_side_frame, "📍 좌표 설정", capture_coord_selected, COLOR_PRIMARY, COLOR_PRIMARY_HOVER, width=12).pack(pady=10)
    create_btn(btn_side_frame, "➕ 단계 추가", add_step, COLOR_INFO, COLOR_INFO_HOVER, width=12).pack(pady=2)
    create_btn(btn_side_frame, "🗑️ 단계 삭제", delete_step, COLOR_DANGER, COLOR_DANGER_HOVER, width=12).pack(pady=2)
    
    def save_and_close():
        save_macro_config(cfg_working)
        global macro_config
        macro_config = cfg_working
        try:
            reg_set = all(s.get("coord") is not None for s in cfg_working["번호등록"] if s.get("type") in ["click_type", "click_only"])
            sms_set = all(s.get("coord") is not None for s in cfg_working["문자발송"] if s.get("type") in ["click_type", "click_only"])
            if reg_set and sms_set:
                btn_coord_reset.config(text="🔄 좌표 재설정 (설정됨)")
            else:
                btn_coord_reset.config(text="🔄 좌표 재설정 (미설정)")
        except:
            pass
        messagebox.showinfo("저장 완료", "매크로 단계 설정이 파일에 저장되었습니다.", parent=win)
        win.destroy()
        
    bottom_frame = tk.Frame(bd, bg="#f8fafc")
    bottom_frame.pack(fill=tk.X, pady=(10, 0))
    create_btn(bottom_frame, "💾 설정 파일 저장", save_and_close, COLOR_SUCCESS, COLOR_SUCCESS_HOVER, width=15).pack(side=tk.LEFT)
    create_btn(bottom_frame, "닫기 (취소)", win.destroy, "#94a3b8", "#64748b", width=10).pack(side=tk.RIGHT)
    
    tk.Radiobutton(mode_frame, text="📋 번호등록 단계", variable=selected_mode, value="번호등록", command=refresh_tree, font=FONT_BOLD_10, bg="#f8fafc", fg="#1e293b", activebackground="#f8fafc", selectcolor="#f8fafc").pack(side=tk.LEFT, padx=15)
    tk.Radiobutton(mode_frame, text="💬 문자발송 단계", variable=selected_mode, value="문자발송", command=refresh_tree, font=FONT_BOLD_10, bg="#f8fafc", fg="#1e293b", activebackground="#f8fafc", selectcolor="#f8fafc").pack(side=tk.LEFT)
    
    refresh_tree()
    show_options_for_type()
    if tree.get_children():
        tree.selection_set(tree.get_children()[0])

# =========================================================
#  관리자 패널
# =========================================================
def open_admin_panel():
    pw=simpledialog.askstring("🔐 관리자","관리자 비밀번호:",show="*")
    if pw is None: return
    if hashlib.sha256(pw.encode()).hexdigest()!=ADMIN_PW_HASH:
        messagebox.showerror("실패","관리자 비밀번호 불일치"); return
    write_usage_log("관리자 패널")
    cfg=load_admin_config()
    
    # Admin Toplevel Window (크기를 560x700으로 세로 확장)
    win=tk.Toplevel(root); win.title("🔧 관리자 설정"); win.geometry("560x700"); win.configure(bg=COLOR_NAVY)
    win.attributes("-topmost",True); win.resizable(False,False)
    
    tk.Label(win,text="🔧 관리자 설정 패널",font=FONT_TITLE,bg=COLOR_NAVY,fg="white",pady=15).pack(fill=tk.X)
    
    bd=tk.Frame(win,bg="#f8fafc",padx=25,pady=20); bd.pack(fill=tk.BOTH,expand=True)
    
    tk.Label(bd,text="🔑 프로그램 접속 비밀번호",font=FONT_BOLD_10,bg="#f8fafc",fg="#1e293b").pack(anchor="w", pady=(0, 4))
    pv=tk.StringVar(value=cfg.get("user_password","0303"))
    tk.Entry(bd,textvariable=pv,font=FONT_REG_10,width=20,bd=1,relief="solid").pack(anchor="w",pady=(0,15))
    
    tk.Label(bd,text="🌐 원격 비밀번호 확인용 URL (선택사항)",font=FONT_BOLD_10,bg="#f8fafc",fg="#1e293b").pack(anchor="w", pady=(0, 4))
    rv=tk.StringVar(value=cfg.get("remote_password_url",""))
    tk.Entry(bd,textvariable=rv,font=FONT_REG_10,width=50,bd=1,relief="solid").pack(anchor="w",pady=(0,5))
    tk.Label(bd,text="* 메모장(.txt) 형태의 웹 링크를 입력하면 실시간으로 해당 파일 안의 비밀번호 목록을 검증합니다.",font=FONT_REG_9,fg=COLOR_MUTED,bg="#f8fafc").pack(anchor="w",pady=(0,15))
    
    tk.Label(bd,text="📅 만료일 (YYYY-MM-DD, 비워두면 평생 무제한)",font=FONT_BOLD_10,bg="#f8fafc",fg="#1e293b").pack(anchor="w", pady=(0, 4))
    ev=tk.StringVar(value=cfg.get("expire_date",""))
    tk.Entry(bd,textvariable=ev,font=FONT_REG_10,width=20,bd=1,relief="solid").pack(anchor="w",pady=(0,15))
    
    tk.Label(bd,text="💬 사용 기간 만료 안내 메시지",font=FONT_BOLD_10,bg="#f8fafc",fg="#1e293b").pack(anchor="w", pady=(0, 4))
    mt=tk.Text(bd,height=3,font=FONT_REG_10,bd=1,relief="solid"); mt.pack(fill=tk.X,pady=(0,15))
    mt.insert(tk.END,cfg.get("expire_message",""))
    
    # 디스코드 웹훅 입력 필드 추가
    tk.Label(bd,text="📢 디스코드 실시간 알림 웹훅 URL",font=FONT_BOLD_10,bg="#f8fafc",fg="#1e293b").pack(anchor="w", pady=(0, 4))
    dw_frame = tk.Frame(bd, bg="#f8fafc")
    dw_frame.pack(fill=tk.X, pady=(0,15))
    dv=tk.StringVar(value=cfg.get("discord_webhook",""))
    tk.Entry(dw_frame,textvariable=dv,font=FONT_REG_10,width=38,bd=1,relief="solid").pack(side=tk.LEFT, ipady=2)
    
    def test_discord_webhook():
        webhook_url = dv.get().strip()
        if not webhook_url:
            messagebox.showwarning("입력 필요", "디스코드 웹훅 URL을 먼저 입력해 주세요.", parent=win)
            return
            
        test_payload = {
            "embeds": [{
                "title": "🧪 CIS Discord Webhook Test Connection",
                "description": "디스코드 알림 연동 테스트가 성공적으로 완료되었습니다!",
                "color": 0x2ecc71,  # Green
                "fields": [
                    {"name": "테스트 시간", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "inline": True},
                    {"name": "접속 PC", "value": pc, "inline": True}
                ],
                "footer": {"text": "Real-time Monitoring System Test"}
            }]
        }
        
        def test_worker():
            try:
                import urllib.request
                import json
                req = urllib.request.Request(
                    webhook_url,
                    data=json.dumps(test_payload).encode("utf-8"),
                    headers={"Content-Type": "application/json", "User-Agent": "urllib-discord-bot"}
                )
                with urllib.request.urlopen(req, timeout=5) as response:
                    pass
                win.after(0, lambda: messagebox.showinfo("성공", "디스코드 테스트 메세지가 전송되었습니다.\n디스코드 채널을 확인하세요.", parent=win))
            except Exception as e:
                win.after(0, lambda: messagebox.showerror("실패", f"디스코드 웹훅 전송 실패:\n{e}", parent=win))
                
        threading.Thread(target=test_worker, daemon=True).start()

    create_btn(dw_frame, "🧪 테스트 전송", test_discord_webhook, COLOR_INFO, COLOR_INFO_HOVER, font=FONT_BOLD_9).pack(side=tk.LEFT, padx=(5,0))
    
    tk.Label(bd,text="📊 프로그램 사용 이력 모니터링",font=FONT_BOLD_10,bg="#f8fafc",fg="#1e293b").pack(anchor="w",pady=(5,5))
    bf=tk.Frame(bd,bg="#f8fafc"); bf.pack(fill=tk.X, pady=(0,15))
    
    def show_recent():
        if not os.path.exists(USAGE_LOG_FILE): messagebox.showinfo("이력","이력 없음"); return
        with open(USAGE_LOG_FILE,"r",encoding="utf-8") as f: lines=f.readlines()[-30:]
        lines.reverse()
        s=tk.Toplevel(win); s.title("최근 30건 이력"); s.geometry("650x400"); s.attributes("-topmost",True)
        t=tk.Text(s,font=FONT_REG_9); t.pack(fill=tk.BOTH,expand=True,padx=10,pady=10)
        t.insert(tk.END,"".join(lines)); t.config(state=tk.DISABLED)
        
    create_btn(bf, "📊 최근 사용 이력 확인", show_recent, COLOR_PRIMARY, COLOR_PRIMARY_HOVER, font=FONT_BOLD_9).pack(side=tk.LEFT,padx=3)
    create_btn(bf, "📄 전체 사용이력.txt 열기", lambda: open_path(USAGE_LOG_FILE) if os.path.exists(USAGE_LOG_FILE) else None, "#64748b", "#475569", font=FONT_BOLD_9).pack(side=tk.LEFT,padx=3)
    
    pc,user=get_user_info()
    tk.Label(bd,text=f"현재 접속 PC: {pc}  |  사용자 계정: {user}",font=FONT_REG_9,fg=COLOR_MUTED,bg="#f8fafc").pack(anchor="w",pady=(5,0))
    
    def save_cfg():
        ne=ev.get().strip()
        if ne:
            try: datetime.strptime(ne,"%Y-%m-%d")
            except: messagebox.showerror("형식 오류","만료일은 YYYY-MM-DD 형식으로 입력하세요."); return
        nc={"user_password":pv.get().strip() or "0303","expire_date":ne,"expire_message":mt.get("1.0",tk.END).strip(),
            "discord_webhook":dv.get().strip(),
            "remote_password_url":rv.get().strip(),
            "created":cfg.get("created",""),"last_modified":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"modified_by":f"{pc}/{user}"}
        save_admin_config(nc); write_usage_log("설정변경",f"만료={ne or '없음'}")
        messagebox.showinfo("저장","관리자 설정이 저장되었습니다."); win.destroy()
        
    sf=tk.Frame(bd,bg="#f8fafc"); sf.pack(fill=tk.X,pady=(15,0))
    create_btn(sf, "💾 설정 저장", save_cfg, COLOR_SUCCESS, COLOR_SUCCESS_HOVER, width=12).pack(side=tk.LEFT)
    create_btn(sf, "닫기", win.destroy, "#94a3b8", "#64748b", width=10).pack(side=tk.RIGHT)

# =========================================================
#  추가 기능: 예시 로드, 수동 등록, 즉시 발송
# =========================================================
def load_example_data():
    global data, registered_indices, failed_indices, start_index
    try:
        example_df = pd.DataFrame([
            {"전화번호": "01012345678", "발송문자": "[CIS 알림] 1차 테스트 발송 예시입니다."},
            {"전화번호": "01098765432", "발송문자": "[CIS 알림] 2차 개별 맞춤형 안내 문자입니다."},
            {"전화번호": "01011112222", "발송문자": "[CIS 알림] 3차 수신 확인용 예시입니다."}
        ])
        process_and_set_data(example_df, silent=True)
        filename_label.config(text="💡 데이터 상태: 예시용 DB 레코드 로드됨")
        write_usage_log("예시 데이터 로드")
        update_status("예시 데이터 로드 완료")
    except Exception as e:
        messagebox.showerror("오류", f"예시 데이터를 로드할 수 없습니다: {e}")

def add_manual_data():
    global data
    try:
        num = manual_phone_entry.get().strip()
        msg = manual_msg_entry.get().strip()
        if not num:
            messagebox.showwarning("입력 오류", "전화번호를 입력하세요.")
            return
        num = num.replace("-", "").replace(" ", "")
        if not num.isdigit():
            messagebox.showwarning("입력 오류", "전화번호는 숫자만 입력 가능합니다.")
            return
        num = num.zfill(11)
        
        new_row = pd.DataFrame([{"전화번호": num, "발송문자": msg}])
        data = pd.concat([data, new_row], ignore_index=True)
        
        manual_phone_entry.delete(0, tk.END)
        manual_msg_entry.delete(0, tk.END)
        
        update_tables()
        update_progress()
        update_status(f"수동 등록 추가: {num}")
    except Exception as e:
        messagebox.showerror("오류", f"수동 추가 중 오류 발생: {e}")

def quick_send_action():
    try:
        phone = quick_phone_entry.get().strip()
        msg = quick_msg_entry.get().strip()
        
        if not phone:
            messagebox.showwarning("입력 오류", "발송할 연락처를 입력하세요.")
            return
        if not msg:
            messagebox.showwarning("입력 오류", "발송할 문구를 입력하세요.")
            return
            
        phone = phone.replace("-", "").replace(" ", "")
        if not phone.isdigit():
            messagebox.showwarning("입력 오류", "연락처는 숫자만 입력 가능합니다.")
            return
        phone = phone.zfill(11)
        
        cfg_macro = load_macro_config()
        coord_steps = [s for s in cfg_macro.get("문자발송", []) if s.get("type") in ["click_type", "click_only"]]
        if any(s.get("coord") is None for s in coord_steps):
            messagebox.showwarning("좌표 미설정", "문자발송을 위한 좌표가 설정되지 않았습니다.\n먼저 [문자발송]에서 좌표 설정을 진행해 주세요.")
            setup_coord_sms()
            return
            
        def worker():
            write_usage_log("즉시 발송", f"연락처:{phone}")
            ok = sms_send_one(-2, phone, msg)
            if ok:
                messagebox.showinfo("성공", f"즉시 발송 완료!\n연락처: {phone}")
                write_log("즉시발송성공", phone, msg)
            else:
                messagebox.showerror("실패", f"즉시 발송 실패!\n연락처: {phone}")
                write_log("즉시발송실패", phone, msg)
                
        threading.Thread(target=worker, daemon=True).start()
    except Exception as e:
        messagebox.showerror("오류", f"즉시 발송 오류: {e}")

# =========================================================
#  GUI 구성
# =========================================================
def create_gui():
    global left_table,right_table,btn_start,btn_pause,btn_resume,btn_stop
    global btn_coord_reset,filename_label,send_mode,fixed_msg_text,root
    global delay_var,batch_var,retry_var,progress_var,lbl_progress,status_var
    global work_mode,mode_frame,sms_opts_frame,reg_opts_frame
    global tpl_var,tpl_dropdown,excel_writeback_var
    global manual_phone_entry, manual_msg_entry, quick_phone_entry, quick_msg_entry

    root=tk.Tk(); root.title("📱 CIS 통합 매크로 — 번호등록 + 문자발송")
    root.geometry("1200x920" if IS_MAC else "1200x950")
    root.configure(bg=BG_MAIN)

    if IS_MAC:
        root.bind_class("Entry", "<Command-c>", lambda e: e.widget.event_generate("<<Copy>>"))
        root.bind_class("Entry", "<Command-v>", lambda e: e.widget.event_generate("<<Paste>>"))
        root.bind_class("Entry", "<Command-x>", lambda e: e.widget.event_generate("<<Cut>>"))
        root.bind_class("Entry", "<Command-a>", lambda e: e.widget.select_range(0, tk.END) or "break")
        root.bind_class("Text", "<Command-c>", lambda e: e.widget.event_generate("<<Copy>>"))
        root.bind_class("Text", "<Command-v>", lambda e: e.widget.event_generate("<<Paste>>"))
        root.bind_class("Text", "<Command-x>", lambda e: e.widget.event_generate("<<Cut>>"))
        root.bind_class("Text", "<Command-a>", lambda e: e.widget.tag_add("sel", "1.0", "end") or "break")

    # ttk 스타일 재지정 (clam 테마 활성화 및 각 컴포넌트 세팅)
    style = ttk.Style()
    style.theme_use("clam")
    
    style.configure("Treeview",
                    background=BG_CARD,
                    foreground="#1e293b",
                    rowheight=26,
                    fieldbackground=BG_CARD,
                    bordercolor=COLOR_BORDER,
                    borderwidth=1,
                    font=FONT_REG_10)
    style.map("Treeview", 
              background=[("selected", COLOR_PRIMARY)], 
              foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading",
                    background="#e2e8f0",
                    foreground="#1e293b",
                    font=FONT_BOLD_9,
                    borderwidth=1)
    
    style.configure("TProgressbar",
                    thickness=15,
                    bordercolor=COLOR_BORDER,
                    troughcolor="#e2e8f0",
                    background=COLOR_PRIMARY)

    style.configure("TCombobox",
                    font=FONT_REG_10,
                    arrowsize=12,
                    fieldbackground=BG_CARD,
                    background=BG_MAIN)

    # 타이틀바 (Dark Navy)
    tf=tk.Frame(root,bg=COLOR_NAVY)
    tf.pack(fill=tk.X)
    tk.Label(tf,text="📱 CIS 단체/개별 선택 문자 발송 매크로",font=FONT_TITLE,bg=COLOR_NAVY,fg="white",pady=12).pack(side=tk.LEFT,padx=15)
    
    cfg=load_admin_config(); es=cfg.get("expire_date","").strip()
    if es:
        try:
            rm=(datetime.strptime(es,"%Y-%m-%d").date()-date.today()).days
            tk.Label(tf,text=f"📅 라이선스: {es} ({rm}일 남음)",font=FONT_BOLD_9,bg=COLOR_NAVY,fg="#10b981" if rm>7 else "#ef4444").pack(side=tk.LEFT,padx=20)
        except: pass
        
    create_btn(tf, "🔧 관리자 설정", open_admin_panel, "#334155", "#475569", fg_color="#f8fafc", font=FONT_BOLD_9).pack(side=tk.RIGHT,padx=15,pady=8)
    create_btn(tf, "🛠️ 매크로 단계 설정", open_step_builder_panel, "#0f766e", "#115e59", fg_color="#f8fafc", font=FONT_BOLD_9).pack(side=tk.RIGHT,padx=5,pady=8)

    # 1. 데이터 소스 및 등록 (White Card Style)
    lf=tk.LabelFrame(root,text=" 1. 데이터 소스 및 연락처 등록 ",font=FONT_BOLD_11,bg=BG_CARD,fg=COLOR_NAVY,bd=1,relief="solid",padx=15,pady=10)
    lf.pack(fill=tk.X,pady=(10,2),padx=20)
    
    # Row 1: 연동 & 예시
    f_row = tk.Frame(lf, bg=BG_CARD)
    f_row.pack(fill=tk.X, pady=(0, 8))
    create_btn(f_row, "📊 활성 엑셀 연동 (B열:전화번호 / C열:발송문구)", load_from_active_excel, COLOR_SUCCESS, COLOR_SUCCESS_HOVER, font=FONT_BOLD_10).pack(side=tk.LEFT,expand=True,fill=tk.X,padx=3)
    create_btn(f_row, "📂 엑셀 파일 선택", load_from_file_path, COLOR_INFO, COLOR_INFO_HOVER, font=FONT_BOLD_10).pack(side=tk.LEFT,expand=True,fill=tk.X,padx=3)
    create_btn(f_row, "💡 예시 데이터 로드", load_example_data, "#8b5cf6", "#7c3aed", font=FONT_BOLD_10).pack(side=tk.LEFT,expand=True,fill=tk.X,padx=3)
    
    # Row 2: 수동 리스트 추가
    m_row = tk.Frame(lf, bg=BG_CARD)
    m_row.pack(fill=tk.X, pady=(4, 0))
    tk.Label(m_row, text="➕ 개별 수동 등록:", font=FONT_BOLD_10, bg=BG_CARD, fg="#1e293b").pack(side=tk.LEFT, padx=(5, 10))
    
    tk.Label(m_row, text="연락처:", font=FONT_REG_10, bg=BG_CARD, fg="#475569").pack(side=tk.LEFT, padx=2)
    manual_phone_entry = tk.Entry(m_row, font=FONT_REG_10, width=15, bd=1, relief="solid")
    manual_phone_entry.pack(side=tk.LEFT, padx=5)
    
    tk.Label(m_row, text="발송문자:", font=FONT_REG_10, bg=BG_CARD, fg="#475569").pack(side=tk.LEFT, padx=(10, 2))
    manual_msg_entry = tk.Entry(m_row, font=FONT_REG_10, width=45, bd=1, relief="solid")
    manual_msg_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    
    create_btn(m_row, "➕ 리스트 추가", add_manual_data, COLOR_PRIMARY, COLOR_PRIMARY_HOVER, font=FONT_BOLD_9).pack(side=tk.RIGHT, padx=5)
    
    filename_label=tk.Label(root,text="연동 상태: 로드된 파일 없음",font=FONT_REG_9,fg=COLOR_MUTED,bg=BG_MAIN)
    filename_label.pack(anchor="w",padx=25,pady=(2,5))

    # 2. 작업 모드 및 옵션
    mode_frame=tk.LabelFrame(root,text=" 2. 작업 모드 설정 ",font=FONT_BOLD_11,bg=BG_CARD,fg=COLOR_NAVY,bd=1,relief="solid",padx=15,pady=8)
    mode_frame.pack(fill=tk.X,pady=(2,0),padx=20)
    
    work_mode=tk.StringVar(value="번호등록")
    tk.Radiobutton(mode_frame,text="📋 수동 번호등록 모드",variable=work_mode,value="번호등록",command=toggle_work_mode,font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b",activebackground=BG_CARD,selectcolor=BG_CARD).pack(side=tk.LEFT,padx=(0,30))
    tk.Radiobutton(mode_frame,text="💬 자동 문자발송 모드",variable=work_mode,value="문자발송",command=toggle_work_mode,font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b",activebackground=BG_CARD,selectcolor=BG_CARD).pack(side=tk.LEFT)

    excel_writeback_var=tk.BooleanVar(value=True)
    tk.Checkbutton(mode_frame,text="📊 엑셀 D열에 결과 실시간 기록",variable=excel_writeback_var,font=FONT_BOLD_9,bg=BG_CARD,fg="#1e293b",activebackground=BG_CARD,selectcolor=BG_CARD).pack(side=tk.RIGHT,padx=10)

    # 2-A 번호등록 옵션
    reg_opts_frame=tk.LabelFrame(root,text=" 번호등록 옵션 ",font=FONT_BOLD_10,bg=BG_CARD,fg=COLOR_NAVY,bd=1,relief="solid",padx=15,pady=8)
    reg_opts_frame.pack(fill=tk.X,pady=(0,5),padx=20)
    
    tk.Label(reg_opts_frame,text="등록 개수:",font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b").pack(side=tk.LEFT)
    batch_var=tk.StringVar(value="100")
    tk.Spinbox(reg_opts_frame,from_=10,to=5000,increment=10,textvariable=batch_var,width=6,font=FONT_BOLD_10,justify="center",bd=1,relief="solid").pack(side=tk.LEFT,padx=5)
    tk.Label(reg_opts_frame,text="개씩 등록 진행",font=FONT_REG_10,bg=BG_CARD,fg=COLOR_MUTED).pack(side=tk.LEFT,padx=3)

    # 2-B 문자발송 옵션
    sms_opts_frame=tk.LabelFrame(root,text=" 문자발송 옵션 ",font=FONT_BOLD_10,bg=BG_CARD,fg=COLOR_NAVY,bd=1,relief="solid",padx=15,pady=8)
    
    st=tk.Frame(sms_opts_frame,bg=BG_CARD)
    st.pack(fill=tk.X,pady=(0,4))
    
    send_mode=tk.StringVar(value="개별")
    tk.Radiobutton(st,text="🗂️ 개별문구 발송",variable=send_mode,value="개별",command=toggle_send_mode,font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b",activebackground=BG_CARD,selectcolor=BG_CARD).pack(side=tk.LEFT,padx=(0,12))
    tk.Radiobutton(st,text="📢 일괄문구 발송",variable=send_mode,value="일괄",command=toggle_send_mode,font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b",activebackground=BG_CARD,selectcolor=BG_CARD).pack(side=tk.LEFT,padx=(0,20))
    
    tk.Label(st,text="⏱️ 발송 딜레이:",font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b").pack(side=tk.LEFT,padx=(8,3))
    delay_var=tk.StringVar(value="3.0")
    tk.Spinbox(st,from_=0.5,to=10,increment=0.5,textvariable=delay_var,width=5,font=FONT_BOLD_10,justify="center",bd=1,relief="solid").pack(side=tk.LEFT)
    tk.Label(st,text="초",font=FONT_REG_10,bg=BG_CARD,fg=COLOR_MUTED).pack(side=tk.LEFT,padx=(2,15))

    tk.Label(st,text="🔄 자동 재시도:",font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b").pack(side=tk.LEFT,padx=(8,3))
    retry_var=tk.StringVar(value="1")
    tk.Spinbox(st,from_=0,to=5,increment=1,textvariable=retry_var,width=4,font=FONT_BOLD_10,justify="center",bd=1,relief="solid").pack(side=tk.LEFT)
    tk.Label(st,text="회",font=FONT_REG_10,bg=BG_CARD,fg=COLOR_MUTED).pack(side=tk.LEFT,padx=2)

    # 문구 템플릿 영역
    tpl_frame=tk.Frame(sms_opts_frame,bg=BG_CARD)
    tpl_frame.pack(fill=tk.X,pady=(4,2))
    
    tk.Label(tpl_frame,text="📝 템플릿 선택:",font=FONT_BOLD_10,bg=BG_CARD,fg="#1e293b").pack(side=tk.LEFT,padx=(0,5))
    tpl_var=tk.StringVar(value="-- 템플릿 선택 --")
    
    tpl_dropdown=ttk.Combobox(tpl_frame,textvariable=tpl_var,state="readonly",width=25,style="TCombobox")
    tpl_dropdown.pack(side=tk.LEFT,padx=3)
    tpl_dropdown.bind("<<ComboboxSelected>>",apply_template)
    
    create_btn(tpl_frame, "적용", apply_template, COLOR_PRIMARY, COLOR_PRIMARY_HOVER, font=FONT_BOLD_9).pack(side=tk.LEFT,padx=2)
    create_btn(tpl_frame, "💾 현재 문구 저장", add_template, COLOR_SUCCESS, COLOR_SUCCESS_HOVER, font=FONT_BOLD_9).pack(side=tk.LEFT,padx=2)
    create_btn(tpl_frame, "🗑 삭제", delete_template, COLOR_DANGER, COLOR_DANGER_HOVER, font=FONT_BOLD_9).pack(side=tk.LEFT,padx=2)

    fixed_msg_text=tk.Text(sms_opts_frame,height=3,font=FONT_REG_10,bd=1,relief="solid",padx=8,pady=5)
    fixed_msg_text.pack(fill=tk.X,pady=(5,2))
    fixed_msg_text.insert(tk.END,"[KT텔레캅] 안녕하세요. 미결제금 확인 요청드립니다.")
    fixed_msg_text.config(state=tk.DISABLED,bg="#f1f5f9")

    # 3. 단건 즉시 발송 (좌표 매크로)
    qf=tk.LabelFrame(root,text=" 3. 단건 즉시 발송 (좌표 매크로) ",font=FONT_BOLD_11,bg=BG_CARD,fg=COLOR_NAVY,bd=1,relief="solid",padx=15,pady=8)
    qf.pack(fill=tk.X,pady=(2,5),padx=20)
    
    tk.Label(qf, text="연락처:", font=FONT_REG_10, bg=BG_CARD, fg="#1e293b").pack(side=tk.LEFT, padx=2)
    quick_phone_entry = tk.Entry(qf, font=FONT_REG_10, width=15, bd=1, relief="solid")
    quick_phone_entry.pack(side=tk.LEFT, padx=5)
    
    tk.Label(qf, text="발송문자:", font=FONT_REG_10, bg=BG_CARD, fg="#1e293b").pack(side=tk.LEFT, padx=(10, 2))
    quick_msg_entry = tk.Entry(qf, font=FONT_REG_10, width=40, bd=1, relief="solid")
    quick_msg_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    
    create_btn(qf, "📍 발송 좌표 설정", setup_coord_sms, COLOR_MUTED, COLOR_MUTED_HOVER, font=FONT_BOLD_9).pack(side=tk.RIGHT, padx=3)
    create_btn(qf, "⚡ 즉시 발송 실행", quick_send_action, COLOR_DANGER, COLOR_DANGER_HOVER, font=FONT_BOLD_9).pack(side=tk.RIGHT, padx=3)

    # 4. 테이블 리스트
    tbf=tk.Frame(root,bg=BG_MAIN)
    tbf.pack(fill=tk.BOTH,expand=True,padx=20,pady=5)
    
    ls=tk.LabelFrame(tbf,text=" ⏳ 발송 대기 목록 ",fg=COLOR_PRIMARY,font=FONT_BOLD_11,bg=BG_CARD,bd=1,relief="solid",padx=8,pady=8)
    ls.pack(side=tk.LEFT,fill=tk.BOTH,expand=True,padx=(0,5))
    
    lsc=tk.Scrollbar(ls,orient="vertical")
    left_table=ttk.Treeview(ls,columns=("n","p","m"),show="headings",yscrollcommand=lsc.set,style="Treeview")
    lsc.config(command=left_table.yview)
    left_table.heading("n",text="No"); left_table.heading("p",text="연락처"); left_table.heading("m",text="발송 문구")
    left_table.column("n",width=50,anchor="center"); left_table.column("p",width=120,anchor="center"); left_table.column("m",width=280)
    lsc.pack(side=tk.RIGHT,fill=tk.Y); left_table.pack(fill=tk.BOTH,expand=True,padx=2,pady=2)
    
    rs=tk.LabelFrame(tbf,text=" ✅ 처리 결과 리스트 ",fg=COLOR_SUCCESS,font=FONT_BOLD_11,bg=BG_CARD,bd=1,relief="solid",padx=8,pady=8)
    rs.pack(side=tk.RIGHT,fill=tk.BOTH,expand=True,padx=(5,0))
    
    rsc=tk.Scrollbar(rs,orient="vertical")
    right_table=ttk.Treeview(rs,columns=("n","p","s"),show="headings",yscrollcommand=rsc.set,style="Treeview")
    rsc.config(command=right_table.yview)
    right_table.heading("n",text="No"); right_table.heading("p",text="연락처"); right_table.heading("s",text="처리 상태")
    right_table.column("n",width=50,anchor="center"); right_table.column("p",width=130,anchor="center"); right_table.column("s",width=120,anchor="center")
    rsc.pack(side=tk.RIGHT,fill=tk.Y); right_table.pack(fill=tk.BOTH,expand=True,padx=2,pady=2)

    # 4. 진행률 (Progress Bar)
    pf=tk.Frame(root,bg=BG_MAIN)
    pf.pack(fill=tk.X,padx=20,pady=(5,3))
    
    progress_var=tk.IntVar(value=0)
    ttk.Progressbar(pf,variable=progress_var,maximum=100,length=400,style="TProgressbar").pack(side=tk.LEFT,fill=tk.X,expand=True,padx=(0,10))
    
    lbl_progress=tk.Label(pf,text="  완료 0 | 실패 0 | 남은 0  (0%)",font=FONT_BOLD_10,bg=BG_MAIN,fg=COLOR_PRIMARY)
    lbl_progress.pack(side=tk.RIGHT)

    # 5. 컨트롤 패널
    ct=tk.Frame(root,bg=BG_MAIN)
    ct.pack(fill=tk.X,padx=20,pady=5)
    
    btn_start=create_btn(ct, "▶ 번호등록 실행", start_worker, COLOR_SUCCESS, COLOR_SUCCESS_HOVER, width=16, font=FONT_BOLD_10)
    btn_start.pack(side=tk.LEFT,padx=3)
    
    btn_pause=create_btn(ct, "⏸ 일시정지", lambda:set_pause_state(True), "#f59e0b", "#d97706", font=FONT_BOLD_10, state=tk.DISABLED)
    btn_pause.pack(side=tk.LEFT,padx=2)
    
    btn_resume=create_btn(ct, "▶ 재개", lambda:set_pause_state(False), "#06b6d4", "#0891b2", font=FONT_BOLD_10, state=tk.DISABLED)
    btn_resume.pack(side=tk.LEFT,padx=2)
    
    btn_stop=create_btn(ct, "⏹ 중지", lambda:stop_run(), COLOR_DANGER, COLOR_DANGER_HOVER, font=FONT_BOLD_10, state=tk.DISABLED)
    btn_stop.pack(side=tk.LEFT,padx=2)
    
    tk.Frame(ct,width=10,bg=BG_MAIN).pack(side=tk.LEFT)
    
    cfg_macro = load_macro_config()
    reg_set = all(s.get("coord") is not None for s in cfg_macro["번호등록"] if s.get("type") in ["click_type", "click_only"])
    sms_set = all(s.get("coord") is not None for s in cfg_macro["문자발송"] if s.get("type") in ["click_type", "click_only"])
    btn_text = "🔄 좌표 재설정 (설정됨)" if (reg_set and sms_set) else "🔄 좌표 재설정 (미설정)"
    btn_coord_reset=create_btn(ct, btn_text, reset_coordinates, COLOR_MUTED, COLOR_MUTED_HOVER, font=FONT_BOLD_10)
    btn_coord_reset.pack(side=tk.LEFT,padx=3)
    
    create_btn(ct, "↺ 리스트 초기화", reset_all, "#94a3b8", "#64748b", font=FONT_BOLD_10).pack(side=tk.LEFT,padx=3)
    create_btn(ct, "종료", root.destroy, COLOR_DANGER, COLOR_DANGER_HOVER, font=FONT_BOLD_10).pack(side=tk.RIGHT,padx=3)

    # 6. 하단 로그 및 상태바
    lgf=tk.Frame(root,bg=BG_MAIN)
    lgf.pack(fill=tk.X,padx=20,pady=(2,5))
    
    create_btn(lgf, "📄 발송로그.txt 열기", open_log_file, "#64748b", "#475569", font=FONT_BOLD_9).pack(side=tk.LEFT,padx=3)
    create_btn(lgf, "📁 저장 폴더 열기", open_log_folder, "#64748b", "#475569", font=FONT_BOLD_9).pack(side=tk.LEFT,padx=3)

    status_var=tk.StringVar(value="대기 중")
    tk.Label(root,textvariable=status_var,font=FONT_BOLD_9,bg="#cbd5e1",fg=COLOR_NAVY,anchor="w",padx=15,pady=5).pack(fill=tk.X,side=tk.BOTTOM)
    tk.Frame(root,bg="#94a3b8",pady=1).pack(fill=tk.X,side=tk.BOTTOM)

    def set_pause_state(s):
        global is_paused; is_paused=s
        update_status("⏸ 일시정지 상태" if s else "▶ 작업 재개")
        
    def stop_run():
        global is_running; is_running=False; update_status("⏹ 작업 중지됨")

    # ✅ 초기 설정 실행
    toggle_work_mode()
    refresh_template_dropdown()

    # ✅ 이전 세션 복원 시도
    root.after(500, load_session)

    root.mainloop()

if __name__=="__main__":
    if verify_password(): create_gui()
