import requests, os, ssl, subprocess, tkinter as tk
from tkinter import messagebox

# --- Ignora SSL corporativo (seguro em rede interna) ---
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings()

# --- Configurações principais ---
REPO = "Kvsl11/Auto-_form"
URL_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
URL_SCRIPT = f"https://raw.githubusercontent.com/{REPO}/main/main.py"
LOCAL_SCRIPT = "main.py"
LOCAL_VERSION_FILE = "version_local.txt"

# Caminho do Python interno (sem console)
PYTHONW_PATH = os.path.join(os.getcwd(), "Python313", "pythonw.exe")

# --- Funções auxiliares ---
def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "0.0.0"

def get_online_version():
    try:
        r = requests.get(URL_VERSION, timeout=10, verify=False)
        if r.status_code == 200:
            return r.text.strip()
    except Exception as e:
        print("Erro ao obter versão online:", e)
    return None

def atualizar_script():
    """Baixa a nova versão do main.py diretamente e substitui a existente."""
    try:
        r = requests.get(URL_SCRIPT, timeout=20, verify=False)
        r.raise_for_status()
        conteudo = r.content

        # Remove o main.py antigo (se existir)
        if os.path.exists(LOCAL_SCRIPT):
            os.remove(LOCAL_SCRIPT)

        # Cria o novo main.py atualizado
        with open(LOCAL_SCRIPT, "wb") as f:
            f.write(conteudo)
        return True
    except Exception as e:
        print("Erro ao atualizar script:", e)
        return False

def save_local_version(ver):
    with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(ver)

# --- Interface visual de atualização ---
def show_update_ui(local_v, online_v):
    win = tk.Tk()
    win.title("Atualização disponível - Auto_form")
    win.geometry("400x300")
    win.configure(bg="#232323")
    win.resizable(False, False)

    tk.Label(win, text="🔄 Atualização disponível!",
             fg="white", bg="#232323", font=("Segoe UI", 13, "bold")).pack(pady=15)
    tk.Label(win, text=f"Versão atual: {local_v}\nNova versão: {online_v}",
             fg="#ccc", bg="#232323").pack(pady=5)
    status_label = tk.Label(win, text="", fg="#3ba55d", bg="#232323")
    status_label.pack(pady=5)

    def atualizar_agora():
        btn_update.config(state="disabled", text="Baixando...")
        win.update()
        ok = atualizar_script()
        if ok:
            save_local_version(online_v)
            status_label.config(text="✅ Atualização concluída!")
            win.update()
            messagebox.showinfo("Atualizado", "O aplicativo foi atualizado com sucesso!")
            win.destroy()
            iniciar_app()
        else:
            messagebox.showerror("Erro", "⚠️ Falha ao baixar atualização. Tente novamente mais tarde.")
            win.destroy()
            iniciar_app()

    btn_update = tk.Button(win, text="Atualizar agora", command=atualizar_agora,
                           bg="#3ba55d", fg="white", font=("Segoe UI", 11, "bold"), width=20)
    btn_update.pack(pady=15)

    tk.Button(win, text="Ignorar", command=lambda: [win.destroy(), iniciar_app()],
              bg="#555", fg="white", width=20).pack(pady=5)
    win.mainloop()

# --- Função para iniciar o app ---
def iniciar_app():
    """Executa o app principal com pythonw.exe sem console."""
    if os.path.exists(PYTHONW_PATH):
        subprocess.Popen(
            [PYTHONW_PATH, LOCAL_SCRIPT],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW
        )
    else:
        subprocess.Popen(["python", LOCAL_SCRIPT])
    os._exit(0)

# --- Execução principal ---
def main():
    print("🔍 Verificando atualizações...")
    local_v = get_local_version()
    online_v = get_online_version()

    if not online_v:
        print("⚠️ Sem conexão ou erro de versão online. Rodando local.")
        iniciar_app()
        return

    if online_v != local_v:
        print(f"🟡 Nova versão detectada: {online_v} (local: {local_v})")
        show_update_ui(local_v, online_v)
    else:
        print(f"🟢 Você está usando a versão mais recente ({local_v}).")
        iniciar_app()

if __name__ == "__main__":
    main()
