import os
import requests
import tkinter as tk
from tkinter import messagebox
import subprocess
import ssl

# --- Ignorar SSL corporativo (opcional para redes com proxy) ---
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings()

# --- Configurações principais ---
GITHUB_USER = "Kvsl11"
REPO_NAME = "Auto_form"
BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/main/"
VERSION_FILE = "version.txt"
SCRIPT_FILE = "main.py"

# --- Caminhos locais ---
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_SCRIPT = os.path.join(APP_DIR, SCRIPT_FILE)
LOCAL_VERSION_FILE = os.path.join(APP_DIR, VERSION_FILE)
PYTHON_EMBUTIDO = os.path.join(APP_DIR, "Python313", "python.exe")

# --- Funções auxiliares ---
def get_remote_version():
    """Obtém a versão online do repositório GitHub."""
    try:
        r = requests.get(BASE_URL + VERSION_FILE, timeout=8, verify=False)
        if r.status_code == 200:
            return r.text.strip()
    except Exception as e:
        print(f"⚠️ Erro ao obter versão remota: {e}")
    return None

def get_local_version():
    """Lê a versão local salva em version.txt."""
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "0.0.0"

def save_local_version(version):
    """Atualiza o arquivo version.txt local."""
    with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(version)

def download_main(version):
    """Baixa o novo main.py atualizado do GitHub."""
    try:
        r = requests.get(BASE_URL + SCRIPT_FILE, timeout=15, verify=False)
        r.raise_for_status()
        with open(LOCAL_SCRIPT, "wb") as f:
            f.write(r.content)
        save_local_version(version)
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"⚠️ Falha ao baixar nova versão: {e}")
        return False

def iniciar_app():
    """Executa o app principal (usando o Python embutido, se existir)."""
    if os.path.exists(PYTHON_EMBUTIDO):
        subprocess.Popen([PYTHON_EMBUTIDO, LOCAL_SCRIPT])
    else:
        subprocess.Popen(["python", LOCAL_SCRIPT])
    exit()

def check_update():
    """Verifica se há atualização e baixa automaticamente, se necessário."""
    local_v = get_local_version()
    remote_v = get_remote_version()

    if not remote_v:
        messagebox.showwarning("Aviso", "⚠️ Não foi possível verificar a versão online.")
        iniciar_app()
        return

    if local_v != remote_v:
        if messagebox.askyesno("Atualização disponível",
                               f"Versão atual: {local_v}\nNova versão: {remote_v}\n\nDeseja atualizar agora?"):
            if download_main(remote_v):
                messagebox.showinfo("Sucesso", f"✅ Atualizado para a versão {remote_v} com sucesso!")
            else:
                messagebox.showerror("Erro", "Falha ao atualizar o aplicativo.")
        else:
            messagebox.showinfo("Ignorado", "🔹 Atualização ignorada pelo usuário.")

    iniciar_app()

# --- Execução principal ---
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    check_update()
