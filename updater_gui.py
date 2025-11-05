import os
import requests
import subprocess
import ssl
import time

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
        headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
        r = requests.get(BASE_URL + VERSION_FILE, timeout=10, verify=False, headers=headers)
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
        print("⬇️ Baixando nova versão do main.py...")
        headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
        r = requests.get(BASE_URL + SCRIPT_FILE, timeout=20, verify=False, headers=headers)
        r.raise_for_status()
        with open(LOCAL_SCRIPT, "wb") as f:
            f.write(r.content)
        save_local_version(version)
        print(f"✅ Atualizado para versão {version}.")
        return True
    except Exception as e:
        print(f"❌ Falha ao baixar nova versão: {e}")
        return False

def iniciar_app():
    """Executa o app principal (usando o Python embutido, se existir)."""
    print("🚀 Iniciando o aplicativo principal...")
    try:
        if os.path.exists(PYTHON_EMBUTIDO):
            subprocess.Popen(
                [PYTHON_EMBUTIDO, LOCAL_SCRIPT],
                creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW
            )
        else:
            subprocess.Popen(["python", LOCAL_SCRIPT])
    except Exception as e:
        print(f"❌ Erro ao iniciar o app: {e}")
    finally:
        os._exit(0)

def check_update():
    """Verifica se há atualização e baixa automaticamente, se necessário."""
    print("🔍 Verificando atualizações...")
    local_v = get_local_version()
    remote_v = get_remote_version()

    print(f"Versão local: {local_v}")
    print(f"Versão online: {remote_v}")

    if not remote_v:
        print("⚠️ Não foi possível verificar a versão online. Executando versão local.")
        iniciar_app()
        return

    if local_v != remote_v:
        print(f"🟡 Nova versão detectada ({remote_v}), iniciando atualização automática...")
        if download_main(remote_v):
            print("♻️ Reiniciando com a nova versão...")
            time.sleep(1)
            iniciar_app()
        else:
            print("❌ Falha ao atualizar. Executando versão local.")
            iniciar_app()
    else:
        print(f"🟢 Versão atual ({local_v}) já está atualizada.")
        iniciar_app()

# --- Execução principal ---
if __name__ == "__main__":
    check_update()
