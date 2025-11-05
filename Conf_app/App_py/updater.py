import requests, os, ssl, subprocess, time

# --- Ignora SSL corporativo (seguro em rede interna) ---
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings()

# --- Configurações principais ---
REPO = "Kvsl11/Auto_form"
URL_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
URL_SCRIPT = f"https://raw.githubusercontent.com/{REPO}/main/main.py"
LOCAL_SCRIPT = "main.py"
LOCAL_VERSION_FILE = "version_local.txt"

# Caminho do Python interno (sem console)
PYTHONW_PATH = os.path.join(os.getcwd(), "Python313", "python.exe")

# --- Funções auxiliares ---
def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "0.0.0"

def get_online_version():
    try:
        headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
        r = requests.get(URL_VERSION, timeout=10, verify=False, headers=headers)
        if r.status_code == 200:
            return r.text.strip()
        print(f"⚠️ Erro HTTP ao buscar versão: {r.status_code}")
    except Exception as e:
        print("⚠️ Erro ao obter versão online:", e)
    return None

def atualizar_script():
    """Baixa a nova versão do main.py diretamente e substitui a existente."""
    try:
        print("⬇️ Baixando nova versão do main.py...")
        headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
        r = requests.get(URL_SCRIPT, timeout=20, verify=False, headers=headers)
        r.raise_for_status()
        conteudo = r.content

        # Remove o main.py antigo (se existir)
        if os.path.exists(LOCAL_SCRIPT):
            os.remove(LOCAL_SCRIPT)

        with open(LOCAL_SCRIPT, "wb") as f:
            f.write(conteudo)

        print("✅ main.py atualizado com sucesso.")
        return True
    except Exception as e:
        print("❌ Erro ao atualizar script:", e)
        return False

def save_local_version(ver):
    with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(ver)
    print(f"💾 Versão local atualizada para: {ver}")

def iniciar_app():
    """Executa o app principal com pythonw.exe sem console."""
    print("🚀 Iniciando app principal...")
    try:
        if os.path.exists(PYTHONW_PATH):
            subprocess.Popen(
                [PYTHONW_PATH, LOCAL_SCRIPT],
                creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW
            )
        else:
            subprocess.Popen(["python", LOCAL_SCRIPT])
    except Exception as e:
        print("❌ Erro ao iniciar o app:", e)
    finally:
        os._exit(0)

# --- Execução principal ---
def main():
    print("🔍 Verificando atualizações...")
    local_v = get_local_version()
    online_v = get_online_version()

    print(f"Versão local: {local_v}")
    print(f"Versão online: {online_v}")

    if not online_v:
        print("⚠️ Falha ao obter versão online. Rodando versão local.")
        iniciar_app()
        return

    if online_v != local_v:
        print(f"🟡 Nova versão detectada: {online_v}. Atualizando automaticamente...")
        ok = atualizar_script()
        if ok:
            save_local_version(online_v)
            print("♻️ Reiniciando com nova versão...")
            time.sleep(1)
            iniciar_app()
        else:
            print("❌ Falha na atualização. Rodando versão atual.")
            iniciar_app()
    else:
        print(f"🟢 Você está usando a versão mais recente ({local_v}).")
        iniciar_app()

if __name__ == "__main__":
    main()
