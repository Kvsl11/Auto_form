import requests
import os
import ssl
import subprocess

# --- Ignora verificação SSL (caso rede corporativa bloqueie) ---
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings()

# --- Configurações principais ---
REPO_USER = "Kvsl11"
REPO_NAME = "Auto_form"
LOCAL_SCRIPT = "main.py"
LOCAL_VERSION_FILE = "version_local.txt"

URL_VERSION_ONLINE = f"https://raw.githubusercontent.com/{REPO_USER}/{REPO_NAME}/main/version.txt"
URL_SCRIPT_ONLINE = f"https://raw.githubusercontent.com/{REPO_USER}/{REPO_NAME}/main/{LOCAL_SCRIPT}"

def ler_versao_local():
    """Lê a versão local salva no arquivo version_local.txt."""
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "0.0.0"

def salvar_versao_local(versao):
    """Atualiza o arquivo version_local.txt com a nova versão."""
    with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(versao)

def obter_versao_online():
    """Obtém a versão mais recente publicada no GitHub."""
    try:
        resp = requests.get(URL_VERSION_ONLINE, timeout=10, verify=False)
        if resp.status_code == 200:
            return resp.text.strip()
        else:
            print(f"⚠️ Erro ao ler versão online: HTTP {resp.status_code}")
    except Exception as e:
        print(f"⚠️ Falha ao verificar versão online: {e}")
    return None

def baixar_script():
    """Baixa o script principal (main.py) do GitHub."""
    try:
        print(f"⬇ Baixando {URL_SCRIPT_ONLINE} ...")
        resp = requests.get(URL_SCRIPT_ONLINE, timeout=15, verify=False)
        resp.raise_for_status()
        with open(LOCAL_SCRIPT, "wb") as f:
            f.write(resp.content)
        print("✅ Atualização baixada com sucesso.")
        return True
    except Exception as e:
        print(f"❌ Erro ao baixar script: {e}")
        return False

def executar_script():
    """Executa o script principal com o Python atual."""
    print("🚀 Iniciando Auto_form...")
    subprocess.run(["python", LOCAL_SCRIPT])

if __name__ == "__main__":
    print("🔍 Verificando atualizações...")
    versao_local = ler_versao_local()
    versao_online = obter_versao_online()

    if not versao_online:
        print("⚠️ Não foi possível verificar atualização. Executando versão local.")
        executar_script()
    elif versao_online != versao_local:
        print(f"🟡 Nova versão detectada: {versao_online} (local: {versao_local})")
        if baixar_script():
            salvar_versao_local(versao_online)
            print("🟢 Versão atualizada. Executando nova versão...")
            executar_script()
        else:
            print("⚠️ Falha ao atualizar. Executando versão local.")
            executar_script()
    else:
        print(f"🟢 Aplicação já está atualizada (v{versao_local}).")
        executar_script()
