import requests, os, subprocess

GITHUB_USER = "Kvsl11"  # <-- substitua
REPO_NAME = "Auto_form"

BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/main/"
version = requests.get(BASE_URL + "version.txt").text.strip()
file_name = f"Version_{version}.py"

print(f"Baixando versão {version}...")
r = requests.get(BASE_URL + file_name)
open(file_name, "wb").write(r.content)
print("Executando aplicação...")
subprocess.Popen(["python", file_name])