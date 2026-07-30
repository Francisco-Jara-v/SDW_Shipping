import requests
from packaging import version
from version import VERSION

OWNER = "Francisco-Jara-v"
REPO = "SDW_Shipping"
API_URL = f"https://api.github.com/repos/{OWNER}/{REPO}/releases/latest"

def obtener_release():
    try:
        response = requests.get(API_URL, timeout=10)
        response.raise_for_status()
        data = response.json()

        download_url = None
        # Buscamos el ejecutable o archivo de instalación dentro del Release
        for asset in data.get("assets", []):
            if asset["name"].endswith(".exe") or asset["name"] == "SDW.zip":
                download_url = asset["browser_download_url"]
                break

        tag_version = data["tag_name"].lstrip("v").strip()

        return {
            "version": tag_version,
            "nombre": data.get("name", ""),
            "descripcion": data.get("body", "Sin detalles de actualización."),
            "download_url": download_url,
            "release_url": data.get("html_url")
        }

    except Exception as e:
        print(f"[UpdateChecker] Error consultando GitHub: {e}")
        return None

def hay_actualizacion():
    release = obtener_release()

    if release is None:
        return False, None

    try:
        # Uso de packaging.version para manejar semantic versioning correctamente
        if version.parse(release["version"]) > version.parse(VERSION):
            return True, release
    except Exception as e:
        print(f"[UpdateChecker] Error comparando versiones: {e}")

    return False, release