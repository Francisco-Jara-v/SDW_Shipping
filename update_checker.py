import requests
from version import VERSION

OWNER = "Francisco-Jara-v"
REPO = "SDW_Shipping"

API_URL = f"https://api.github.com/repos/{OWNER}/{REPO}/releases/latest"


def obtener_release():
    try:
        response = requests.get(API_URL, timeout=10)
        response.raise_for_status()

        data = response.json()

        return {
            "version": data["tag_name"].lstrip("v"),
            "nombre": data["name"],
            "descripcion": data["body"],
            "assets": data["assets"],
        }

    except Exception as e:
        print("Error consultando GitHub:", e)
        return None


def comparar_versiones(v1, v2):
    """
    Devuelve True si v1 es mayor que v2
    """
    a = tuple(map(int, v1.split(".")))
    b = tuple(map(int, v2.split(".")))
    return a > b


def hay_actualizacion():
    release = obtener_release()

    if release is None:
        return False, None

    if comparar_versiones(release["version"], VERSION):
        return True, release

    return False, release