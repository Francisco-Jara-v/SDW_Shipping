from update_checker import hay_actualizacion
from version import VERSION

print(f"Versión instalada: {VERSION}")

hay, release = hay_actualizacion()

if hay:
    print("\nHay una nueva versión disponible.")
    print(f"Versión: {release['version']}")
    print(f"Nombre: {release['nombre']}")

else:
    print("\nYa tienes la última versión.")