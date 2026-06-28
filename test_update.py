from update_checker import hay_actualizacion
from version import VERSION

print("Versión instalada:", VERSION)

hay, release = hay_actualizacion()

if hay:

    print("\nNueva versión encontrada")

    print("Versión:", release["version"])

    print("URL:", release["download_url"])

else:

    print("\nNo hay actualizaciones.")