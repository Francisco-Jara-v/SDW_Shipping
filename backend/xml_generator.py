from datetime import datetime
import re

# -------------------------
# 🔥 ITEMS DINÁMICOS (SIN VEHÍCULO NI PESO)
# -------------------------
def generar_items_xml(items, datos):
    xml = ""

    for i, item in enumerate(items, start=1):
        contenedor = item.get("delivered_container", "")

        match = re.match(r'^([A-Za-z]{4})(\d{6})(\d)$', contenedor)

        if match:
            sigla = match.group(1)
            numero = match.group(2)
            digito = match.group(3)

            carga_cnt = "S"

            contenedor_xml = f"""
                <Contenedores>
                    <contenedor>
                        <sigla>{sigla}</sigla>
                        <numero>{numero}</numero>
                        <digito>{digito}</digito>
                        <tipo-cnt>42P3</tipo-cnt>
                        <peso>{int(datos['peso_total'])}</peso>
                        <nombre-operador>{sigla}</nombre-operador>
                        <status>LCL/LCL</status>
                    </contenedor>
                </Contenedores>
            """
        else:
            carga_cnt = "N"
            contenedor_xml = ""

        xml += f"""
        <item>
            <numero-item>{i}</numero-item>
            <marcas>NO MARCAS</marcas>
            <carga-peligrosa>N</carga-peligrosa>
            <tipo-bulto>85</tipo-bulto>
            <descripcion>{item.get('specification', '')}</descripcion>
            <cantidad>1</cantidad>
            <peso-bruto>{int(datos['peso_total'])}</peso-bruto>
            <unidad-peso>KGM</unidad-peso>
            <observaciones>{item.get('specification', '')}</observaciones>
            <carga-cnt>{carga_cnt}</carga-cnt>

            {contenedor_xml if carga_cnt == "S" else ""}
            
            <Vehiculos>
                <vehiculo>
                    <modelo>{datos.get('type', '') or "S/M"}</modelo>
                    <chassis>{datos.get('vin', '')}</chassis>
                    <marca>{datos.get('make', '')}</marca>
                    <observacion>{item.get('specification', '')}</observacion>
                </vehiculo>
            </Vehiculos>

        </item>
        """
    return xml.strip()


# -------------------------
# 🚛 VEHÍCULO (FUERA DEL LOOP)
# -------------------------
def generar_vehiculo_xml(datos):
    return f"""
    <Vehiculos>
        <vehiculo>
            <modelo>{datos.get('type', '') or "S/M"}</modelo>
            <chassis>{datos.get('vin', '')}</chassis>
            <marca>{datos.get('make', '')}</marca>
            <observacion>{datos.get('items')[0].get('specification', '')}</observacion>
        </vehiculo>
    </Vehiculos>
    """


# -------------------------
# 🔥 GENERADOR PRINCIPAL
# -------------------------
def generar_xml(template_path, output_path, datos):
    # leer template
    with open(template_path, "r", encoding="ISO-8859-1") as f:
        xml = f.read()

    # generar bloques
    items_xml = generar_items_xml(datos["items"], datos)
    vehiculo_xml = generar_vehiculo_xml(datos)
    transit = (datos.get("transit") or "").strip()

    shipping = "" if transit else "Y/O SDW SHIPPING CHILE SPA"
    sentido = "TR" if transit else "I"

    # reemplazos
    reemplazos = {
        "${bl}": datos["bl"],
        "${bultos}": str(datos["bultos"]),
        "${peso_total}": str(int(datos["peso_total"])),
        "${consignee}": datos["consignee"],
        "${pol}": datos["pol"],
        "${pod}": datos["pod"],
        "${fecha_bl}": datos["fecha_bl"],
        "${fecha_pres}": datos["fecha_pres"],
        "${nave}": datos["nave"],
        "${items_xml}": items_xml,
        "${vehiculo_xml}": vehiculo_xml,
        "${shipping}": shipping,
        "${sentido}": sentido
    }

    # aplicar reemplazos
    for key, value in reemplazos.items():
        xml = xml.replace(key, str(value))

    # guardar
    with open(output_path, "w", encoding="ISO-8859-1") as f:
        f.write(xml)