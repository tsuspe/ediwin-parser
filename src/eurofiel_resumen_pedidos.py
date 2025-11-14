import re
import argparse
from pathlib import Path

import pdfplumber
import pandas as pd


def split_orders(full_text: str):
    """
    Divide el texto completo del PDF en bloques,
    cada uno correspondiente a un pedido (PEDIDO / REEMPLAZO / ANULACIÓN).
    """
    matches = list(re.finditer(r"Nº Pedido\s*:", full_text))
    chunks = []

    for i, m in enumerate(matches):
        start = m.start()
        # Buscamos la línea anterior para incluir el TIPO (PEDIDO, ANULACIÓN PEDIDO…)
        prev_nl = full_text.rfind("\n", 0, start)
        if prev_nl != -1:
            prev_prev_nl = full_text.rfind("\n", 0, prev_nl)
            order_start = prev_prev_nl + 1 if prev_prev_nl != -1 else 0
        else:
            order_start = 0

        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        chunk = full_text[order_start:end].strip()
        chunks.append(chunk)

    return chunks


def parse_detail_line(line: str):
    """
    Parsea una línea de detalle de artículo.

    Ejemplos:
    1 8447571299747 3RC240/NARANJA/XS 0863769/66/01 1 50 50 0 EUR
    1 8447571186818 2TB060/AZUL OSCUR/S 0832547/11/04 4 27 27 189 EUR

    Devuelve (MODELO, PATRON, PRECIO) o None si la línea no es de detalle.
    """
    parts = line.split()
    if len(parts) < 8:
        return None

    # Índice de línea
    if not parts[0].isdigit():
        return None

    # EAN 13
    if not re.fullmatch(r"\d{13}", parts[1]):
        return None

    # Buscar el primer token que tenga formato d+/d+/d+ => Cod Cliente/Color/Talla
    cli_idx = None
    for i in range(2, len(parts)):
        if re.fullmatch(r"\d+/\d+/\d+", parts[i]):
            cli_idx = i
            break

    if cli_idx is None or cli_idx + 4 >= len(parts):
        return None

    # Cod Proveedor/Color/Talla puede tener espacios (ej. "2TB060/AZUL OSCUR/S")
    cod_prov_full = " ".join(parts[2:cli_idx])
    cod_cli_full = parts[cli_idx]

    # Luego vienen: Cantidad, P.Bruto, P.Neto, PVP, (Dto), Moneda
    qty = parts[cli_idx + 1]
    p_bruto = parts[cli_idx + 2]
    p_neto = parts[cli_idx + 3]
    pvp = parts[cli_idx + 4]

    # Modelo = Cod Proveedor/Color (quitamos la talla final /XXS, /S…)
    modelo = re.sub(r"/[^/]+$", "", cod_prov_full)
    # Patrón = Cod Cliente/Color (quitamos la talla final /01, /04…)
    patron = re.sub(r"/[^/]+$", "", cod_cli_full)

    # Usamos P.Neto como PRECIO
    precio = p_neto.replace(",", ".")

    return modelo, patron, precio


def parse_order(order_text: str):
    """
    Parsea un bloque de texto correspondiente a un solo pedido.
    Devuelve un dict con todos los campos necesarios para el Excel.
    """
    lines = [ln for ln in order_text.splitlines() if ln.strip()]
    first_line = lines[0].strip() if lines else ""
    tipo = first_line  # PEDIDO / REEMPLAZO PEDIDO / ANULACIÓN PEDIDO

    def search(pattern: str):
        m = re.search(pattern, order_text)
        return m.group(1).strip() if m else ""

    pedido = search(r"Nº Pedido\s*:\s*(\S+)")
    fecha_entrega = search(r"Fecha Entrega\s*:\s*(\d{2}/\d{2}/\d{4})")

    # País: ( CR ) COSTA RICA  -> nos quedamos con "COSTA RICA"
    pais = ""
    m_pais = re.search(r"País:\s*\([^)]*\)\s*([A-ZÁÉÍÓÚÜÑ ]+)", order_text)
    if m_pais:
        pais = m_pais.group(1).strip()

    descripcion = search(r"Descripción:\s*(.+)")
    total_unidades = search(r"Total Unidades\s+(\d+)")

    modelo = ""
    patron = ""
    precio = ""

    # Buscamos la primera línea de detalle válida
    for ln in lines:
        parsed = parse_detail_line(ln)
        if parsed:
            modelo, patron, precio = parsed
            break

    return {
        "TIPO": tipo,
        "PEDIDO": pedido,
        "FECHA_ENTREGA": fecha_entrega,
        "PAIS": pais,
        "DESCRIPCION": descripcion,
        "MODELO": modelo,
        "PATRON": patron,
        "PRECIO": precio,
        "TOTAL_UNIDADES": total_unidades,
    }


def parse_pdf(path: Path):
    """
    Lee el PDF completo, lo trocea en pedidos y devuelve
    una lista de dicts (uno por pedido).
    """
    with pdfplumber.open(path) as pdf:
        full_text = "\n".join(page.extract_text() for page in pdf.pages)

    orders = split_orders(full_text)
    rows = [parse_order(o) for o in orders]

    # Quitamos posibles bloques raros sin nº de pedido
    rows = [r for r in rows if r.get("PEDIDO")]
    return rows


def main():
    parser = argparse.ArgumentParser(description="Resumen de pedidos EUROFIEL desde PDF EDIWIN")
    parser.add_argument("--pdf", required=True, help="Ruta al PDF de EUROFIEL (EDIWIN)")
    parser.add_argument("--out", required=True, help="Ruta de salida del Excel")

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    out_path = Path(args.out)

    if not pdf_path.exists():
        raise SystemExit(f"No se encuentra el PDF: {pdf_path}")

    rows = parse_pdf(pdf_path)

    # Orden de columnas en el Excel final
    columns = [
        "TIPO",
        "PEDIDO",
        "FECHA_ENTREGA",
        "PAIS",
        "DESCRIPCION",
        "MODELO",
        "PATRON",
        "PRECIO",
        "TOTAL_UNIDADES",
    ]

    df = pd.DataFrame(rows, columns=columns)

    # Creamos carpeta de salida si no existe
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_path, index=False)
    print(f"Resumen generado con {len(df)} pedidos -> {out_path}")


if __name__ == "__main__":
    main()
