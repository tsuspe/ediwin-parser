import re
from pathlib import Path
import argparse

import pdfplumber
import pandas as pd


# ============= PARSER ECI =============

def parse_page_eci(text: str):
    """
    Parsea una página de pedido de ECI.
    Devuelve una lista de dicts, una fila por (pedido, modelo, color).
    """

    # Limpiamos líneas
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    def search(pattern: str):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else ""

    # TIPO: Pedido / Reposición / Anulación Pedido...
    tipo = ""
    for ln in lines:
        low = ln.lower()
        if low in ("pedido", "reposicion", "reposición", "anulacion pedido", "anulación pedido"):
            tipo = ln.upper()
            break

    # Cabecera
    n_pedido = search(r"Nº Pedido\s+(\d+)")
    departamento = search(r"Dpto\. venta\s+(\d+)")
    fecha_entrega = search(r"Fecha Entrega\s+(\d{2}/\d{2}/\d{4})")

    # Sucursal entrega (01 0050, 02 0062, etc.)
    suc_entrega = search(r"Sucursal Destino que Pide\s+([0-9 ]+)\s+[A-ZÁÉÍÓÚÜÑ]")
    if not suc_entrega:
        suc_entrega = search(r"Sucursal de Entrega\s+([0-9 ]+)\s+[A-ZÁÉÍÓÚÜÑ]")

    rows = []

    for i, ln in enumerate(lines):
        # Línea de detalle principal (nº + EAN13 + resto)
        if not re.match(r"^\d+\s+\d{13}\s", ln):
            continue

        parts = ln.split()

        # Índices de tokens numéricos (nos sirven para localizar QTY y precios)
        num_indices = [idx for idx, tok in enumerate(parts)
                       if re.fullmatch(r"[\d.,]+", tok)]
        if len(num_indices) < 6:
            # Si no hay suficientes números, pasamos
            continue

        # Patrón: ... DESCRIPCION QTY 1 P_BRUTO P_NETO PVP NETO_LINEA
        # → los últimos 4 números son precios
        # → antes hay un "1" (factor) y antes la cantidad
        qty_idx = num_indices[-6]
        p_bruto_idx = num_indices[-4]

        qty = parts[qty_idx]
        p_bruto_raw = parts[p_bruto_idx]

        # Normalizamos precio: 53,000 → 53.000
        precio = p_bruto_raw.replace(".", "").replace(",", ".")

        # Descripción: después de los 3 códigos (serie + ref + colorcode)
        # Formato visto: 1 EAN SERIE REF COLORCODE DESCRIPCION... QTY ...
        desc_tokens = parts[5:qty_idx]
        descripcion = " ".join(desc_tokens)

        # Línea siguiente puede ser segunda parte de la descripción (PUNTO ASIM FALDA)
        extra_desc = ""
        if i + 1 < len(lines):
            next_ln = lines[i + 1]
            if (not re.match(r"^\d+\s+\d{13}\s", next_ln)     # no es otra cabecera
                and not re.match(r"^[A-Z0-9]{5,}\s+\d{3}\s", next_ln)  # no es línea de modelo/color
                and "WOMAN FIESTA" not in next_ln):
                extra_desc = next_ln.strip()

        if extra_desc:
            descripcion = f"{descripcion} {extra_desc}"

        # Línea de modelo/color:
        # si hay extra_desc, está 2 líneas debajo; si no, 1 línea debajo
        j = i + 1 + (1 if extra_desc else 0)
        modelo = ""
        color = ""

        if j < len(lines):
            info_ln = lines[j]
            info_parts = info_ln.split()

            # Formato visto: 47D262G 983 PRINT NEGRO003 3
            if len(info_parts) >= 3 and re.fullmatch(r"[A-Z0-9]+", info_parts[0]):
                modelo = info_parts[0]
                color_code = info_parts[1] if len(info_parts) >= 2 else ""
                color_name1 = info_parts[2] if len(info_parts) >= 3 else ""
                color_name2 = ""
                if len(info_parts) >= 4:
                    # NEGRO003 → NEGRO
                    color_name2 = re.sub(r"\d+$", "", info_parts[3])
                color = " ".join([p for p in [color_code, color_name1, color_name2] if p])

        rows.append({
            "TIPO": tipo,                         # PEDIDO / REPOSICION / ANULACION...
            "N_PEDIDO": n_pedido,                 # 74245201
            "DEPARTAMENTO": departamento,         # 0056
            "DESCRIPCION": descripcion,           # VEST LARGO... PUNTO ASIM FALDA
            "MODELO": modelo,                     # 47D262G
            "COLOR": color,                       # 983 PRINT NEGRO
            "PRECIO": precio,                     # 53.000 (P. Bruto)
            "FECHA_ENTREGA": fecha_entrega,       # 06/02/2025
            "SUC_ENTREGA": suc_entrega,           # 01 0050 / 02 0062...
            "TOTAL_UNIDADES": qty,                # 134 / 125 / 34 / 31...
        })

    return rows


def parse_pdf_eci(pdf_path: Path) -> pd.DataFrame:
    """
    Abre un PDF de ECI y devuelve un DataFrame con:
    - 1 fila por (PEDIDO + MODELO + COLOR)
    - TOTAL_UNIDADES: suma por modelo/color/pedido
    """
    all_rows = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_rows.extend(parse_page_eci(text))

    if not all_rows:
        return pd.DataFrame()

    df = pd.DataFrame(all_rows)

    # Total unidades como número
    df["TOTAL_UNIDADES"] = pd.to_numeric(df["TOTAL_UNIDADES"],
                                         errors="coerce").fillna(0).astype(int)

    # Agrupamos por pedido + modelo + color (por si en algún PDF repiten línea)
    group_cols = [
        "TIPO",
        "N_PEDIDO",
        "DEPARTAMENTO",
        "DESCRIPCION",
        "MODELO",
        "COLOR",
        "PRECIO",
        "FECHA_ENTREGA",
        "SUC_ENTREGA",
    ]

    df_grouped = (
        df.groupby(group_cols, as_index=False)["TOTAL_UNIDADES"]
          .sum()
    )

    return df_grouped


# ============= CLI =============

def main():
    parser = argparse.ArgumentParser(
        description="Extrae resumen de pedidos ECI (N_PEDIDO, MODELO, COLOR, unidades...) desde un PDF."
    )
    parser.add_argument(
        "--pdf",
        required=True,
        help="Ruta al PDF de El Corte Inglés (formato EDIWIN)."
    )
    parser.add_argument(
        "--out",
        required=True,
        help="Ruta de salida para el Excel (por ejemplo, output/eci_resumen.xlsx)."
    )

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = parse_pdf_eci(pdf_path)

    if df.empty:
        print("⚠ No se han detectado líneas de pedido en el PDF ECI.")
        return

    df.to_excel(out_path, index=False)
    print(f"✅ Exportado {len(df)} filas a: {out_path}")


if __name__ == "__main__":
    main()
