#!/usr/bin/env python3
"""
Eurofiel EDIWIN PDF -> CSV/Excel extractor (v2)
- Extrae cabecera (pedido, fecha, fecha_entrega, destino)
- Extrae líneas con fallback por regex (EAN-13 como ancla)
- Detecta MODELO y PATRON (p.ej., 3RC240/NARANJA y 0863769/66)
- Saca PRECIO neto unitario, UNIDADES y calcula UNIDADES_TOTALES e IMPORTE_TOTAL por pedido
Uso:
  python eurofiel_parser.py \
    --pdf input/Ejemplo_PDF_Eurofiel.pdf \
    --map equivalencias/tabla_equivalencias_Eurofiel.xlsx \
    --out output/eurofiel_resultado.xlsx
"""

import argparse, re, shutil
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple
from pathlib import Path
from datetime import datetime
import pdfplumber
import pandas as pd

TARGET_COLS = [
    "TIPO","PEDIDO","FECHA","FECHA_ENTREGA","DESTINO",
    "MODELO","PATRON","UNIDADES","PRECIO",
    "UNIDADES_TOTALES","IMPORTE_TOTAL"
]

@dataclass
class Header:
    tipo: str = "Pedido"
    pedido: str = ""
    fecha: str = ""
    fecha_entrega: str = ""
    destino: str = ""

@dataclass
class Linea:
    modelo: str = ""
    patron: str = ""
    precio: Optional[float] = None
    unidades: Optional[int] = None
    ean: str = ""
    page: int = 0
    desc: str = ""

def parse_equivalences(xlsx_path: Optional[str]) -> Dict[str, Dict[str,str]]:
    eq: Dict[str, Dict[str,str]] = {}
    if not xlsx_path:
        return eq
    df = pd.read_excel(xlsx_path)
    cols = list(df.columns)
    if len(cols) < 3:
        return eq
    group_col, src_col, dst_col = cols[:3]
    for _, row in df.iterrows():
        group = str(row.get(group_col, "")).strip()
        src = str(row.get(src_col, "")).strip()
        dst = str(row.get(dst_col, "")).strip()
        if not group or not src:
            continue
        eq.setdefault(group, {})[src] = dst
    return eq

def apply_eq(eq: Dict[str, Dict[str,str]], group: str, value: str) -> str:
    if not value:
        return value
    return eq.get(group, {}).get(value, value)

# ---------- Helpers ----------
def norm_date(s: str) -> str:
    for fmt in ("%d/%m/%Y","%d-%m-%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).strftime("%d/%m/%Y")
        except: pass
    return s.strip()

def clean_money(s: str) -> Optional[float]:
    if not s: return None
    s = s.replace("\xa0"," ").replace("€","").strip()
    s = s.replace(".", "").replace(",", ".")
    m = re.search(r"(\d+(?:\.\d{2})?)", s)
    if not m: return None
    try:
        return float(m.group(1))
    except:
        return None

def to_int(s: str) -> Optional[int]:
    if not s: return None
    m = re.search(r"\b(\d{1,5})\b", s)
    if not m: return None
    try:
        return int(m.group(1))
    except:
        return None

# Detecta un MODELO formato proveedor (p.ej., 3RC240/NARANJA)
MODELO_RE = re.compile(r"\b([A-Z0-9]{2,}[A-Z0-9/_\-]{2,})/[A-ZÁÉÍÓÚÑ0-9\-]{2,}\b")
# Detecta un PATRON formato cliente (p.ej., 0863769/66)
PATRON_RE = re.compile(r"\b(\d{5,9}/\d{1,3})\b")

HEADER_PEDIDO_RE = re.compile(r"(?:N[ºo]\s*doc|N[ºo]\s*Pedido|Pedido)\s*:\s*([A-Z]?\d[\w\-./]*)", re.IGNORECASE)
HEADER_FECHA_RE = re.compile(r"Fecha\s*:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
HEADER_FECHA_ENTREGA_RE = re.compile(r"Fecha\s*Entrega\s*:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
HEADER_DESTINO_RE = re.compile(r"(?:Destino|Destinatario|Lug\.?Entreg\.)\s*:\s*(.+)", re.IGNORECASE)

def parse_pdf(pdf_path: str) -> Tuple[Header, List[Linea]]:
    header = Header()
    lineas: List[Linea] = []
    with pdfplumber.open(pdf_path) as pdf:
        current_pedido = ""
        current_fecha = ""
        current_fecha_entrega = ""
        current_destino = ""
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1.5, y_tolerance=1.5) or ""

            mp = HEADER_PEDIDO_RE.search(text)
            if mp: current_pedido = mp.group(1).strip()
            mf = HEADER_FECHA_RE.search(text)
            if mf: current_fecha = mf.group(1).strip()
            mfe = HEADER_FECHA_ENTREGA_RE.search(text)
            if mfe: current_fecha_entrega = mfe.group(1).strip()
            md = HEADER_DESTINO_RE.search(text)
            if md:
                current_destino = md.group(1).splitlines()[0].strip()

            # Guarda en header (nos quedamos con el último visto en el pedido)
            header.pedido = header.pedido or current_pedido
            header.fecha = header.fecha or current_fecha
            header.fecha_entrega = header.fecha_entrega or current_fecha_entrega
            header.destino = header.destino or current_destino

            # 1) INTENTO TABLAS
            got_rows = 0
            tables = page.extract_tables(table_settings={
                "vertical_strategy":"text","horizontal_strategy":"text",
                "text_x_tolerance": 2, "text_y_tolerance": 2
            })
            for tb in tables or []:
                if not tb or len(tb) < 2: 
                    continue
                header_row = [ (c or "").strip().lower() for c in tb[0] ]
                if not any(("ref" in h) or ("ean" in h) or ("descripción" in h) or ("cantidad" in h) for h in header_row):
                    continue
                for row in tb[1:]:
                    cells = [ (c or "").strip() for c in row ]
                    if not any(cells): 
                        continue
                    # MODELO/PATRON
                    modelo = ""
                    patron = ""
                    for c in cells:
                        if not modelo:
                            m = MODELO_RE.search(c); 
                            if m: modelo = m.group(1)
                        if not patron:
                            p = PATRON_RE.search(c)
                            if p: patron = p.group(1)
                    # UNIDADES
                    qty = None
                    # Heurística: toma el primer entero “limpio” que no parezca CP o año
                    for c in cells:
                        cc = to_int(c)
                        if cc is not None and 0 < cc < 10000:
                            qty = cc; break
                    # PRECIO (neto unitario): último número con decimales antes de “EUR” en la fila
                    price = None
                    for c in reversed(cells):
                        mm = re.search(r"(\d{1,3}(?:[.,]\d{2}))\s*(?:EUR|€)", c)
                        if mm:
                            price = clean_money(mm.group(1))
                            break
                    if any([modelo, patron, qty, price]):
                        lineas.append(Linea(
                            modelo=modelo.split("/")[0],
                            patron=patron.split("/")[0] if patron else "",
                            precio=price, unidades=qty, ean="", page=i
                        ))
                        got_rows += 1

            # 2) FALLBACK REGEX (si no detectó tabla o se quedó corto)
            if got_rows == 0:
                # Buscamos EAN-13 y contexto
                for m in re.finditer(r"\b(\d{13})\b", text):
                    start = max(0, m.start()-240)
                    end = min(len(text), m.end()+240)
                    ctx = text[start:end]
                    # Descripción
                    desc = ""
                    md = re.search(r"Descripción:\s*([^\n\r]+)", ctx, re.IGNORECASE)
                    if md: desc = md.group(1).strip()
                    # MODELO y PATRON
                    modelo = ""
                    patron = ""
                    mm = MODELO_RE.search(ctx)
                    if mm: modelo = mm.group(1)
                    pp = PATRON_RE.search(ctx)
                    if pp: patron = pp.group(1)
                    # UNIDADES (suelen ir justo antes del EAN en Eurofiel)
                    # patrón típico: "... Descripción ...  3  8447571xxxxxxxx"
                    qty = None
                    mq = re.search(r"Descripción:[^\n\r]*?\b(\d{1,4})\b\s+\d{13}", ctx, re.IGNORECASE)
                    if mq:
                        qty = int(mq.group(1))
                    else:
                        # plan B: primer entero pequeño en el contexto
                        for n in re.findall(r"\b(\d{1,4})\b", ctx):
                            v = int(n)
                            if 0 < v < 10000:
                                qty = v; break
                    # PRECIO (buscamos neto unitario cercano a EUR en el contexto)
                    price = None
                    mp = re.findall(r"(\d{1,3}(?:[.,]\d{2}))\s*(?:EUR|€)", ctx)
                    if mp:
                        price = clean_money(mp[-1])  # el último suele ser neto unitario
                    if any([modelo, patron, qty, price, desc]):
                        lineas.append(Linea(
                            modelo=modelo.split("/")[0] if modelo else "",
                            patron=patron.split("/")[0] if patron else "",
                            precio=price, unidades=qty, ean=m.group(1), page=i, desc=desc
                        ))

    # Completa cabecera con últimas vistas
    header.pedido = header.pedido or ""
    header.fecha = norm_date(header.fecha) if header.fecha else ""
    header.fecha_entrega = norm_date(header.fecha_entrega) if header.fecha_entrega else ""
    header.destino = header.destino or ""
    return header, lineas

def build_dataframe(header: Header, lineas: List[Linea], eq: Dict[str, Dict[str,str]]) -> pd.DataFrame:
    rows = []
    tipo_norm = apply_eq(eq, "TIPO", header.tipo) or "Pedido"
    destino_norm = apply_eq(eq, "DESTINO", header.destino)

    for ln in lineas:
        rows.append({
            "TIPO": tipo_norm,
            "PEDIDO": header.pedido,
            "FECHA": header.fecha,
            "FECHA_ENTREGA": header.fecha_entrega,
            "DESTINO": destino_norm,
            "MODELO": apply_eq(eq, "MODELO", ln.modelo),
            "PATRON": apply_eq(eq, "PATRON", ln.patron),
            "UNIDADES": ln.unidades if ln.unidades is not None else "",
            "PRECIO": ln.precio if ln.precio is not None else "",
            # se rellenan luego:
            "UNIDADES_TOTALES": "",
            "IMPORTE_TOTAL": ""
        })

    df = pd.DataFrame(rows, columns=TARGET_COLS)
    # Totales por pedido
    if not df.empty:
        df["__u__"] = pd.to_numeric(df["UNIDADES"], errors="coerce").fillna(0)
        df["__p__"] = pd.to_numeric(df["PRECIO"], errors="coerce").fillna(0.0)
        df["__imp_linea__"] = df["__u__"] * df["__p__"]
        g = df.groupby("PEDIDO", dropna=False).agg(
            UNIDADES_TOTALES=("__u__", "sum"),
            IMPORTE_TOTAL=("__imp_linea__", "sum"),
        ).reset_index()
        df = df.merge(g, on="PEDIDO", how="left")
        df.drop(columns=["__u__","__p__","__imp_linea__"], inplace=True)
        # Reordena columnas por si se descolocaron:
        df = df[TARGET_COLS]
    return df

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--map", default=None)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    eq = parse_equivalences(args.map) if args.map else {}
    header, lineas = parse_pdf(args.pdf)
    df = build_dataframe(header, lineas, eq)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(out.suffix + ".tmp")

    if out.suffix.lower() in (".xlsx",".xls"):
        df.to_excel(tmp, index=False)
    else:
        df.to_csv(tmp, index=False)
    shutil.move(str(tmp), str(out))
    print(f"OK: {len(df)} líneas → {out}")

if __name__ == "__main__":
    main()
