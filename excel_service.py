import os, re, shutil, threading
from datetime import datetime, date
from pathlib import Path
import openpyxl

BASE_DIR = Path(__file__).resolve().parent
BASE_DIR = Path(__file__).resolve().parent
POSIBLES = [
    BASE_DIR / "data" / "inventario.xlsx",
    BASE_DIR / "data" / "Sistema_Etoile_2.xlsx",
    BASE_DIR / "data" / "Sistema_Etoile_2.xls",
    BASE_DIR / "Sistema_Etoile_2.xls",
    BASE_DIR / "Sistema_Etoile_2.xlsx",
]
EXCEL_PATH = next((p for p in POSIBLES if p.exists()), BASE_DIR / "data" / "inventario.xlsx")
BACKUP_DIR = BASE_DIR / "data" / "backups"

INV_COLS = {"CODIGO":1,"PRENDA":2,"COLOR":3,"TALLA":4,"PRECIO":5,"STOCK":6,"RESERVADO":7,"DISPONIBLE":8}
HIST_COLS = {"PEDIDO_ID":1,"FECHA":2,"CLIENTE":3,"TELEFONO":4,"CODIGO":5,"PRENDA":6,"COLOR":7,"TALLA":8,"CANT":9,"PRECIO":10,"SUBTOTAL":11,"METODO_PAGO":12,"NUM_TRANSACCION":13}

INV_FIRST_DATA_ROW = 2
HIST_FIRST_DATA_ROW = 2

_lock = threading.Lock()

def _load():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontró {EXCEL_PATH}")
    return openpyxl.load_workbook(EXCEL_PATH)

def _backup():
    try:
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy(EXCEL_PATH, BACKUP_DIR / f"inventario_{ts}.xlsx")
    except Exception:
        pass

def _parse_date(v):
    if not v: return None
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except: pass
    return None

def _norm(s):
    return re.sub(r"\s+"," ", str(s or "").strip()).lower()

def get_prendas_agrupadas(q=None):
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        grupos = {}
        for r in range(INV_FIRST_DATA_ROW, ws.max_row+1):
            prenda = ws.cell(r, INV_COLS["PRENDA"]).value
            if not prenda: continue
            if q and q.lower() not in str(prenda).lower():
                # busca también en color/codigo
                color = str(ws.cell(r, INV_COLS["COLOR"]).value or "")
                codigo = str(ws.cell(r, INV_COLS["CODIGO"]).value or "")
                if q.lower() not in color.lower() and q.lower() not in codigo.lower():
                    continue
            key = str(prenda).strip()
            if key not in grupos:
                grupos[key] = {"prenda": key, "variantes": []}
            grupos[key]["variantes"].append({
                "codigo": ws.cell(r, INV_COLS["CODIGO"]).value,
                "color": ws.cell(r, INV_COLS["COLOR"]).value,
                "talla": ws.cell(r, INV_COLS["TALLA"]).value,
                "precio": ws.cell(r, INV_COLS["PRECIO"]).value or 0,
                "stock": ws.cell(r, INV_COLS["STOCK"]).value or 0,
                "disponible": ws.cell(r, INV_COLS["DISPONIBLE"]).value or 0,
            })
        wb.close()
        return list(grupos.values())

def get_inventario_con_disponible():
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        out=[]
        for r in range(INV_FIRST_DATA_ROW, ws.max_row+1):
            codigo = ws.cell(r, INV_COLS["CODIGO"]).value
            if not codigo: continue
            out.append({
                "codigo": codigo,
                "prenda": ws.cell(r, INV_COLS["PRENDA"]).value,
                "color": ws.cell(r, INV_COLS["COLOR"]).value,
                "talla": ws.cell(r, INV_COLS["TALLA"]).value,
                "precio": ws.cell(r, INV_COLS["PRECIO"]).value or 0,
                "stock": ws.cell(r, INV_COLS["STOCK"]).value or 0,
                "disponible": ws.cell(r, INV_COLS["DISPONIBLE"]).value or 0,
            })
        wb.close()
        return out

def actualizar_stock(codigo, nuevo_stock):
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        for r in range(INV_FIRST_DATA_ROW, ws.max_row+1):
            if str(ws.cell(r, INV_COLS["CODIGO"]).value).strip() == str(codigo).strip():
                ws.cell(r, INV_COLS["STOCK"]).value = int(nuevo_stock)
                _backup()
                wb.save(EXCEL_PATH)
                wb.close()
                return True
        wb.close()
        raise ValueError("Código no encontrado")

def get_prendas_nombres():
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        nombres=set()
        for r in range(INV_FIRST_DATA_ROW, ws.max_row+1):
            v=ws.cell(r, INV_COLS["PRENDA"]).value
            if v: nombres.add(str(v).strip())
        wb.close()
        return sorted(nombres)

def sugerir_codigo(prenda, color, talla):
    p = (prenda[:4].upper() if prenda else "PREN")
    c = (color[:2].upper() if color else "XX")
    t = (str(talla).upper() if talla else "U")
    return f"{p}-{c}-{t}"

def agregar_prenda(prenda, color, talla, precio, codigo=None):
    if not prenda or not precio:
        raise ValueError("Falta prenda o precio")
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        if not codigo:
            codigo = sugerir_codigo(prenda,color,talla)
        # verificar duplicado
        for r in range(INV_FIRST_DATA_ROW, ws.max_row+1):
            if str(ws.cell(r, INV_COLS["CODIGO"]).value).strip() == codigo.strip():
                wb.close()
                raise ValueError(f"El código {codigo} ya existe")
        nr = ws.max_row+1
        ws.cell(nr, INV_COLS["CODIGO"]).value = codigo
        ws.cell(nr, INV_COLS["PRENDA"]).value = prenda
        ws.cell(nr, INV_COLS["COLOR"]).value = color
        ws.cell(nr, INV_COLS["TALLA"]).value = talla
        ws.cell(nr, INV_COLS["PRECIO"]).value = float(precio)
        ws.cell(nr, INV_COLS["STOCK"]).value = 0
        ws.cell(nr, INV_COLS["RESERVADO"]).value = 0
        ws.cell(nr, INV_COLS["DISPONIBLE"]).value = 0
        _backup()
        wb.save(EXCEL_PATH)
        wb.close()
        return codigo

def _siguiente_pedido_id(ws):
    max_n=0
    for r in range(HIST_FIRST_DATA_ROW, ws.max_row+1):
        pid = ws.cell(r, HIST_COLS["PEDIDO_ID"]).value
        if pid:
            m=re.search(r"(\d+)", str(pid))
            if m:
                max_n=max(max_n, int(m.group(1)))
    return f"PED-{max_n+1:04d}"

def crear_pedido(cliente, telefono, items, metodo_pago, num_transaccion):
    with _lock:
        wb = _load()
        ws_inv = wb["INVENTARIO"]
        ws_hist = wb["HISTORIAL"] if "HISTORIAL" in wb.sheetnames else wb.create_sheet("HISTORIAL")
        pedido_id = _siguiente_pedido_id(ws_hist)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
        for it in items:
            codigo = it.get("codigo")
            cant = int(it.get("cant",1))
            # buscar precio en inventario
            precio=0
            prenda=color=talla=""
            for r in range(INV_FIRST_DATA_ROW, ws_inv.max_row+1):
                if str(ws_inv.cell(r, INV_COLS["CODIGO"]).value).strip() == str(codigo).strip():
                    prenda = ws_inv.cell(r, INV_COLS["PRENDA"]).value
                    color = ws_inv.cell(r, INV_COLS["COLOR"]).value
                    talla = ws_inv.cell(r, INV_COLS["TALLA"]).value
                    precio = ws_inv.cell(r, INV_COLS["PRECIO"]).value or 0
                    break
            subtotal = float(precio)*cant
            nr = ws_hist.max_row+1
            ws_hist.cell(nr, HIST_COLS["PEDIDO_ID"]).value = pedido_id
            ws_hist.cell(nr, HIST_COLS["FECHA"]).value = fecha
            ws_hist.cell(nr, HIST_COLS["CLIENTE"]).value = cliente
            ws_hist.cell(nr, HIST_COLS["TELEFONO"]).value = telefono
            ws_hist.cell(nr, HIST_COLS["CODIGO"]).value = codigo
            ws_hist.cell(nr, HIST_COLS["PRENDA"]).value = prenda
            ws_hist.cell(nr, HIST_COLS["COLOR"]).value = color
            ws_hist.cell(nr, HIST_COLS["TALLA"]).value = talla
            ws_hist.cell(nr, HIST_COLS["CANT"]).value = cant
            ws_hist.cell(nr, HIST_COLS["PRECIO"]).value = precio
            ws_hist.cell(nr, HIST_COLS["SUBTOTAL"]).value = subtotal
            ws_hist.cell(nr, HIST_COLS["METODO_PAGO"]).value = metodo_pago
            ws_hist.cell(nr, HIST_COLS["NUM_TRANSACCION"]).value = num_transaccion
        _backup()
        wb.save(EXCEL_PATH)
        wb.close()
        return pedido_id

def get_pedidos_agrupados(desde=None, hasta=None, cliente=None, codigo=None):
    with _lock:
        wb = _load()
        if "HISTORIAL" not in wb.sheetnames:
            wb.close()
            return []
        ws = wb["HISTORIAL"]
        d_desde = _parse_date(desde)
        d_hasta = _parse_date(hasta)
        pedidos={}
        for r in range(HIST_FIRST_DATA_ROW, ws.max_row+1):
            pid = ws.cell(r, HIST_COLS["PEDIDO_ID"]).value
            if not pid: continue
            cli = ws.cell(r, HIST_COLS["CLIENTE"]).value or ""
            cod = ws.cell(r, HIST_COLS["CODIGO"]).value or ""
            fecha_raw = ws.cell(r, HIST_COLS["FECHA"]).value
            f_date = _parse_date(fecha_raw)
            if cliente and cliente.lower() not in str(cli).lower(): continue
            if codigo and codigo.lower() not in str(cod).lower(): continue
            if d_desde and f_date and f_date < d_desde: continue
            if d_hasta and f_date and f_date > d_hasta: continue
            if pid not in pedidos:
                pedidos[pid] = {
                    "pedido_id": pid,
                    "fecha": str(fecha_raw or ""),
                    "cliente": cli,
                    "telefono": ws.cell(r, HIST_COLS["TELEFONO"]).value,
                    "metodo_pago": ws.cell(r, HIST_COLS["METODO_PAGO"]).value,
                    "num_transaccion": ws.cell(r, HIST_COLS["NUM_TRANSACCION"]).value,
                    "items":[],
                    "total":0
                }
            sub = ws.cell(r, HIST_COLS["SUBTOTAL"]).value or 0
            pedidos[pid]["items"].append({
                "codigo": cod,
                "prenda": ws.cell(r, HIST_COLS["PRENDA"]).value,
                "color": ws.cell(r, HIST_COLS["COLOR"]).value,
                "talla": ws.cell(r, HIST_COLS["TALLA"]).value,
                "cant": ws.cell(r, HIST_COLS["CANT"]).value,
                "precio": ws.cell(r, HIST_COLS["PRECIO"]).value,
                "subtotal": sub
            })
            pedidos[pid]["total"] += float(sub or 0)
        wb.close()
        return sorted(pedidos.values(), key=lambda x: x["pedido_id"], reverse=True)

def eliminar_pedido(pedido_id):
    if not pedido_id:
        raise ValueError("Falta el ID del pedido")
    with _lock:
        wb = _load()
        if "HISTORIAL" not in wb.sheetnames:
            wb.close()
            raise ValueError("No existe la hoja HISTORIAL")
        ws = wb["HISTORIAL"]
        filas_a_borrar = []
        for row in range(HIST_FIRST_DATA_ROW, ws.max_row + 1):
            pid = ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"]).value
            if pid and str(pid).strip() == str(pedido_id).strip():
                filas_a_borrar.append(row)
        if not filas_a_borrar:
            wb.close()
            raise ValueError(f"No se encontró el pedido {pedido_id}")
        for row in reversed(filas_a_borrar):
            ws.delete_rows(row, 1)
        _backup()
        wb.save(EXCEL_PATH)
        wb.close()
    return True
