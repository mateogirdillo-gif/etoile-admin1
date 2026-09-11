import os
import shutil
import threading
from datetime import datetime
import re
import openpyxl
from openpyxl.utils import column_index_from_string

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.environ.get(
    "ETOILE_EXCEL_PATH", os.path.join(BASE_DIR, "data", "Sistema_Etoile_2.xlsx")
)
BACKUP_DIR = os.path.join(BASE_DIR, "data", "backups")
_lock = threading.Lock()

INV_HEADER_ROW = 3
INV_FIRST_DATA_ROW = 4
INV_COLS = {"CODIGO": 1, "PRENDA": 2, "COLOR": 3, "TALLA": 4, "PRECIO": 5}

HIST_HEADER_ROW = 4
HIST_FIRST_DATA_ROW = 5
HIST_LAST_DATA_ROW = 5000 # <-- Aumentado para que no se llene nunca
HIST_COLS = {
    "FECHA": 1, "CODIGO": 2, "PRENDA": 3, "COLOR": 4, "TALLA": 5,
    "CANT": 6, "PRECIO": 7, "SUBTOTAL": 8, "CLIENTE": 9,
    "PEDIDO_ID": 10, "TELEFONO": 11, "METODO_PAGO": 12, "NUM_TRANSACCION": 13,
}

_REF_SIMPLE = re.compile(r"^=([A-Z]+)(\d+)$")

def _ensure_excel_exists():
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "INVENTARIO"
        ws1.cell(row=3, column=1, value="CODIGO")
        ws1.cell(row=3, column=2, value="PRENDA")
        ws1.cell(row=3, column=3, value="COLOR")
        ws1.cell(row=3, column=4, value="TALLA")
        ws1.cell(row=3, column=5, value="PRECIO")
        ws2 = wb.create_sheet("HISTORIAL")
        headers = ["FECHA","CODIGO","PRENDA","COLOR","TALLA","CANT","PRECIO","SUBTOTAL","CLIENTE","PEDIDO_ID","TELEFONO","METODO_PAGO","NUM_TRANSACCION"]
        for i, h in enumerate(headers, 1):
            ws2.cell(row=4, column=i, value=h)
        wb.save(EXCEL_PATH)
        wb.close()

def _backup():
    # en Render no hacemos backup si no hay disco, para no llenarlo
    if os.environ.get("RENDER"): return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"Sistema_Etoile_2_{stamp}.xlsx")
    try:
        shutil.copy2(EXCEL_PATH, dest)
    except: pass

def _load(data_only=False):
    _ensure_excel_exists()
    return openpyxl.load_workbook(EXCEL_PATH, data_only=data_only)

def _resolver_celda(ws, row, col, visitados=None):
    if visitados is None: visitados = set()
    val = ws.cell(row=row, column=col).value
    if not (isinstance(val, str) and val.startswith("=")): return val
    m = _REF_SIMPLE.match(val)
    if not m: return None
    ref_col = column_index_from_string(m.group(1))
    ref_row = int(m.group(2))
    if (ref_row, ref_col) in visitados: return None
    visitados.add((ref_row, ref_col))
    return _resolver_celda(ws, ref_row, ref_col, visitados)

def get_inventario_raw(wb=None):
    own = wb is None
    wb = wb or _load(data_only=True)
    if "INVENTARIO" not in wb.sheetnames:
        if own: wb.close()
        return []
    ws = wb["INVENTARIO"]
    items = []
    for row in range(INV_FIRST_DATA_ROW, ws.max_row + 1):
        codigo = ws.cell(row=row, column=INV_COLS["CODIGO"]).value
        if not codigo: continue
        items.append({
            "codigo": str(codigo).strip(),
            "prenda": _resolver_celda(ws, row, INV_COLS["PRENDA"]),
            "color": _resolver_celda(ws, row, INV_COLS["COLOR"]),
            "talla": _resolver_celda(ws, row, INV_COLS["TALLA"]),
            "precio": ws.cell(row=row, column=INV_COLS["PRECIO"]).value or 0,
            "_row": row,
        })
    if own: wb.close()
    return items

def get_historial_raw(wb=None):
    own = wb is None
    wb = wb or _load(data_only=True)
    if "HISTORIAL" not in wb.sheetnames:
        if own: wb.close()
        return []
    ws = wb["HISTORIAL"]
    items = []
    for row in range(HIST_FIRST_DATA_ROW, ws.max_row + 1):
        codigo = ws.cell(row=row, column=HIST_COLS["CODIGO"]).value
        if not codigo: continue
        if isinstance(codigo, str) and codigo.startswith("="): continue
        items.append({
            "fecha": ws.cell(row=row, column=HIST_COLS["FECHA"]).value,
            "codigo": str(codigo).strip(),
            "prenda": ws.cell(row=row, column=HIST_COLS["PRENDA"]).value,
            "color": ws.cell(row=row, column=HIST_COLS["COLOR"]).value,
            "talla": ws.cell(row=row, column=HIST_COLS["TALLA"]).value,
            "cant": ws.cell(row=row, column=HIST_COLS["CANT"]).value or 0,
            "precio": ws.cell(row=row, column=HIST_COLS["PRECIO"]).value or 0,
            "subtotal": ws.cell(row=row, column=HIST_COLS["SUBTOTAL"]).value or 0,
            "cliente": ws.cell(row=row, column=HIST_COLS["CLIENTE"]).value,
            "pedido_id": ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"]).value,
            "telefono": ws.cell(row=row, column=HIST_COLS["TELEFONO"]).value,
            "metodo_pago": ws.cell(row=row, column=HIST_COLS["METODO_PAGO"]).value,
            "num_transaccion": ws.cell(row=row, column=HIST_COLS["NUM_TRANSACCION"]).value,
        })
    if own: wb.close()
    return items

# --- FUNCIONES QUE FALTABAN PARA QUE NO CRASHEE EN LA NUBE ---
def actualizar_stock(codigo, stock):
    # Ya no usas stock, pero tu API lo sigue llamando. Lo dejamos como no-op para que no crashee
    return True

def get_inventario_con_disponible():
    with _lock:
        wb = _load(data_only=True)
        inventario = get_inventario_raw(wb)
        historial = get_historial_raw(wb)
        wb.close()
    vendido_por_codigo = {}
    for h in historial:
        vendido_por_codigo[h["codigo"]] = vendido_por_codigo.get(h["codigo"], 0) + (h["cant"] or 0)
    for item in inventario:
        item["vendido"] = vendido_por_codigo.get(item["codigo"], 0)
    return inventario

def get_prendas_agrupadas(busqueda=None):
    inventario = get_inventario_con_disponible()
    if busqueda:
        b = busqueda.strip().lower()
        inventario = [i for i in inventario if b in (i["prenda"] or "").lower()]
    agrupado = {}
    for item in inventario:
        nombre = item["prenda"] or "(sin nombre)"
        agrupado.setdefault(nombre, []).append(item)
    resultado = []
    for nombre, variantes in sorted(agrupado.items()):
        colores = sorted(set(v["color"] for v in variantes if v["color"]))
        tallas = sorted(set(v["talla"] for v in variantes if v["talla"]))
        precio = variantes[0]["precio"] if variantes else 0
        resultado.append({"prenda": nombre, "colores": colores, "tallas": tallas, "precio": precio, "variantes": variantes})
    return resultado

def _abreviar(texto, largo):
    limpio = "".join(ch for ch in (texto or "") if ch.isalnum())
    return limpio[:largo].upper() if limpio else "XX"

def sugerir_codigo(prenda, color, talla, wb=None):
    palabras = [p for p in (prenda or "").split() if p]
    if len(palabras) >= 2: prendaAbrev = _abreviar(palabras[0], 2) + _abreviar(palabras[1], 2)
    elif len(palabras) == 1: prendaAbrev = _abreviar(palabras[0], 4)
    else: prendaAbrev = "XXXX"
    colorAbrev = _abreviar(color, 2)
    tallaTxt = "".join(ch for ch in (talla or "").strip().upper() if ch.isalnum()) or "XX"
    base = f"{prendaAbrev}-{colorAbrev}-{tallaTxt}"
    own = wb is None
    wb = wb or _load()
    inventario = get_inventario_raw(wb)
    if own: wb.close()
    existentes = {i["codigo"] for i in inventario}
    if base not in existentes: return base
    n = 2
    while f"{base}-{n}" in existentes: n += 1
    return f"{base}-{n}"

def get_prendas_nombres():
    inventario = get_inventario_raw()
    return sorted(set(i["prenda"] for i in inventario if i["prenda"] and not str(i["prenda"]).startswith("=")))

def agregar_prenda(prenda, color, talla, precio, codigo=None):
    prenda = (prenda or "").strip()
    color = (color or "").strip()
    talla = (talla or "").strip()
    if not prenda: raise ValueError("Falta el nombre de la prenda")
    if precio is None or float(precio) < 0: raise ValueError("Precio inválido")
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        inventario = get_inventario_raw(wb)
        existentes = {i["codigo"] for i in inventario}
        if codigo:
            codigo = codigo.strip().upper()
            if codigo in existentes:
                wb.close()
                raise ValueError(f"El código {codigo} ya existe.")
        else:
            codigo = sugerir_codigo(prenda, color, talla, wb)
        fila = INV_FIRST_DATA_ROW
        for row in range(INV_FIRST_DATA_ROW, ws.max_row + 2):
            if not ws.cell(row=row, column=INV_COLS["CODIGO"]).value:
                fila = row
                break
        else: fila = ws.max_row + 1
        ws.cell(row=fila, column=INV_COLS["CODIGO"], value=codigo)
        ws.cell(row=fila, column=INV_COLS["PRENDA"], value=prenda)
        ws.cell(row=fila, column=INV_COLS["COLOR"], value=color)
        ws.cell(row=fila, column=INV_COLS["TALLA"], value=talla)
        ws.cell(row=fila, column=INV_COLS["PRECIO"], value=float(precio))
        _backup()
        wb.save(EXCEL_PATH)
        wb.close()
    return codigo

def siguiente_pedido_id(wb):
    ws = wb["HISTORIAL"]
    max_id = 0
    for row in range(HIST_FIRST_DATA_ROW, ws.max_row + 1):
        pid = ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"]).value
        if pid:
            try:
                num = int(str(pid).replace("PED-", ""))
                max_id = max(max_id, num)
            except: continue
    return f"PED-{max_id + 1:04d}"

def crear_pedido(cliente, telefono, items, metodo_pago, num_transaccion):
    if not items: raise ValueError("El pedido no tiene prendas")
    if metodo_pago not in ("efectivo", "transferencia"): raise ValueError("Método de pago inválido")
    if metodo_pago == "transferencia" and not num_transaccion: raise ValueError("Falta el número de transacción")
    with _lock:
        wb = _load()
        ws_hist = wb["HISTORIAL"]
        pedido_id = siguiente_pedido_id(wb)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
        fila = None
        for row in range(HIST_FIRST_DATA_ROW, ws_hist.max_row + 2):
            valor = ws_hist.cell(row=row, column=HIST_COLS["CODIGO"]).value
            if not valor:
                fila = row
                break
        if fila is None: fila = ws_hist.max_row + 1
        for it in items:
            subtotal = round(it["precio"] * it["cantidad"], 2)
            ws_hist.cell(row=fila, column=HIST_COLS["FECHA"], value=fecha)
            ws_hist.cell(row=fila, column=HIST_COLS["CODIGO"], value=it["codigo"])
            ws_hist.cell(row=fila, column=HIST_COLS["PRENDA"], value=it.get("prenda"))
            ws_hist.cell(row=fila, column=HIST_COLS["COLOR"], value=it.get("color"))
            ws_hist.cell(row=fila, column=HIST_COLS["TALLA"], value=it.get("talla"))
            ws_hist.cell(row=fila, column=HIST_COLS["CANT"], value=it["cantidad"])
            ws_hist.cell(row=fila, column=HIST_COLS["PRECIO"], value=it["precio"])
            ws_hist.cell(row=fila, column=HIST_COLS["SUBTOTAL"], value=subtotal)
            ws_hist.cell(row=fila, column=HIST_COLS["CLIENTE"], value=cliente)
            ws_hist.cell(row=fila, column=HIST_COLS["PEDIDO_ID"], value=pedido_id)
            ws_hist.cell(row=fila, column=HIST_COLS["TELEFONO"], value=telefono)
            ws_hist.cell(row=fila, column=HIST_COLS["METODO_PAGO"], value=metodo_pago)
            ws_hist.cell(row=fila, column=HIST_COLS["NUM_TRANSACCION"], value=num_transaccion if metodo_pago == "transferencia" else "EFECTIVO")
            fila += 1
        _backup()
        wb.save(EXCEL_PATH)
        wb.close()
    return pedido_id

def get_pedidos_agrupados(desde=None, hasta=None, cliente=None, codigo=None):
    historial = get_historial_raw()
    def pasa_filtros(h):
        if desde and (not h["fecha"] or str(h["fecha"])[:10] < desde): return False
        if hasta and (not h["fecha"] or str(h["fecha"])[:10] > hasta): return False
        if cliente and cliente.lower() not in (h["cliente"] or "").lower(): return False
        if codigo and codigo.lower() not in (h["codigo"] or "").lower(): return False
        return True
    historial = [h for h in historial if pasa_filtros(h)]
    pedidos = {}
    for h in historial:
        pid = h["pedido_id"] or f"SIN-ID-{h['fecha']}-{h['cliente']}"
        if pid not in pedidos:
            pedidos[pid] = {"pedido_id": pid, "fecha": h["fecha"], "cliente": h["cliente"], "telefono": h["telefono"], "metodo_pago": h["metodo_pago"], "num_transaccion": h["num_transaccion"], "items": [], "total": 0}
        pedidos[pid]["items"].append(h)
        pedidos[pid]["total"] += h["subtotal"] or 0
    resultado = list(pedidos.values())
    resultado.sort(key=lambda p: p["fecha"] or "", reverse=True)
    return resultado

def eliminar_pedido(pedido_id):
        if not pedido_id:
            raise ValueError("Falta el ID del pedido")
            with _lock:
                wb=_load()
                raise ValueError("No existe la hoja HISTORIAL")
                ws = wb["HISTORIAL"]
                filas_a_borrar=[]
        for row in range(HIST_FIRST_DATA_ROW, ws.max_row + 1):
            pid=ws.cell(row=row, column_HIST_COLS["PEDIDO_ID"]).value
            if pid and str(pid).strip()==str(pedido_id).strip()
            filas_a_borrar.apppend(row)
        if not filas_a_borrar:
            wb.close()
        raise ValueError(f"No se encontro el pedido {pedido_id}")
        for row in reverse(filas_a_borrar):
            ws.delete_rows(row, 1)
            _backup()
            wb.sav(EXCEL_PATH)
            WB.close()
    return True