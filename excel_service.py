import shutil, threading, re
from datetime import datetime, date
from pathlib import Path
import openpyxl

BASE_DIR = Path(__file__).resolve().parent
# busca tu archivo con cualquier nombre
POSIBLES = [
    BASE_DIR / "data" / "Sistema_Etoile_2.xlsx",
    BASE_DIR / "data" / "Sistema_Etoile_2.xls",
    BASE_DIR / "Sistema_Etoile_2.xlsx",
]
EXCEL_PATH = next((p for p in POSIBLES if p.exists()), BASE_DIR / "data" / "Sistema_Etoile_2.xlsx")

INV_COLS = {"CODIGO":1,"PRENDA":2,"COLOR":3,"TALLA":4,"PRECIO":5}
HIST_COLS = {"PEDIDO_ID":1,"FECHA":2,"CLIENTE":3,"TELEFONO":4,"CODIGO":5,"PRENDA":6,"COLOR":7,"TALLA":8,"CANT":9,"PRECIO":10,"SUBTOTAL":11,"METODO_PAGO":12,"NUM_TRANSACCION":13}
_lock = threading.Lock()

def _load(): return openpyxl.load_workbook(EXCEL_PATH)

def get_prendas_agrupadas(q=None):
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        grupos = {}
        for r in range(1, ws.max_row+1):
            codigo = ws.cell(r, INV_COLS["CODIGO"]).value
            prenda = ws.cell(r, INV_COLS["PRENDA"]).value
            if not codigo or not prenda: continue
            # ESTO ARREGLA TU FOTO - ignora el encabezado
            if str(codigo).strip().upper() == "CODIGO": continue
            if str(prenda).strip().upper() == "PRENDA": continue

            if q and q.lower() not in str(prenda).lower() and q.lower() not in str(codigo).lower():
                continue

            key = str(prenda).strip()
            if key not in grupos: grupos[key] = {"prenda": key, "variantes": []}

            precio_raw = ws.cell(r, INV_COLS["PRECIO"]).value
            try: precio = float(precio_raw)
            except:
                if str(precio_raw).upper() == "PRECIO": continue
                precio = 0

            grupos[key]["variantes"].append({
                "codigo": str(codigo).strip(),
                "color": ws.cell(r, INV_COLS["COLOR"]).value,
                "talla": ws.cell(r, INV_COLS["TALLA"]).value,
                "precio": precio,
                "stock": 999,
                "disponible": 999,
            })
        wb.close()
        return list(grupos.values())

def get_inventario_con_disponible():
    # para compatibilidad
    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        out=[]
        for r in range(1, ws.max_row+1):
            codigo = ws.cell(r, INV_COLS["CODIGO"]).value
            prenda = ws.cell(r, INV_COLS["PRENDA"]).value
            if not codigo or not prenda: continue
            if str(codigo).strip().upper() == "CODIGO": continue
            precio_raw = ws.cell(r, INV_COLS["PRECIO"]).value
            try: precio = float(precio_raw)
            except: continue
            out.append({"codigo": str(codigo).strip(),"prenda": prenda,"color": ws.cell(r, 3).value,"talla": ws.cell(r, 4).value,"precio": precio,"stock":999,"disponible":999})
        wb.close()
        return out

# resto funciones iguales pero sin stock
def get_prendas_nombres():
    with _lock:
        wb=_load(); ws=wb["INVENTARIO"]; nombres=set()
        for r in range(1, ws.max_row+1):
            v=ws.cell(r, 2).value; c=ws.cell(r, 1).value
            if v and c and str(c).upper()!="CODIGO" and str(v).upper()!="PRENDA": nombres.add(str(v).strip())
        wb.close(); return sorted(nombres)

def sugerir_codigo(p,c,t): return f"{p[:4].upper()}-{c[:2].upper()}-{t}" if p else "PREN-XX-U"
def agregar_prenda(prenda,color,talla,precio,codigo=None):
    with _lock:
        wb=_load(); ws=wb["INVENTARIO"]
        if not codigo: codigo=sugerir_codigo(prenda,color,talla)
        nr=ws.max_row+1
        ws.cell(nr,1).value=codigo; ws.cell(nr,2).value=prenda; ws.cell(nr,3).value=color; ws.cell(nr,4).value=talla; ws.cell(nr,5).value=float(precio); ws.cell(nr,6).value=999
        wb.save(EXCEL_PATH); wb.close(); return codigo

def _siguiente_pedido_id(ws):
    max_n=0
    for r in range(1, ws.max_row+1):
        pid=ws.cell(r,1).value
        if pid:
            m=re.search(r"(\d+)", str(pid))
            if m: max_n=max(max_n,int(m.group(1)))
    return f"PED-{max_n+1:04d}"

def crear_pedido(cliente,telefono,items,metodo_pago,num_transaccion):
    with _lock:
        wb=_load(); ws_inv=wb["INVENTARIO"]; ws_hist=wb["HISTORIAL"] if "HISTORIAL" in wb.sheetnames else wb.create_sheet("HISTORIAL")
        if ws_hist.max_row==1 and not ws_hist.cell(1,1).value:
            headers=["PEDIDO_ID","FECHA","CLIENTE","TELEFONO","CODIGO","PRENDA","COLOR","TALLA","CANT","PRECIO","SUBTOTAL","METODO_PAGO","NUM_TRANSACCION"]
            for i,h in enumerate(headers,1): ws_hist.cell(1,i).value=h
        pedido_id=_siguiente_pedido_id(ws_hist); fecha=datetime.now().strftime("%Y-%m-%d %H:%M")
        for it in items:
            codigo=it.get("codigo"); cant=int(it.get("cant",1)); precio=0; prenda=color=talla=""
            for r in range(1, ws_inv.max_row+1):
                if str(ws_inv.cell(r,1).value).strip()==str(codigo).strip():
                    prenda=ws_inv.cell(r,2).value; color=ws_inv.cell(r,3).value; talla=ws_inv.cell(r,4).value; precio=ws_inv.cell(r,5).value or 0; break
            subtotal=float(precio)*cant; nr=ws_hist.max_row+1
            ws_hist.cell(nr,1).value=pedido_id; ws_hist.cell(nr,2).value=fecha; ws_hist.cell(nr,3).value=cliente; ws_hist.cell(nr,4).value=telefono; ws_hist.cell(nr,5).value=codigo; ws_hist.cell(nr,6).value=prenda; ws_hist.cell(nr,7).value=color; ws_hist.cell(nr,8).value=talla; ws_hist.cell(nr,9).value=cant; ws_hist.cell(nr,10).value=precio; ws_hist.cell(nr,11).value=subtotal; ws_hist.cell(nr,12).value=metodo_pago; ws_hist.cell(nr,13).value=num_transaccion
        wb.save(EXCEL_PATH); wb.close(); return pedido_id

def get_pedidos_agrupados(desde=None, hasta=None, cliente=None, codigo=None):
    with _lock:
        wb=_load()
        if "HISTORIAL" not in wb.sheetnames: wb.close(); return []
        ws=wb["HISTORIAL"]; pedidos={}
        for r in range(2, ws.max_row+1):
            pid=ws.cell(r,1).value
            if not pid or str(pid).upper()=="PEDIDO_ID": continue
            cli=ws.cell(r,3).value or ""; cod=ws.cell(r,5).value or ""
            if cliente and cliente.lower() not in str(cli).lower(): continue
            if codigo and codigo.lower() not in str(cod).lower(): continue
            if pid not in pedidos:
                pedidos[pid]={"pedido_id":pid,"fecha":str(ws.cell(r,2).value or ""),"cliente":cli,"telefono":ws.cell(r,4).value,"metodo_pago":ws.cell(r,12).value,"num_transaccion":ws.cell(r,13).value,"items":[],"total":0}
            sub=ws.cell(r,11).value or 0
            pedidos[pid]["items"].append({"codigo":cod,"prenda":ws.cell(r,6).value,"color":ws.cell(r,7).value,"talla":ws.cell(r,8).value,"cant":ws.cell(r,9).value,"precio":ws.cell(r,10).value,"subtotal":sub})
            pedidos[pid]["total"]+=float(sub or 0)
        wb.close()
        return sorted(pedidos.values(), key=lambda x: x["pedido_id"], reverse=True)

def eliminar_pedido(pedido_id):
    with _lock:
        wb=_load(); ws=wb["HISTORIAL"]; filas=[]
        for row in range(1, ws.max_row+1):
            pid=ws.cell(row,1).value
            if pid and str(pid).strip()==str(pedido_id).strip(): filas.append(row)
        if not filas: wb.close(); raise ValueError("No existe")
        for row in reversed(filas): ws.delete_rows(row,1)
        wb.save(EXCEL_PATH); wb.close()
    return True
def actualizar_stock(codigo,nuevo_stock): return True
