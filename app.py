import os
import json
import sqlite3
import datetime
from typing import List, Dict, Any
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, HTTPException, Depends
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
DB_PATH = os.path.join(BASE_DIR, "facturacion.db")
os.makedirs(STATIC_DIR, exist_ok=True)

app = FastAPI(title="Control de Facturación - Multi-usuario en Tiempo Real")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Authentication Config
USERS_CONFIG = {
    "1143851865": {"name": "Sol", "role": "admin", "theme": "sol", "filter": "Sol"},
    "1107078586": {"name": "Edwar", "role": "admin", "theme": "edwar", "filter": "Edwar"},
    "1234": {"name": "Invitado", "role": "viewer", "theme": "guest", "filter": "ALL"}
}

# Initial Clients Definition (41 clients)
CLIENTS_MASTER = [
    {"id": 1, "name": "AIR FRANCE - BOGOTA - TRIPULACIONES", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Sol", "contact": "Portal API Global", "obs": "Se genera factura, se carga en API Global", "reqInf": False, "reqOc": False, "canal": "Portal API Global", "keyDay": "Primeros 5 días"},
    {"id": 2, "name": "ANDI", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Sol", "contact": "Marisol Marín", "obs": "Se envía informe a Marisol Marín, dada la respuesta se factura.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 3, "name": "API", "prog": "* MENSUAL", "freqType": "MENSUAL", "resp": "Sol", "contact": "API", "obs": "Se genera factura, primeros días del mes siguiente.", "reqInf": False, "reqOc": False, "canal": "Correo electrónico", "keyDay": "Días 1-5 mes siguiente"},
    {"id": 4, "name": "APTAR", "prog": "CADA 02-16", "freqType": "02-16", "resp": "Edwar", "contact": "Cristina Rojas (cc Erika Zapata)", "obs": "Se envía informe a Cristina Rojas con copia a Erika Zapata, a vuelta de correo envían orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 5, "name": "AVIANCA - BOGOTA", "prog": "CADA 05-23", "freqType": "05-23", "resp": "Sol", "contact": "Edison Torres, A. Peralta, L. Ramirez", "obs": "Se envía informe a Edison Torres - Alejandro Peralta - Luis Ramirez y se espera orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 05 y Día 23"},
    {"id": 6, "name": "AVIANCA - CALI", "prog": "CADA 05-23", "freqType": "05-23", "resp": "Sol", "contact": "Edison Torres, A. Peralta, L. Ramirez", "obs": "Se envía informe a Edison Torres - Alejandro Peralta - Luis Ramirez y se espera orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 05 y Día 23"},
    {"id": 7, "name": "AVIANCA - EQUIPAJES", "prog": "SEMANAL", "freqType": "SEMANAL", "resp": "Sol", "contact": "Avianca Equipajes", "obs": "Se realiza informes semanales junto los vouchers de transporte y al finalizar el mes, se debe generar un resumen en el informe.", "reqInf": True, "reqOc": True, "canal": "Correo + Vouchers", "keyDay": "Semanal + Cierre mes"},
    {"id": 8, "name": "BANCO DE LA REPUBLICA", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Don Hector", "obs": "Se genera factura de inmediato, don Hector facilita la información.", "reqInf": False, "reqOc": False, "canal": "Factura electrónica", "keyDay": "Inmediato tras servicio"},
    {"id": 9, "name": "BANCO DE OCCIDENTE", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Autorizador de viaje", "obs": "Revisar el número de viaje y la información de quién autoriza.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 10, "name": "BMS", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Mary Holguín, Paula Martinez y equipo", "obs": "Se realizan facturas, una por cada orden de compra. 1 solo informe con 1 hoja de Excel por OC.", "reqInf": True, "reqOc": True, "canal": "Correo con Excel x OC", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 11, "name": "CINEMARK COLOMBIA", "prog": "MES ANTICIPADO", "freqType": "ANTICIPADO", "resp": "Sol", "contact": "Alexander Victoria", "obs": "Se genera factura, mes anticipado y se envía factura a Alexander Victoria (más tardar a 5 del mes en curso)", "reqInf": False, "reqOc": False, "canal": "Correo electrónico", "keyDay": "Máximo día 05 mes curso"},
    {"id": 12, "name": "COLEGIO COLOMBO BRITANICO", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Administración CCB", "obs": "Se genera factura", "reqInf": False, "reqOc": False, "canal": "Factura electrónica", "keyDay": "Inmediato"},
    {"id": 13, "name": "COMFENALCO", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Fabian Campuzano", "obs": "Factura debe enviarse lunes a miércoles antes de los 20 del mes siguiente. Factura + Informe a Fabian Campuzano", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Lun-Mié antes día 20"},
    {"id": 14, "name": "COPA", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Zulay", "obs": "Se genera factura, luego de generar factura se envía factura + informe a Zulay.", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 15, "name": "EML", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Laura Vega", "obs": "Se envía la información a Laura Vega para aprobación de factura", "reqInf": True, "reqOc": True, "canal": "Correo aprobación", "keyDay": "Inmediato tras servicio"},
    {"id": 16, "name": "ESTELAR", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Sra. Luz Cano", "obs": "Se envía la información a la señora Luz Cano, debe ir por número de referencia", "reqInf": True, "reqOc": False, "canal": "Correo por Ref", "keyDay": "Inmediato tras servicio"},
    {"id": 17, "name": "EXPEDITORS", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Sol", "contact": "Patricia Delgado", "obs": "Luego de generar factura se envía factura + informe a Patricia Delgado", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Inmediato tras servicio"},
    {"id": 18, "name": "FEDERACION COLOMBIANA DE BALONCESTO", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Sra. Fanny Ceballos", "obs": "Se envía información a Fanny Ceballos, una vez apruebe pide link de pago", "reqInf": True, "reqOc": True, "canal": "Correo + Link pago", "keyDay": "Inmediato tras servicio"},
    {"id": 19, "name": "FLEISCHMANN", "prog": "CADA 16", "freqType": "FIJOS", "resp": "Edwar", "contact": "Melissa Peña", "obs": "Se envía la información cada 16 a Melissa Peña, a vuelta de correo da orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 16 de cada mes"},
    {"id": 20, "name": "FONDO DE EMPLEADOS BANCO DE OCCIDENTE", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Kevin Mauricio Castañeda", "obs": "Se envía informe a Kevin Mauricio Castañeda", "reqInf": True, "reqOc": False, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 21, "name": "FUNDACION VALLE DEL LILI", "prog": "* INMEDIATO", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "Solicitante del servicio", "obs": "Se envía el informe a quien solicito el servicio para que generen orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo a solicitante", "keyDay": "Inmediato tras servicio"},
    {"id": 22, "name": "GENFAR", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Cada solicitante", "obs": "Se envía informe a cada persona que solicita.", "reqInf": True, "reqOc": False, "canal": "Correo a solicitantes", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 23, "name": "GENFAR - PLANTA", "prog": "CADA 02-16", "freqType": "02-16", "resp": "Edwar", "contact": "Katerine Arcila", "obs": "Se envía informe a Katerine Arcila, servicios de planta y adicionales", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 24, "name": "HENRY LUIS VILLEGAS", "prog": "CADA 02-16", "freqType": "02-16", "resp": "Edwar", "contact": "Hector Mejía", "obs": "Se envía informe a Hector Mejía, luego se factura.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 25, "name": "HOTEL CASA VALLECAUCANA", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Administración Hotel", "obs": "Se envía informe y luego de la respuesta se factura.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 26, "name": "IN BOND GEMA", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Sol", "contact": "Fabian Soto", "obs": "Se envía la factura a Fabian Soto", "reqInf": False, "reqOc": False, "canal": "Correo a Fabian Soto", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 27, "name": "LATAM - BOGOTA", "prog": "RECEPCIÓN DE LATAM", "freqType": "INMEDIATO", "resp": "Edwar", "contact": "LATAM / Sistema GOW", "obs": "Se revisa información, modificaciones en GOW (si las hay) y se factura.", "reqInf": True, "reqOc": False, "canal": "Factura tras GOW", "keyDay": "Al recibir corte LATAM"},
    {"id": 28, "name": "LATAM - CALI", "prog": "CADA 15 - FIN DE MES", "freqType": "15-FIN", "resp": "Sol", "contact": "Cada Área", "obs": "Se envía informe cada área, se espera HES y se factura", "reqInf": True, "reqOc": True, "canal": "Factura tras HES", "keyDay": "Día 15 y Fin de mes"},
    {"id": 29, "name": "LATAM - EQUIPAJES", "prog": "CADA 15 - FIN DE MES", "freqType": "15-FIN", "resp": "Sol", "contact": "Cada Área", "obs": "Se envía informe cada área, se espera HES y se factura", "reqInf": True, "reqOc": True, "canal": "Factura tras HES", "keyDay": "Día 15 y Fin de mes"},
    {"id": 30, "name": "MENZIES", "prog": "CADA 23", "freqType": "FIJOS", "resp": "Sol", "contact": "Natalia Avendaño", "obs": "Se envía la factura a Natalia Avendaño", "reqInf": False, "reqOc": False, "canal": "Correo a Natalia", "keyDay": "Día 23 de cada mes"},
    {"id": 31, "name": "MINA SERVICIOS", "prog": "CADA 15", "freqType": "FIJOS", "resp": "Sol", "contact": "Natalia Jaramillo", "obs": "Se envía el informe a Natalia Jaramillo y genera orden de compra para facturar", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 15 de cada mes"},
    {"id": 32, "name": "PGI", "prog": "CADA 02-16", "freqType": "02-16", "resp": "Edwar", "contact": "Maritzabel Barrezueta", "obs": "Se envía información a Maritzabel Barrezueta, si a 5 días no hay novedad se factura", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 33, "name": "PRODUCTOS RAMO", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Sol", "contact": "Contacto Ramo", "obs": "Se envía informe y una vez aprueben se factura", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 34, "name": "QBCO", "prog": "CADA 23", "freqType": "FIJOS", "resp": "Sol", "contact": "Contacto QBCO", "obs": "Se envía informe y una vez aprueben se factura", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 23 de cada mes"},
    {"id": 35, "name": "S.O.S.", "prog": "CADA 14", "freqType": "FIJOS", "resp": "Sol", "contact": "Betty Peréz", "obs": "Luego de generar factura se envía factura + informe a Betty Peréz", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Día 14 de cada mes"},
    {"id": 36, "name": "SANOFI", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Cada solicitante", "obs": "Se envía informe a cada persona que solicita.", "reqInf": True, "reqOc": False, "canal": "Correo a solicitantes", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 37, "name": "SPIRAX", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "Administración Spirax", "obs": "Se factura y se envía el PDF con el informe", "reqInf": True, "reqOc": False, "canal": "Correo (Factura PDF + Inf)", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 38, "name": "STF -  BOGOTA RUTA", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Edwar", "contact": "asistente.comercial@studiof.com.co", "obs": "Se envían informe al correo comercial y se espera orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 39, "name": "STF -  CALI", "prog": "CADA 02-16", "freqType": "02-16", "resp": "Edwar", "contact": "Víctor / Funcionaria STF", "obs": "STF envía 01 y 16 informe. Revisar con Víctor, aprobar y esperar entrada contable. Suministros es el 19% del total.", "reqInf": True, "reqOc": True, "canal": "Factura tras Entrada Contable", "keyDay": "Día 01-02 y Día 16"},
    {"id": 40, "name": "SUMINISTROS DE LA INDUSTRIA", "prog": "CADA 02-16", "freqType": "02-16", "resp": "Edwar", "contact": "Relacionado con STF Cali", "obs": "Facturación quincenal ligada a entrada contable de STF Cali (19% del valor total).", "reqInf": False, "reqOc": True, "canal": "Factura 19% STF", "keyDay": "Día 02 y Día 16"},
    {"id": 41, "name": "VIRUTEX", "prog": "MENSUAL", "freqType": "MENSUAL", "resp": "Sol", "contact": "Administración Virutex", "obs": "Se factura y se envía el PDF con el informe", "reqInf": True, "reqOc": False, "canal": "Correo (Factura PDF + Inf)", "keyDay": "Fin de mes / Días 1-5"}
]

# Database Setup
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS client_states (
        client_id INTEGER,
        period TEXT,
        step1 BOOLEAN DEFAULT 0,
        step1_date TEXT DEFAULT '',
        step2 BOOLEAN DEFAULT 0,
        step2_oc TEXT DEFAULT '',
        step2_date TEXT DEFAULT '',
        step3 BOOLEAN DEFAULT 0,
        step3_fac TEXT DEFAULT '',
        step3_date TEXT DEFAULT '',
        step4 BOOLEAN DEFAULT 0,
        step4_date TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        updated_by TEXT DEFAULT '',
        updated_at TEXT DEFAULT '',
        PRIMARY KEY (client_id, period)
    )
    """)
    conn.commit()

    # Pre-populate defaults for active period if not exists
    default_period = "OCT_1Q"
    for c in CLIENTS_MASTER:
        cursor.execute("SELECT 1 FROM client_states WHERE client_id = ? AND period = ?", (c["id"], default_period))
        if not cursor.fetchone():
            # If not reqInf, step1 is considered true by default
            step1_def = 1 if not c["reqInf"] else 0
            step2_def = 1 if not c["reqOc"] else 0
            cursor.execute("""
            INSERT INTO client_states 
            (client_id, period, step1, step1_date, step2, step2_oc, step2_date, step3, step3_fac, step3_date, step4, step4_date, notes, updated_by, updated_at)
            VALUES (?, ?, ?, '', ?, '', '', 0, '', '', 0, '', '', 'Sistema', ?)
            """, (c["id"], default_period, step1_def, step2_def, datetime.datetime.now().isoformat()))
    conn.commit()
    conn.close()

init_db()

# WebSocket Connection Manager for Live Broadcasts
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: dict):
        dead_connections = []
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except Exception:
                dead_connections.append(connection)
        for dc in dead_connections:
            self.disconnect(dc)

manager = ConnectionManager()

# Data Models
class LoginRequest(BaseModel):
    password: str

class UpdateFieldRequest(BaseModel):
    client_id: int
    period: str
    field: str
    value: Any
    updated_by: str

# Endpoints
@app.post("/api/login")
def login(req: LoginRequest):
    pwd = req.password.strip()
    if pwd in USERS_CONFIG:
        user_info = USERS_CONFIG[pwd]
        return {
            "success": True,
            "user": user_info["name"],
            "role": user_info["role"],
            "theme": user_info["theme"],
            "filter": user_info["filter"]
        }
    raise HTTPException(status_code=401, detail="Clave incorrecta. Acceso restringido.")

@app.get("/api/clients")
def get_clients(period: str = "OCT_1Q"):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM client_states WHERE period = ?", (period,))
    rows = cursor.fetchall()
    conn.close()

    states_by_id = {r["client_id"]: dict(r) for r in rows}

    result = []
    for c in CLIENTS_MASTER:
        st = states_by_id.get(c["id"], {
            "client_id": c["id"],
            "period": period,
            "step1": False if c["reqInf"] else True,
            "step1_date": "",
            "step2": False if c["reqOc"] else True,
            "step2_oc": "",
            "step2_date": "",
            "step3": False,
            "step3_fac": "",
            "step3_date": "",
            "step4": False,
            "step4_date": "",
            "notes": "",
            "updated_by": "Sistema",
            "updated_at": ""
        })
        merged = {**c, **st}
        result.append(merged)

    return {"clients": result, "period": period}

@app.post("/api/update")
async def update_field(req: UpdateFieldRequest):
    allowed_fields = [
        "step1", "step1_date", "step2", "step2_oc", "step2_date",
        "step3", "step3_fac", "step3_date", "step4", "step4_date", "notes"
    ]
    if req.field not in allowed_fields:
        raise HTTPException(status_code=400, detail="Campo inválido")

    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Ensure row exists
    cursor.execute("SELECT 1 FROM client_states WHERE client_id = ? AND period = ?", (req.client_id, req.period))
    if not cursor.fetchone():
        cursor.execute("""
        INSERT INTO client_states (client_id, period, updated_by, updated_at)
        VALUES (?, ?, ?, ?)
        """, (req.client_id, req.period, req.updated_by, now_iso))

    sql = f"UPDATE client_states SET {req.field} = ?, updated_by = ?, updated_at = ? WHERE client_id = ? AND period = ?"
    cursor.execute(sql, (req.value, req.updated_by, now_iso, req.client_id, req.period))
    conn.commit()
    conn.close()

    # Broadcast to all connected WebSockets
    broadcast_data = {
        "type": "FIELD_UPDATED",
        "client_id": req.client_id,
        "period": req.period,
        "field": req.field,
        "value": req.value,
        "updated_by": req.updated_by,
        "updated_at": now_iso
    }
    await manager.broadcast(broadcast_data)

    return {"success": True, "data": broadcast_data}

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            data = await websocket.receive_text()
            # Heartbeat ping/pong or custom message
            if data == "ping":
                await websocket.send_text("pong")
    except WebSocketDisconnect:
        manager.disconnect(websocket)
    except Exception:
        manager.disconnect(websocket)

@app.get("/api/export/excel")
def export_excel(period: str = "OCT_1Q"):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM client_states WHERE period = ?", (period,))
    rows = {r["client_id"]: dict(r) for r in cursor.fetchall()}
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Control_{period}"

    # Header styling
    NAVY = "1E3A8A"
    STEP1_COL = "0284C7"
    STEP2_COL = "D97706"
    STEP3_COL = "059669"
    STEP4_COL = "7C3AED"

    ws.merge_cells("A1:U1")
    ws["A1"] = f"CONTROL OPERATIVO DE FACTURACIÓN E INFORMES - PERÍODO: {period}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        ("ID", NAVY), ("CLIENTE", NAVY), ("RESPONSABLE", NAVY), ("PROGRAMACIÓN", NAVY), ("DÍA CLAVE", NAVY),
        ("P1: REQ. INFORME", STEP1_COL), ("P1: FECHA ENVÍO", STEP1_COL), ("P1: ¿ENVIADO?", STEP1_COL), ("P1: CONTACTO", STEP1_COL),
        ("P2: REQ. OC", STEP2_COL), ("P2: Nº OC / HES", STEP2_COL), ("P2: FECHA OC", STEP2_COL), ("P2: ¿OC RECIBIDA?", STEP2_COL),
        ("P3: Nº FACTURA", STEP3_COL), ("P3: FECHA FACTURA", STEP3_COL), ("P3: ¿GENERADA?", STEP3_COL),
        ("P4: CANAL ENTREGA", STEP4_COL), ("P4: FECHA RADICACIÓN", STEP4_COL), ("P4: ¿ENTREGADA?", STEP4_COL),
        ("ESTADO ACTUAL", "475569"), ("ÚLTIMO RESPONSABLE", "475569")
    ]

    for col_idx, (text, color) in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=text)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    border_thin = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                         top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

    for r_idx, c in enumerate(CLIENTS_MASTER, start=4):
        st = rows.get(c["id"], {})
        s1 = "SÍ" if st.get("step1") else ("N/A" if not c["reqInf"] else "NO")
        s2 = "SÍ" if st.get("step2") else ("N/A" if not c["reqOc"] else "NO")
        s3 = "SÍ" if st.get("step3") else "NO"
        s4 = "SÍ" if st.get("step4") else "NO"

        # Calculate status
        if st.get("step4"):
            status_text = "🟢 COMPLETADO"
        elif st.get("step3"):
            status_text = "🔵 FACTURADO (Pend. Envío)"
        elif (st.get("step2") or not c["reqOc"]) and (st.get("step1") or not c["reqInf"]):
            status_text = "🟠 APROBADO (Pend. Facturar)"
        elif st.get("step1"):
            status_text = "🟡 ESPERANDO OC/APROB."
        else:
            status_text = "⚪ PENDIENTE INFORME"

        row_vals = [
            c["id"], c["name"], c["resp"], c["prog"], c["keyDay"],
            "SÍ" if c["reqInf"] else "NO", st.get("step1_date", ""), s1, c["contact"],
            "SÍ" if c["reqOc"] else "NO", st.get("step2_oc", ""), st.get("step2_date", ""), s2,
            st.get("step3_fac", ""), st.get("step3_date", ""), s3,
            c["canal"], st.get("step4_date", ""), s4,
            status_text, st.get("updated_by", "Sistema")
        ]

        fill_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_idx, column=col_idx, value=val)
            cell.border = border_thin
            cell.font = Font(name="Calibri", size=10)
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            if col_idx in [1, 3, 4, 6, 8, 10, 13, 16, 19]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Set column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Control_Facturacion_{period}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# Serve Frontend SPA
@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))

app.mount("/", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    # Bind to 0.0.0.0 so all computers on LAN (Sol, Edwar) can access!
    print("Iniciando servidor de facturación en http://0.0.0.0:8000 ...")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)
