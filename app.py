import os
import json
import sqlite3
import datetime
import uuid
from typing import List, Dict, Any, Optional
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

app = FastAPI(title="Control de Facturación - Multi-Registro en Tiempo Real")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 41 Clients Master Seed Data
CLIENTS_MASTER = [
    {"id": 1, "name": "AIR FRANCE - BOGOTA - TRIPULACIONES", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Sol", "contact": "Portal API Global", "obs": "Se genera factura, se carga en API Global", "reqInf": False, "reqOc": False, "canal": "Portal API Global", "keyDay": "Primeros 5 días"},
    {"id": 2, "name": "ANDI", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Sol", "contact": "Marisol Marín", "obs": "Se envía informe a Marisol Marín, dada la respuesta se factura.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 3, "name": "API", "prog": "* MENSUAL", "freqType": "Mensual", "resp": "Sol", "contact": "API", "obs": "Se genera factura, primeros días del mes siguiente.", "reqInf": False, "reqOc": False, "canal": "Correo electrónico", "keyDay": "Días 1-5 mes siguiente"},
    {"id": 4, "name": "APTAR", "prog": "CADA 02-16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Cristina Rojas (cc Erika Zapata)", "obs": "Se envía informe a Cristina Rojas con copia a Erika Zapata, a vuelta de correo envían orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 5, "name": "AVIANCA - BOGOTA", "prog": "CADA 05-23", "freqType": "Quincenal", "resp": "Sol", "contact": "Edison Torres, A. Peralta, L. Ramirez", "obs": "Se envía informe a Edison Torres - Alejandro Peralta - Luis Ramirez y se espera orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 05 y Día 23"},
    {"id": 6, "name": "AVIANCA - CALI", "prog": "CADA 05-23", "freqType": "Quincenal", "resp": "Sol", "contact": "Edison Torres, A. Peralta, L. Ramirez", "obs": "Se envía informe a Edison Torres - Alejandro Peralta - Luis Ramirez y se espera orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 05 y Día 23"},
    {"id": 7, "name": "AVIANCA - EQUIPAJES", "prog": "SEMANAL", "freqType": "Semanal", "resp": "Sol", "contact": "Avianca Equipajes", "obs": "Se realiza informes semanales junto los vouchers de transporte y al finalizar el mes, se debe generar un resumen en el informe.", "reqInf": True, "reqOc": True, "canal": "Correo + Vouchers", "keyDay": "Semanal + Cierre mes"},
    {"id": 8, "name": "BANCO DE LA REPUBLICA", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Don Hector", "obs": "Se genera factura de inmediato, don Hector facilita la información.", "reqInf": False, "reqOc": False, "canal": "Factura electrónica", "keyDay": "Inmediato tras servicio"},
    {"id": 9, "name": "BANCO DE OCCIDENTE", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Autorizador de viaje", "obs": "Revisar el número de viaje y la información de quién autoriza.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 10, "name": "BMS", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Mary Holguín, Paula Martinez y equipo", "obs": "Se realizan facturas, una por cada orden de compra. 1 solo informe con 1 hoja de Excel por OC.", "reqInf": True, "reqOc": True, "canal": "Correo con Excel x OC", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 11, "name": "CINEMARK COLOMBIA", "prog": "MES ANTICIPADO", "freqType": "Mensual", "resp": "Sol", "contact": "Alexander Victoria", "obs": "Se genera factura, mes anticipado y se envía factura a Alexander Victoria (más tardar a 5 del mes en curso)", "reqInf": False, "reqOc": False, "canal": "Correo electrónico", "keyDay": "Máximo día 05 mes curso"},
    {"id": 12, "name": "COLEGIO COLOMBO BRITANICO", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Administración CCB", "obs": "Se genera factura", "reqInf": False, "reqOc": False, "canal": "Factura electrónica", "keyDay": "Inmediato"},
    {"id": 13, "name": "COMFENALCO", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Fabian Campuzano", "obs": "Factura debe enviarse lunes a miércoles antes de los 20 del mes siguiente. Factura + Informe a Fabian Campuzano", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Lun-Mié antes día 20"},
    {"id": 14, "name": "COPA", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Zulay", "obs": "Se genera factura, luego de generar factura se envía factura + informe a Zulay.", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 15, "name": "EML", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Laura Vega", "obs": "Se envía la información a Laura Vega para aprobación de factura", "reqInf": True, "reqOc": True, "canal": "Correo aprobación", "keyDay": "Inmediato tras servicio"},
    {"id": 16, "name": "ESTELAR", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Sra. Luz Cano", "obs": "Se envía la información a la señora Luz Cano, debe ir por número de referencia", "reqInf": True, "reqOc": False, "canal": "Correo por Ref", "keyDay": "Inmediato tras servicio"},
    {"id": 17, "name": "EXPEDITORS", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Sol", "contact": "Patricia Delgado", "obs": "Luego de generar factura se envía factura + informe a Patricia Delgado", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Inmediato tras servicio"},
    {"id": 18, "name": "FEDERACION COLOMBIANA DE BALONCESTO", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Sra. Fanny Ceballos", "obs": "Se envía información a Fanny Ceballos, una vez apruebe pide link de pago", "reqInf": True, "reqOc": True, "canal": "Correo + Link pago", "keyDay": "Inmediato tras servicio"},
    {"id": 19, "name": "FLEISCHMANN", "prog": "CADA 16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Melissa Peña", "obs": "Se envía la información cada 16 a Melissa Peña, a vuelta de correo da orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 16 de cada mes"},
    {"id": 20, "name": "FONDO DE EMPLEADOS BANCO DE OCCIDENTE", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Kevin Mauricio Castañeda", "obs": "Se envía informe a Kevin Mauricio Castañeda", "reqInf": True, "reqOc": False, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 21, "name": "FUNDACION VALLE DEL LILI", "prog": "* INMEDIATO", "freqType": "Por Evento", "resp": "Edwar", "contact": "Solicitante del servicio", "obs": "Se envía el informe a quien solicito el servicio para que generen orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo a solicitante", "keyDay": "Inmediato tras servicio"},
    {"id": 22, "name": "GENFAR", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Cada solicitante", "obs": "Se envía informe a cada persona que solicita.", "reqInf": True, "reqOc": False, "canal": "Correo a solicitantes", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 23, "name": "GENFAR - PLANTA", "prog": "CADA 02-16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Katerine Arcila", "obs": "Se envía informe a Katerine Arcila, servicios de planta y adicionales", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 24, "name": "HENRY LUIS VILLEGAS", "prog": "CADA 02-16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Hector Mejía", "obs": "Se envía informe a Hector Mejía, luego se factura.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 25, "name": "HOTEL CASA VALLECAUCANA", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Administración Hotel", "obs": "Se envía informe y luego de la respuesta se factura.", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 26, "name": "IN BOND GEMA", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Sol", "contact": "Fabian Soto", "obs": "Se envía la factura a Fabian Soto", "reqInf": False, "reqOc": False, "canal": "Correo a Fabian Soto", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 27, "name": "LATAM - BOGOTA", "prog": "RECEPCIÓN DE LATAM", "freqType": "Por Evento", "resp": "Edwar", "contact": "LATAM / Sistema GOW", "obs": "Se revisa información, modificaciones en GOW (si las hay) y se factura.", "reqInf": True, "reqOc": False, "canal": "Factura tras GOW", "keyDay": "Al recibir corte LATAM"},
    {"id": 28, "name": "LATAM - CALI", "prog": "CADA 15 - FIN DE MES", "freqType": "Quincenal", "resp": "Sol", "contact": "Cada Área", "obs": "Se envía informe cada área, se espera HES y se factura", "reqInf": True, "reqOc": True, "canal": "Factura tras HES", "keyDay": "Día 15 y Fin de mes"},
    {"id": 29, "name": "LATAM - EQUIPAJES", "prog": "CADA 15 - FIN DE MES", "freqType": "Quincenal", "resp": "Sol", "contact": "Cada Área", "obs": "Se envía informe cada área, se espera HES y se factura", "reqInf": True, "reqOc": True, "canal": "Factura tras HES", "keyDay": "Día 15 y Fin de mes"},
    {"id": 30, "name": "MENZIES", "prog": "CADA 23", "freqType": "Quincenal", "resp": "Sol", "contact": "Natalia Avendaño", "obs": "Se envía la factura a Natalia Avendaño", "reqInf": False, "reqOc": False, "canal": "Correo a Natalia", "keyDay": "Día 23 de cada mes"},
    {"id": 31, "name": "MINA SERVICIOS", "prog": "CADA 15", "freqType": "Quincenal", "resp": "Sol", "contact": "Natalia Jaramillo", "obs": "Se envía el informe a Natalia Jaramillo y genera orden de compra para facturar", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 15 de cada mes"},
    {"id": 32, "name": "PGI", "prog": "CADA 02-16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Maritzabel Barrezueta", "obs": "Se envía información a Maritzabel Barrezueta, si a 5 días no hay novedad se factura", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 02 y Día 16"},
    {"id": 33, "name": "PRODUCTOS RAMO", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Sol", "contact": "Contacto Ramo", "obs": "Se envía informe y una vez aprueben se factura", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 34, "name": "QBCO", "prog": "CADA 23", "freqType": "Quincenal", "resp": "Sol", "contact": "Contacto QBCO", "obs": "Se envía informe y una vez aprueben se factura", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Día 23 de cada mes"},
    {"id": 35, "name": "S.O.S.", "prog": "CADA 14", "freqType": "Quincenal", "resp": "Sol", "contact": "Betty Peréz", "obs": "Luego de generar factura se envía factura + informe a Betty Peréz", "reqInf": True, "reqOc": False, "canal": "Correo (Factura + Inf)", "keyDay": "Día 14 de cada mes"},
    {"id": 36, "name": "SANOFI", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Cada solicitante", "obs": "Se envía informe a cada persona que solicita.", "reqInf": True, "reqOc": False, "canal": "Correo a solicitantes", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 37, "name": "SPIRAX", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "Administración Spirax", "obs": "Se factura y se envía el PDF con el informe", "reqInf": True, "reqOc": False, "canal": "Correo (Factura PDF + Inf)", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 38, "name": "STF -  BOGOTA RUTA", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Edwar", "contact": "asistente.comercial@studiof.com.co", "obs": "Se envían informe al correo comercial y se espera orden de compra", "reqInf": True, "reqOc": True, "canal": "Correo electrónico", "keyDay": "Fin de mes / Días 1-5"},
    {"id": 39, "name": "STF -  CALI", "prog": "CADA 02-16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Víctor / Funcionaria STF", "obs": "STF envía 01 y 16 informe. Revisar con Víctor, aprobar y esperar entrada contable. Suministros es el 19% del total.", "reqInf": True, "reqOc": True, "canal": "Factura tras Entrada Contable", "keyDay": "Día 01-02 y Día 16"},
    {"id": 40, "name": "SUMINISTROS DE LA INDUSTRIA", "prog": "CADA 02-16", "freqType": "Quincenal", "resp": "Edwar", "contact": "Relacionado con STF Cali", "obs": "Facturación quincenal ligada a entrada contable de STF Cali (19% del valor total).", "reqInf": False, "reqOc": True, "canal": "Factura 19% STF", "keyDay": "Día 02 y Día 16"},
    {"id": 41, "name": "VIRUTEX", "prog": "MENSUAL", "freqType": "Mensual", "resp": "Sol", "contact": "Administración Virutex", "obs": "Se factura y se envía el PDF con el informe", "reqInf": True, "reqOc": False, "canal": "Correo (Factura PDF + Inf)", "keyDay": "Fin de mes / Días 1-5"}
]

# Database Setup & Migrations
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # 1. System Users Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        name TEXT,
        role TEXT DEFAULT 'admin',
        cargo TEXT DEFAULT 'Analista de Facturación',
        theme TEXT DEFAULT 'edwar',
        filter TEXT DEFAULT 'ALL',
        is_active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT '',
        updated_at TEXT DEFAULT ''
    )
    """)

    # 2. Clients Table (Persistent Client Master)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY,
        name TEXT,
        prog TEXT,
        freq_type TEXT,
        resp TEXT,
        contact TEXT,
        obs TEXT,
        req_inf INTEGER DEFAULT 1,
        req_oc INTEGER DEFAULT 1,
        canal TEXT DEFAULT 'Correo electrónico',
        key_day TEXT DEFAULT '',
        is_blocked INTEGER DEFAULT 0
    )
    """)

    # 3. Cargos Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cargos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT UNIQUE
    )
    """)

    # 4. Matrix Cuatrimestre Tracking Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS matrix_cells (
        client_id INTEGER,
        period_col TEXT,
        status TEXT DEFAULT 'PENDIENTE',
        invoice_num TEXT DEFAULT '',
        invoice_val TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        updated_by TEXT DEFAULT '',
        updated_at TEXT DEFAULT '',
        PRIMARY KEY (client_id, period_col)
    )
    """)

    # 5. Primary Multi-Entry Billing Records Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS billing_records (
        record_id TEXT PRIMARY KEY,
        client_id INTEGER,
        client_name TEXT,
        freq_type TEXT DEFAULT 'Mensual',
        month TEXT DEFAULT 'Octubre',
        period_detail TEXT DEFAULT 'Mes Completo',
        period_key TEXT DEFAULT 'OCT_1Q',
        step1 BOOLEAN DEFAULT 0,
        step1_date TEXT DEFAULT '',
        pref_num TEXT DEFAULT '',
        pref_val TEXT DEFAULT '',
        step2 BOOLEAN DEFAULT 0,
        step2_oc TEXT DEFAULT '',
        step2_date TEXT DEFAULT '',
        oc_val TEXT DEFAULT '',
        step3 BOOLEAN DEFAULT 0,
        step3_fac TEXT DEFAULT '',
        step3_date TEXT DEFAULT '',
        fac_val TEXT DEFAULT '',
        step4 BOOLEAN DEFAULT 0,
        step4_date TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        updated_by TEXT DEFAULT '',
        updated_at TEXT DEFAULT ''
    )
    """)
    conn.commit()

    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Seed Default Users if empty
    cursor.execute("SELECT COUNT(*) FROM system_users")
    if cursor.fetchone()[0] == 0:
        default_users = [
            ("master", ".soporte", "Administrador Master", "master", "Administrador Master del Sistema", "edwar", "ALL"),
            ("1143851865", "1143851865", "Sol", "admin", "Especialista de Facturación & Cartera", "sol", "Sol"),
            ("1107078586", "1107078586", "Edwar", "admin", "Coordinador de Facturación Operativa", "edwar", "Edwar"),
            ("1234", "1234", "Invitado", "viewer", "Consultor Externo / Modo Lectura", "guest", "ALL")
        ]
        for u, p, n, r, c, t, f in default_users:
            cursor.execute("""
            INSERT INTO system_users (username, password, name, role, cargo, theme, filter, is_active, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
            """, (u, p, n, r, c, t, f, now_iso, now_iso))
        conn.commit()

    # Ensure Master User exists even if table was pre-existing
    cursor.execute("SELECT COUNT(*) FROM system_users WHERE role = 'master'")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
        INSERT INTO system_users (username, password, name, role, cargo, theme, filter, is_active, created_at, updated_at)
        VALUES ('master', '.soporte', 'Administrador Master', 'master', 'Administrador Master del Sistema', 'edwar', 'ALL', 1, ?, ?)
        """, (now_iso, now_iso))
        conn.commit()

    # Seed Default Clients if empty
    cursor.execute("SELECT COUNT(*) FROM clients")
    if cursor.fetchone()[0] == 0:
        for c in CLIENTS_MASTER:
            cursor.execute("""
            INSERT INTO clients (id, name, prog, freq_type, resp, contact, obs, req_inf, req_oc, canal, key_day, is_blocked)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
            """, (c["id"], c["name"], c["prog"], c["freqType"], c["resp"], c["contact"], c["obs"],
                  1 if c["reqInf"] else 0, 1 if c["reqOc"] else 0, c["canal"], c["keyDay"]))
        conn.commit()

    # Seed Default Cargos if empty
    cursor.execute("SELECT COUNT(*) FROM cargos")
    if cursor.fetchone()[0] == 0:
        default_cargos = [
            "Administrador Master del Sistema",
            "Especialista de Facturación & Cartera",
            "Coordinador de Facturación Operativa",
            "Analista Senior de Facturación",
            "Analista Junior de Facturación",
            "Líder de Facturación & Cartera",
            "Asistente de Facturación",
            "Consultor Externo / Modo Lectura"
        ]
        for cargo in default_cargos:
            cursor.execute("INSERT OR IGNORE INTO cargos (title) VALUES (?)", (cargo,))
        conn.commit()

    # Seed Default Billing Records if empty
    cursor.execute("SELECT COUNT(*) FROM billing_records")
    if cursor.fetchone()[0] == 0:
        cursor.execute("SELECT * FROM clients ORDER BY id ASC")
        c_rows = cursor.fetchall()
        for c in c_rows:
            cid, cname, cprog, cfreq, cresp, ccontact, cobs, creq_inf, creq_oc, ccanal, ckey_day, cblocked = c
            rec_id = f"rec_{cid}_default"
            p_prog = cprog or ""
            if "02-16" in p_prog or "05-23" in p_prog or "15" in p_prog:
                f_type = "Quincenal"
                p_detail = "1Q"
            elif "SEMANAL" in p_prog:
                f_type = "Semanal"
                p_detail = "Semana 1"
            elif "INMEDIATO" in p_prog:
                f_type = "Por Evento"
                p_detail = "Servicio"
            else:
                f_type = "Mensual"
                p_detail = "Mes Completo"

            step1_def = 1 if not creq_inf else 0
            step2_def = 1 if not creq_oc else 0

            cursor.execute("""
            INSERT INTO billing_records
            (record_id, client_id, client_name, freq_type, month, period_detail, period_key, step1, step1_date, pref_num, pref_val, step2, step2_oc, step2_date, oc_val, step3, step3_fac, step3_date, fac_val, step4, step4_date, notes, updated_by, updated_at)
            VALUES (?, ?, ?, ?, 'Octubre', ?, 'OCT_1Q', ?, '', '', '', ?, '', '', '', 0, '', '', '', 0, '', '', 'Sistema', ?)
            """, (rec_id, cid, cname, f_type, p_detail, step1_def, step2_def, now_iso))
        conn.commit()

    # Seed Matrix Cells if empty
    cursor.execute("SELECT COUNT(*) FROM matrix_cells")
    if cursor.fetchone()[0] == 0:
        cursor.execute("SELECT id FROM clients")
        c_ids = [r[0] for r in cursor.fetchall()]
        cols = ["sep_1q", "sep_2q", "oct_1q", "oct_2q", "nov_1q", "nov_2q", "dic_1q", "dic_2q"]
        for cid in c_ids:
            for col in cols:
                # Default status
                cursor.execute("""
                INSERT OR IGNORE INTO matrix_cells (client_id, period_col, status, invoice_num, invoice_val, updated_by, updated_at)
                VALUES (?, ?, 'PENDIENTE', '', '', 'Sistema', ?)
                """, (cid, col, now_iso))
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
    username: Optional[str] = None

class UpdateProfileRequest(BaseModel):
    username: str
    new_password: Optional[str] = None
    theme: Optional[str] = None

class CreateRecordRequest(BaseModel):
    client_id: int
    freq_type: str = "Quincenal"
    month: str = "Octubre"
    period_detail: str = "1Q"
    updated_by: str = "Sistema"

class UpdateRecordFieldRequest(BaseModel):
    record_id: str
    field: str
    value: Any
    updated_by: str

class AdminUserCreateRequest(BaseModel):
    username: str
    password: str
    name: str
    role: str = "admin"
    cargo: str = "Analista de Facturación"
    theme: str = "edwar"
    filter: str = "ALL"

class AdminUserUpdateRequest(BaseModel):
    id: int
    username: str
    password: str
    name: str
    role: str
    cargo: str
    theme: str
    filter: str
    is_active: int = 1

class AdminClientCreateRequest(BaseModel):
    name: str
    prog: str
    freq_type: str
    resp: str
    contact: str
    obs: str
    req_inf: bool = True
    req_oc: bool = True
    canal: str = "Correo electrónico"
    key_day: str = ""

class AdminClientUpdateRequest(BaseModel):
    id: int
    name: str
    prog: str
    freq_type: str
    resp: str
    contact: str
    obs: str
    req_inf: bool
    req_oc: bool
    canal: str
    key_day: str
    is_blocked: int = 0

class AdminToggleBlockRequest(BaseModel):
    id: int

class CargoCreateRequest(BaseModel):
    title: str

class MatrixUpdateRequest(BaseModel):
    client_id: int
    period_col: str
    status: str
    invoice_num: str = ""
    invoice_val: str = ""
    updated_by: str = "Sistema"

# Endpoints
@app.post("/api/login")
def login(req: LoginRequest):
    pwd = (req.password or "").strip()
    usr = (req.username or "").strip()

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    user_row = None
    # 1. If username and password provided
    if usr and pwd:
        cursor.execute("SELECT * FROM system_users WHERE (username = ? OR password = ?) AND password = ? AND is_active = 1", (usr, pwd, pwd))
        user_row = cursor.fetchone()
    elif pwd:
        # Check by password
        cursor.execute("SELECT * FROM system_users WHERE password = ? AND is_active = 1", (pwd,))
        user_row = cursor.fetchone()
        if not user_row:
            # Check by username/cédula
            cursor.execute("SELECT * FROM system_users WHERE username = ? AND is_active = 1", (pwd,))
            user_row = cursor.fetchone()
    
    conn.close()

    if user_row:
        return {
            "success": True,
            "user": user_row["name"],
            "username": user_row["username"],
            "role": user_row["role"],
            "cargo": user_row["cargo"],
            "theme": user_row["theme"],
            "filter": user_row["filter"]
        }
    raise HTTPException(status_code=401, detail="Credenciales no autorizadas o usuario inactivo.")

@app.post("/api/user/update-profile")
async def update_profile(req: UpdateProfileRequest):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM system_users WHERE username = ?", (req.username,))
    user_row = cursor.fetchone()
    if not user_row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuario no encontrado")

    new_pwd = user_row["password"]
    if req.new_password and req.new_password.strip():
        new_pwd = req.new_password.strip()

    new_theme = user_row["theme"]
    if req.theme and req.theme in ["sol", "edwar"]:
        if user_row["role"] != "viewer":
            new_theme = req.theme

    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("UPDATE system_users SET password = ?, theme = ?, updated_at = ? WHERE username = ?",
                   (new_pwd, new_theme, now_iso, req.username))
    conn.commit()
    conn.close()

    return {
        "success": True,
        "username": req.username,
        "theme": new_theme,
        "message": "Perfil actualizado exitosamente"
    }

# Master Admin Users Management
@app.get("/api/admin/users")
def get_admin_users():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, password, name, role, cargo, theme, filter, is_active FROM system_users ORDER BY id ASC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return {"users": rows}

@app.post("/api/admin/users/create")
def create_admin_user(req: AdminUserCreateRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        cursor.execute("""
        INSERT INTO system_users (username, password, name, role, cargo, theme, filter, is_active, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        """, (req.username.strip(), req.password.strip(), req.name.strip(), req.role, req.cargo, req.theme, req.filter, now_iso, now_iso))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        raise HTTPException(status_code=400, detail="El usuario o cédula ya existe en el sistema.")
    conn.close()
    return {"success": True, "message": "Usuario creado exitosamente"}

@app.post("/api/admin/users/update")
def update_admin_user(req: AdminUserUpdateRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
    UPDATE system_users SET username = ?, password = ?, name = ?, role = ?, cargo = ?, theme = ?, filter = ?, is_active = ?, updated_at = ?
    WHERE id = ?
    """, (req.username.strip(), req.password.strip(), req.name.strip(), req.role, req.cargo, req.theme, req.filter, req.is_active, now_iso, req.id))
    conn.commit()
    conn.close()
    return {"success": True, "message": "Usuario actualizado exitosamente"}

@app.delete("/api/admin/users/{user_id}")
def delete_admin_user(user_id: int):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT role FROM system_users WHERE id = ?", (user_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    if row[0] == "master":
        conn.close()
        raise HTTPException(status_code=400, detail="No se puede eliminar el usuario Master principal.")
    cursor.execute("DELETE FROM system_users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    return {"success": True, "message": "Usuario eliminado"}

# Master Admin Clients / Companies Management
@app.get("/api/admin/clients")
def get_admin_clients():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM clients ORDER BY id ASC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return {"clients": rows}

@app.post("/api/admin/clients/create")
async def create_admin_client(req: AdminClientCreateRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(id) FROM clients")
    max_id = cursor.fetchone()[0] or 0
    new_id = max_id + 1

    cursor.execute("""
    INSERT INTO clients (id, name, prog, freq_type, resp, contact, obs, req_inf, req_oc, canal, key_day, is_blocked)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
    """, (new_id, req.name.strip(), req.prog.strip(), req.freq_type, req.resp, req.contact.strip(), req.obs.strip(),
          1 if req.req_inf else 0, 1 if req.req_oc else 0, req.canal.strip(), req.key_day.strip()))
    conn.commit()

    # Also create default billing record
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rec_id = f"rec_{new_id}_default"
    p_detail = "1Q" if req.freq_type == "Quincenal" else ("Semana 1" if req.freq_type == "Semanal" else ("Servicio" if req.freq_type == "Por Evento" else "Mes Completo"))
    step1_def = 0 if req.req_inf else 1
    step2_def = 0 if req.req_oc else 1
    cursor.execute("""
    INSERT INTO billing_records
    (record_id, client_id, client_name, freq_type, month, period_detail, period_key, step1, step2, updated_by, updated_at)
    VALUES (?, ?, ?, ?, 'Octubre', ?, 'OCT_1Q', ?, ?, 'Master', ?)
    """, (rec_id, new_id, req.name.strip(), req.freq_type, p_detail, step1_def, step2_def, now_iso))

    # Also insert default matrix rows
    cols = ["sep_1q", "sep_2q", "oct_1q", "oct_2q", "nov_1q", "nov_2q", "dic_1q", "dic_2q"]
    for col in cols:
        cursor.execute("INSERT OR IGNORE INTO matrix_cells (client_id, period_col, status, updated_by, updated_at) VALUES (?, ?, 'PENDIENTE', 'Master', ?)", (new_id, col, now_iso))

    conn.commit()
    conn.close()

    await manager.broadcast({"type": "CLIENT_CREATED", "client_id": new_id, "name": req.name})
    return {"success": True, "id": new_id, "message": "Empresa creada exitosamente"}

@app.post("/api/admin/clients/update")
async def update_admin_client(req: AdminClientUpdateRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE clients SET name = ?, prog = ?, freq_type = ?, resp = ?, contact = ?, obs = ?, req_inf = ?, req_oc = ?, canal = ?, key_day = ?, is_blocked = ?
    WHERE id = ?
    """, (req.name.strip(), req.prog.strip(), req.freq_type, req.resp, req.contact.strip(), req.obs.strip(),
          1 if req.req_inf else 0, 1 if req.req_oc else 0, req.canal.strip(), req.key_day.strip(), req.is_blocked, req.id))
    cursor.execute("UPDATE billing_records SET client_name = ? WHERE client_id = ?", (req.name.strip(), req.id))
    conn.commit()
    conn.close()

    await manager.broadcast({"type": "CLIENT_UPDATED", "client_id": req.id, "name": req.name})
    return {"success": True, "message": "Empresa actualizada"}

@app.post("/api/admin/clients/toggle-block")
async def toggle_block_client(req: AdminToggleBlockRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE clients SET is_blocked = CASE WHEN is_blocked = 1 THEN 0 ELSE 1 END WHERE id = ?", (req.id,))
    cursor.execute("SELECT is_blocked, name FROM clients WHERE id = ?", (req.id,))
    row = cursor.fetchone()
    conn.commit()
    conn.close()
    if not row:
        raise HTTPException(status_code=404, detail="Empresa no encontrada")

    await manager.broadcast({"type": "CLIENT_BLOCKED_TOGGLED", "client_id": req.id, "is_blocked": row[0]})
    return {"success": True, "is_blocked": row[0], "name": row[1]}

@app.delete("/api/admin/clients/{client_id}")
async def delete_admin_client(client_id: int):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM clients WHERE id = ?", (client_id,))
    cursor.execute("DELETE FROM billing_records WHERE client_id = ?", (client_id,))
    cursor.execute("DELETE FROM matrix_cells WHERE client_id = ?", (client_id,))
    conn.commit()
    conn.close()

    await manager.broadcast({"type": "CLIENT_DELETED", "client_id": client_id})
    return {"success": True, "message": "Empresa eliminada"}

# Cargos Endpoints
@app.get("/api/cargos")
def get_cargos():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT title FROM cargos ORDER BY id ASC")
    cargos = [r[0] for r in cursor.fetchall()]
    conn.close()
    return {"cargos": cargos}

@app.post("/api/cargos/create")
def create_cargo(req: CargoCreateRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO cargos (title) VALUES (?)", (req.title.strip(),))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        raise HTTPException(status_code=400, detail="Este cargo ya existe")
    conn.close()
    return {"success": True, "title": req.title.strip()}

# Matrix Endpoints
@app.get("/api/matrix")
def get_matrix_data():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM clients ORDER BY id ASC")
    clients = [dict(c) for c in cursor.fetchall()]

    cursor.execute("SELECT * FROM matrix_cells")
    cells = cursor.fetchall()
    conn.close()

    matrix_map = {}
    for c in cells:
        cid = c["client_id"]
        if cid not in matrix_map:
            matrix_map[cid] = {}
        matrix_map[cid][c["period_col"]] = {
            "status": c["status"],
            "invoice_num": c["invoice_num"],
            "invoice_val": c["invoice_val"],
            "notes": c["notes"],
            "updated_by": c["updated_by"],
            "updated_at": c["updated_at"]
        }

    return {
        "clients": clients,
        "matrix": matrix_map
    }

@app.post("/api/matrix/update")
async def update_matrix_cell(req: MatrixUpdateRequest):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
    INSERT INTO matrix_cells (client_id, period_col, status, invoice_num, invoice_val, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(client_id, period_col) DO UPDATE SET
    status = excluded.status,
    invoice_num = excluded.invoice_num,
    invoice_val = excluded.invoice_val,
    updated_by = excluded.updated_by,
    updated_at = excluded.updated_at
    """, (req.client_id, req.period_col, req.status, req.invoice_num, req.invoice_val, req.updated_by, now_iso))
    conn.commit()
    conn.close()

    broadcast_data = {
        "type": "MATRIX_CELL_UPDATED",
        "client_id": req.client_id,
        "period_col": req.period_col,
        "status": req.status,
        "invoice_num": req.invoice_num,
        "invoice_val": req.invoice_val,
        "updated_by": req.updated_by,
        "updated_at": now_iso
    }
    await manager.broadcast(broadcast_data)
    return {"success": True, "data": broadcast_data}

# Backward Compatibility Alias
@app.get("/api/clients")
def get_clients_alias(period: Optional[str] = None):
    res = get_records()
    res["clients"] = res["records"]
    return res

@app.post("/api/update")
async def update_alias(req: dict):
    record_id = req.get("record_id")
    client_id = req.get("client_id")
    field = req.get("field")
    value = req.get("value")
    updated_by = req.get("updated_by", "Sistema")
    
    if not record_id and client_id:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT record_id FROM billing_records WHERE client_id = ? ORDER BY record_id ASC LIMIT 1", (client_id,))
        row = c.fetchone()
        conn.close()
        if row:
            record_id = row[0]
            
    if not record_id:
        raise HTTPException(status_code=400, detail="record_id requerido")
        
    return await update_record_field(UpdateRecordFieldRequest(
        record_id=record_id,
        field=field,
        value=value,
        updated_by=updated_by
    ))

@app.get("/api/records")
def get_records(month: Optional[str] = None):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # Load clients from clients table
    cursor.execute("SELECT * FROM clients ORDER BY id ASC")
    c_rows = cursor.fetchall()
    client_dict = {}
    for c in c_rows:
        client_dict[c["id"]] = {
            "id": c["id"],
            "name": c["name"],
            "prog": c["prog"],
            "freqType": c["freq_type"],
            "resp": c["resp"],
            "contact": c["contact"],
            "obs": c["obs"],
            "reqInf": bool(c["req_inf"]),
            "reqOc": bool(c["req_oc"]),
            "canal": c["canal"],
            "keyDay": c["key_day"],
            "is_blocked": bool(c["is_blocked"])
        }

    if month and month != "ALL":
        cursor.execute("SELECT * FROM billing_records WHERE month = ? ORDER BY client_id ASC, record_id ASC", (month,))
    else:
        cursor.execute("SELECT * FROM billing_records ORDER BY client_id ASC, record_id ASC")
    
    rows = cursor.fetchall()
    
    # Also fetch all rows to compute duplicate maps across the entire system
    cursor.execute("SELECT record_id, client_id, client_name, month, period_detail, pref_num, pref_val, step2_oc, oc_val, step3_fac, fac_val FROM billing_records")
    all_rows = cursor.fetchall()
    conn.close()

    # Build Duplicate & Aggregation Maps
    # 1. Invoices Map
    invoices_map = {}
    for r in all_rows:
        fac = str(r["step3_fac"] or "").strip().upper()
        if fac:
            if fac not in invoices_map:
                invoices_map[fac] = []
            invoices_map[fac].append({
                "record_id": r["record_id"],
                "client_id": r["client_id"],
                "client_name": r["client_name"],
                "period": f"{r['month']} - {r['period_detail']}",
                "fac_val": r["fac_val"]
            })

    # 2. Pre-invoices Map
    preinvoices_map = {}
    for r in all_rows:
        pref = str(r["pref_num"] or "").strip().upper()
        if pref:
            if pref not in preinvoices_map:
                preinvoices_map[pref] = []
            preinvoices_map[pref].append({
                "record_id": r["record_id"],
                "client_id": r["client_id"],
                "client_name": r["client_name"],
                "period": f"{r['month']} - {r['period_detail']}",
                "pref_val": r["pref_val"]
            })

    # 3. OCs Map
    ocs_map = {}
    for r in all_rows:
        oc = str(r["step2_oc"] or "").strip().upper()
        if oc:
            if oc not in ocs_map:
                ocs_map[oc] = []
            ocs_map[oc].append({
                "record_id": r["record_id"],
                "client_id": r["client_id"],
                "client_name": r["client_name"],
                "period": f"{r['month']} - {r['period_detail']}",
                "fac_val": r["fac_val"],
                "oc_val": r["oc_val"]
            })

    # Merge client metadata
    records = []
    for r in rows:
        rec_dict = dict(r)
        c_meta = client_dict.get(rec_dict["client_id"], {})
        merged = {**c_meta, **rec_dict}
        records.append(merged)

    return {
        "records": records,
        "clients": list(client_dict.values()),
        "invoices_map": invoices_map,
        "preinvoices_map": preinvoices_map,
        "ocs_map": ocs_map
    }

@app.post("/api/records/create")
async def create_record(req: CreateRecordRequest):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM clients WHERE id = ?", (req.client_id,))
    c_row = cursor.fetchone()
    if not c_row:
        conn.close()
        raise HTTPException(status_code=404, detail="Cliente no encontrado")

    c_meta = dict(c_row)
    new_id = f"rec_{req.client_id}_{uuid.uuid4().hex[:8]}"
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    step1_def = 1 if not c_meta["req_inf"] else 0
    step2_def = 1 if not c_meta["req_oc"] else 0

    cursor.execute("""
    INSERT INTO billing_records
    (record_id, client_id, client_name, freq_type, month, period_detail, period_key, step1, step1_date, pref_num, pref_val, step2, step2_oc, step2_date, oc_val, step3, step3_fac, step3_date, fac_val, step4, step4_date, notes, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, 'CUSTOM', ?, '', '', '', ?, '', '', '', 0, '', '', '', 0, '', '', ?, ?)
    """, (new_id, req.client_id, c_meta["name"], req.freq_type, req.month, req.period_detail, step1_def, step2_def, req.updated_by, now_iso))
    conn.commit()
    conn.close()

    new_record = {
        "id": c_meta["id"],
        "name": c_meta["name"],
        "prog": c_meta["prog"],
        "freqType": c_meta["freq_type"],
        "resp": c_meta["resp"],
        "contact": c_meta["contact"],
        "obs": c_meta["obs"],
        "reqInf": bool(c_meta["req_inf"]),
        "reqOc": bool(c_meta["req_oc"]),
        "canal": c_meta["canal"],
        "keyDay": c_meta["key_day"],
        "is_blocked": bool(c_meta["is_blocked"]),
        "record_id": new_id,
        "client_id": req.client_id,
        "client_name": c_meta["name"],
        "freq_type": req.freq_type,
        "month": req.month,
        "period_detail": req.period_detail,
        "step1": step1_def,
        "step1_date": "",
        "pref_num": "",
        "pref_val": "",
        "step2": step2_def,
        "step2_oc": "",
        "step2_date": "",
        "oc_val": "",
        "step3": 0,
        "step3_fac": "",
        "step3_date": "",
        "fac_val": "",
        "step4": 0,
        "step4_date": "",
        "notes": "",
        "updated_by": req.updated_by,
        "updated_at": now_iso
    }

    # Broadcast creation to all connected peers
    await manager.broadcast({
        "type": "RECORD_CREATED",
        "record": new_record,
        "author": req.updated_by
    })

    return {"success": True, "record": new_record}

@app.post("/api/records/update")
async def update_record_field(req: UpdateRecordFieldRequest):
    allowed_fields = [
        "freq_type", "month", "period_detail",
        "step1", "step1_date", "pref_num", "pref_val",
        "step2", "step2_oc", "step2_date", "oc_val",
        "step3", "step3_fac", "step3_date", "fac_val",
        "step4", "step4_date", "notes"
    ]
    if req.field not in allowed_fields:
        raise HTTPException(status_code=400, detail="Campo no permitido")

    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    sql = f"UPDATE billing_records SET {req.field} = ?, updated_by = ?, updated_at = ? WHERE record_id = ?"
    cursor.execute(sql, (req.value, req.updated_by, now_iso, req.record_id))
    conn.commit()
    conn.close()

    # Broadcast field update to all connected peers
    broadcast_data = {
        "type": "RECORD_FIELD_UPDATED",
        "record_id": req.record_id,
        "field": req.field,
        "value": req.value,
        "updated_by": req.updated_by,
        "updated_at": now_iso
    }
    await manager.broadcast(broadcast_data)

    return {"success": True, "data": broadcast_data}

@app.delete("/api/records/{record_id}")
async def delete_record(record_id: str, user: str = "Sistema"):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT client_id, client_name FROM billing_records WHERE record_id = ?", (record_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Registro no encontrado")

    client_id = row[0]
    cursor.execute("SELECT COUNT(*) FROM billing_records WHERE client_id = ?", (client_id,))
    cnt = cursor.fetchone()[0]
    if cnt <= 1:
        conn.close()
        raise HTTPException(status_code=400, detail="No se puede eliminar la única fila del cliente. Debe quedar al menos un registro.")

    cursor.execute("DELETE FROM billing_records WHERE record_id = ?", (record_id,))
    conn.commit()
    conn.close()

    await manager.broadcast({
        "type": "RECORD_DELETED",
        "record_id": record_id,
        "client_id": client_id,
        "author": user
    })

    return {"success": True, "deleted_id": record_id}

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            data = await websocket.receive_text()
            if data == "ping":
                await websocket.send_text("pong")
    except WebSocketDisconnect:
        manager.disconnect(websocket)
    except Exception:
        manager.disconnect(websocket)

# Backup, Restore & LocalStorage Sync Endpoints
@app.get("/api/backup/export")
def export_backup():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM clients")
    clients = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM billing_records")
    records = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM matrix_cells")
    matrix = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM cargos")
    cargos = [r["title"] for r in cursor.fetchall()]

    cursor.execute("SELECT id, username, password, name, role, cargo, theme, filter, is_active FROM system_users")
    users = [dict(r) for r in cursor.fetchall()]

    conn.close()

    backup_data = {
        "version": "1.0",
        "exported_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "clients": clients,
        "billing_records": records,
        "matrix_cells": matrix,
        "cargos": cargos,
        "users": users
    }
    return backup_data

@app.post("/api/backup/import")
async def import_backup(data: dict):
    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="Formato de respaldo inválido")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Restore clients
    if "clients" in data and isinstance(data["clients"], list):
        for c in data["clients"]:
            cursor.execute("""
            INSERT OR REPLACE INTO clients (id, name, prog, freq_type, resp, contact, obs, req_inf, req_oc, canal, key_day, is_blocked)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (c.get("id"), c.get("name"), c.get("prog"), c.get("freq_type"), c.get("resp"), c.get("contact"), c.get("obs"),
                  1 if c.get("req_inf") else 0, 1 if c.get("req_oc") else 0, c.get("canal", ""), c.get("key_day", ""), 1 if c.get("is_blocked") else 0))

    # Restore billing records
    if "billing_records" in data and isinstance(data["billing_records"], list):
        for r in data["billing_records"]:
            cursor.execute("""
            INSERT OR REPLACE INTO billing_records
            (record_id, client_id, client_name, freq_type, month, period_detail, period_key, step1, step1_date, pref_num, pref_val, step2, step2_oc, step2_date, oc_val, step3, step3_fac, step3_date, fac_val, step4, step4_date, notes, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                r.get("record_id"), r.get("client_id"), r.get("client_name"), r.get("freq_type"), r.get("month"), r.get("period_detail"),
                r.get("period_key", "OCT_1Q"), 1 if r.get("step1") else 0, r.get("step1_date", ""), r.get("pref_num", ""), r.get("pref_val", ""),
                1 if r.get("step2") else 0, r.get("step2_oc", ""), r.get("step2_date", ""), r.get("oc_val", ""),
                1 if r.get("step3") else 0, r.get("step3_fac", ""), r.get("step3_date", ""), r.get("fac_val", ""),
                1 if r.get("step4") else 0, r.get("step4_date", ""), r.get("notes", ""), r.get("updated_by", "Restauración"), now_iso
            ))

    # Restore matrix cells
    if "matrix_cells" in data and isinstance(data["matrix_cells"], list):
        for m in data["matrix_cells"]:
            cursor.execute("""
            INSERT OR REPLACE INTO matrix_cells (client_id, period_col, status, invoice_num, invoice_val, notes, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                m.get("client_id"), m.get("period_col"), m.get("status", "PENDIENTE"),
                m.get("invoice_num", ""), m.get("invoice_val", ""), m.get("notes", ""),
                m.get("updated_by", "Restauración"), now_iso
            ))

    # Restore cargos
    if "cargos" in data and isinstance(data["cargos"], list):
        for title in data["cargos"]:
            cursor.execute("INSERT OR IGNORE INTO cargos (title) VALUES (?)", (title,))

    # Restore users (preserve master)
    if "users" in data and isinstance(data["users"], list):
        for u in data["users"]:
            if u.get("role") == "master":
                cursor.execute("UPDATE system_users SET password = ? WHERE role = 'master'", (u.get("password"),))
            else:
                cursor.execute("""
                INSERT OR REPLACE INTO system_users (id, username, password, name, role, cargo, theme, filter, is_active, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    u.get("id"), u.get("username"), u.get("password"), u.get("name"), u.get("role", "admin"),
                    u.get("cargo", "Analista"), u.get("theme", "edwar"), u.get("filter", "ALL"),
                    1 if u.get("is_active", 1) else 0, now_iso, now_iso
                ))

    conn.commit()
    conn.close()

    await manager.broadcast({"type": "DATABASE_RESTORED"})
    return {"success": True, "message": "Base de datos restaurada correctamente"}

@app.post("/api/sync/client_cache")
async def sync_client_cache(payload: dict):
    records = payload.get("records", [])
    if not records:
        return {"success": True, "synced": 0}

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    synced_count = 0
    for r in records:
        rec_id = r.get("record_id")
        if not rec_id:
            continue
        cursor.execute("""
        INSERT OR REPLACE INTO billing_records
        (record_id, client_id, client_name, freq_type, month, period_detail, period_key, step1, step1_date, pref_num, pref_val, step2, step2_oc, step2_date, oc_val, step3, step3_fac, step3_date, fac_val, step4, step4_date, notes, updated_by, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            rec_id, r.get("client_id"), r.get("client_name"), r.get("freq_type", "Mensual"), r.get("month", "Octubre"),
            r.get("period_detail", "Mes Completo"), r.get("period_key", "OCT_1Q"),
            1 if r.get("step1") else 0, r.get("step1_date", ""), r.get("pref_num", ""), r.get("pref_val", ""),
            1 if r.get("step2") else 0, r.get("step2_oc", ""), r.get("step2_date", ""), r.get("oc_val", ""),
            1 if r.get("step3") else 0, r.get("step3_fac", ""), r.get("step3_date", ""), r.get("fac_val", ""),
            1 if r.get("step4") else 0, r.get("step4_date", ""), r.get("notes", ""), r.get("updated_by", "AutoSync"), now_iso
        ))
        synced_count += 1

    conn.commit()
    conn.close()
    return {"success": True, "synced": synced_count}

# Dynamic Excel Export
@app.get("/api/export/excel")
def export_excel(month: Optional[str] = "Octubre"):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    if month and month != "ALL":
        cursor.execute("SELECT * FROM billing_records WHERE month = ? ORDER BY client_id ASC, record_id ASC", (month,))
    else:
        cursor.execute("SELECT * FROM billing_records ORDER BY client_id ASC, record_id ASC")
    rows = cursor.fetchall()

    cursor.execute("SELECT * FROM clients")
    c_rows = cursor.fetchall()
    conn.close()

    client_dict = {c["id"]: dict(c) for c in c_rows}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Control_{month or 'General'}"

    NAVY = "1E3A8A"
    STEP1_COL = "0284C7"
    STEP2_COL = "D97706"
    STEP3_COL = "059669"
    STEP4_COL = "7C3AED"

    ws.merge_cells("A1:W1")
    ws["A1"] = f"CONTROL OPERATIVO DE FACTURACIÓN E INFORMES - {month or 'GENERAL'}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        ("ID", NAVY), ("CLIENTE", NAVY), ("RESPONSABLE", NAVY), ("PROGRAMACIÓN BASE", NAVY),
        ("TIPO CORTE", NAVY), ("MES", NAVY), ("PERÍODO / DETALLE", NAVY),
        ("P1: REQ. INFORME", STEP1_COL), ("P1: FECHA ENVÍO", STEP1_COL), ("P1: ¿ENVIADO?", STEP1_COL),
        ("P1: Nº PRE-FACTURA", STEP1_COL), ("P1: VALOR PRE-FACTURA", STEP1_COL),
        ("P2: REQ. OC", STEP2_COL), ("P2: Nº OC / HES", STEP2_COL), ("P2: FECHA OC", STEP2_COL), ("P2: ¿OC RECIBIDA?", STEP2_COL),
        ("P3: FECHA FACTURA", STEP3_COL), ("P3: Nº FACTURA", STEP3_COL), ("P3: VALOR FACTURA", STEP3_COL), ("P3: ¿GENERADA?", STEP3_COL),
        ("COMPARATIVA / AUDITORÍA", "D97706"),
        ("P4: CANAL ENTREGA", STEP4_COL), ("P4: ¿ENTREGADA?", STEP4_COL),
        ("ESTADO ACTUAL", "475569")
    ]

    for col_idx, (text, color) in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=text)
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    border_thin = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                         top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

    for r_idx, r in enumerate(rows, start=4):
        c_meta = client_dict.get(r["client_id"], {})
        s1 = "SÍ" if r["step1"] else ("N/A" if not c_meta.get("req_inf") else "NO")
        s2 = "SÍ" if r["step2"] else ("N/A" if not c_meta.get("req_oc") else "NO")
        s3 = "SÍ" if r["step3"] else "NO"
        s4 = "SÍ" if r["step4"] else "NO"

        # Math Difference
        p_val_str = str(r["pref_val"] or "").replace("$", "").replace(".", "").replace(",", "").strip()
        f_val_str = str(r["fac_val"] or "").replace("$", "").replace(".", "").replace(",", "").strip()
        try:
            p_val = float(p_val_str) if p_val_str else 0.0
            f_val = float(f_val_str) if f_val_str else 0.0
            diff = f_val - p_val
            diff_text = "✓ Coincide" if (p_val > 0 and f_val > 0 and diff == 0) else (f"Dif: ${diff:,.0f}" if (p_val > 0 or f_val > 0) else "")
        except Exception:
            diff_text = ""

        # Status
        if r["step4"]:
            st_text = "🟢 COMPLETADO"
        elif r["step3"]:
            st_text = "🔵 FACTURADO (Pend. Envío)"
        elif (r["step2"] or not c_meta.get("req_oc")) and (r["step1"] or not c_meta.get("req_inf")):
            st_text = "🟠 APROBADO (Pend. Facturar)"
        elif r["step1"]:
            st_text = "🟡 ESPERANDO OC/APROB."
        else:
            st_text = "⚪ PENDIENTE INFORME"

        row_vals = [
            r["client_id"], r["client_name"], c_meta.get("resp", ""), c_meta.get("prog", ""),
            r["freq_type"], r["month"], r["period_detail"],
            "SÍ" if c_meta.get("req_inf") else "NO", r["step1_date"], s1,
            r["pref_num"], r["pref_val"],
            "SÍ" if c_meta.get("req_oc") else "NO", r["step2_oc"], r["step2_date"], s2,
            r["step3_date"], r["step3_fac"], r["fac_val"], s3,
            diff_text,
            c_meta.get("canal", ""), s4,
            st_text
        ]

        fill_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_idx, column=col_idx, value=val)
            cell.border = border_thin
            cell.font = Font(name="Calibri", size=10)
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            if col_idx in [1, 3, 4, 5, 6, 7, 8, 10, 13, 16, 20, 23]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Control_Facturacion_{month}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
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
    print("Iniciando servidor de facturación en http://0.0.0.0:8000 ...")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)
