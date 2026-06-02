"""
Script para generar/actualizar la base de datos SQLite desde el Excel planilla_dov.xls
Uso: python3 crear_db.py
"""
import re
import xlrd
import sqlite3

def excel_date(val, wb):
    if not val:
        return ''
    try:
        f = float(val)
        if f > 1000:
            return xlrd.xldate_as_datetime(f, wb.datemode).strftime('%d/%m/%Y')
    except:
        pass
    s = str(val).strip()
    if s in ('0.0', '0'):
        return ''
    # Normaliza fechas escritas como texto (ej. "20.03.2026" o "20-03-2026") a DD/MM/YYYY
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$', s)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return s

def clean(val):
    if val == '' or val is None:
        return ''
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return str(f)
    except:
        return str(val).strip()

def clean_hora(val):
    if not val:
        return ''
    try:
        f = float(val)
        total_min = round(f * 24 * 60)
        h = total_min // 60
        m = total_min % 60
        return f"{h:02d}:{m:02d}"
    except:
        return str(val).strip()

wb = xlrd.open_workbook('planilla_dov.xls')
sheet = wb.sheet_by_name('REGISTRO_GENERAL')

conn = sqlite3.connect('reportes.db')
c = conn.cursor()
c.execute('DROP TABLE IF EXISTS registro_general')
c.execute('''CREATE TABLE registro_general (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ods TEXT, ubicacion_tecnica TEXT, sem TEXT, bim TEXT, anio TEXT,
    fecha_generacion TEXT, rcc TEXT, primera_lect TEXT, aud TEXT,
    consumo_cero TEXT, integral TEXT, distrito TEXT, analisis TEXT,
    repite TEXT, fecha_entrega TEXT, inspector TEXT, exp_ref TEXT,
    nro_serie TEXT, material TEXT, lectura TEXT, fecha_inspeccion TEXT,
    hora_llegada TEXT, hora_salida TEXT, funcionamiento TEXT, resid TEXT,
    actividad TEXT, corte TEXT, precinto TEXT, pozo TEXT,
    valvula_retencion TEXT, estado_conexion TEXT, tipo_caja TEXT,
    movil TEXT, cx_libres TEXT, observaciones TEXT, resuelve TEXT,
    motivo TEXT, ods_inconsistencia_rcc TEXT, zrec TEXT,
    ods_recambio TEXT, alto_consumo TEXT, fecha_comunicacion TEXT,
    nro_contacto TEXT
)''')

for i in range(1, sheet.nrows):
    r = sheet.row_values(i)
    c.execute('''INSERT INTO registro_general VALUES
        (NULL,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
        clean(r[0]), clean(r[1]), clean(r[2]), clean(r[3]), clean(r[4]),
        excel_date(r[5], wb), clean(r[6]), clean(r[7]), clean(r[8]), clean(r[9]),
        clean(r[10]), clean(r[11]), clean(r[12]), clean(r[13]),
        excel_date(r[14], wb), str(r[15]).strip(), clean(r[16]), clean(r[17]),
        str(r[18]).strip(), clean(r[19]), excel_date(r[20], wb),
        clean_hora(r[21]), clean_hora(r[22]), str(r[23]).strip(),
        str(r[24]).strip(), str(r[25]).strip(), str(r[26]).strip(),
        str(r[27]).strip(), str(r[28]).strip(), str(r[29]).strip(),
        str(r[30]).strip(), str(r[31]).strip(), str(r[32]).strip(),
        str(r[33]).strip(), str(r[34]).strip(), str(r[35]).strip(),
        str(r[36]).strip(), str(r[37]).strip(), str(r[38]).strip(),
        str(r[39]).strip(), str(r[40]).strip(), excel_date(r[41], wb),
        str(r[42]).strip()
    ))

conn.commit()
c.execute('SELECT COUNT(*) FROM registro_general')
print(f"Base de datos creada. Registros: {c.fetchone()[0]}")
conn.close()
