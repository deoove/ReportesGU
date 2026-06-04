"""
Servidor de Reportes GU - Red Interna
Uso: python servidor.py
Acceso desde otras PCs: http://<IP-de-esta-PC>:8000
"""
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import sqlite3, os, re

app = Flask(__name__, static_folder='.')
CORS(app)

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reportes.db')

CAMPOS = [
    'ods','ubicacion_tecnica','sem','bim','anio','fecha_generacion','rcc',
    'primera_lect','aud','consumo_cero','integral','distrito','analisis','repite',
    'fecha_entrega','inspector','exp_ref','nro_serie','material','lectura',
    'fecha_inspeccion','hora_llegada','hora_salida','funcionamiento','resid',
    'actividad','corte','precinto','pozo','valvula_retencion','estado_conexion',
    'tipo_caja','movil','cx_libres','observaciones','resuelve','motivo',
    'ods_inconsistencia_rcc','zrec','ods_recambio','alto_consumo',
    'fecha_comunicacion','nro_contacto'
]

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def normalizar_fecha(val):
    if not val:
        return ''
    s = str(val).strip()
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$', s)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return s

# ── Servir archivos estáticos ─────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory('.', filename)

# ── API: obtener todos los registros ─────────────────────────────────────────
@app.route('/api/registros', methods=['GET'])
def get_registros():
    conn = get_db()
    rows = conn.execute('SELECT * FROM registro_general ORDER BY ods').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

# ── API: agregar múltiples registros (desde Excel) ───────────────────────────
@app.route('/api/registros', methods=['POST'])
def add_registros():
    data = request.get_json()
    rows = data.get('rows', [])
    overwrite = data.get('overwrite', [])  # lista de ODS a sobrescribir

    conn = get_db()
    agregados = 0
    actualizados = 0

    for row in rows:
        ods = str(row.get('ods', '')).strip()
        if not ods:
            continue
        row['fecha_generacion'] = normalizar_fecha(row.get('fecha_generacion',''))
        row['fecha_entrega']    = normalizar_fecha(row.get('fecha_entrega',''))
        row['fecha_inspeccion'] = normalizar_fecha(row.get('fecha_inspeccion',''))

        existe = conn.execute('SELECT id FROM registro_general WHERE ods=?', (ods,)).fetchone()

        if existe:
            if ods in overwrite:
                sets = ', '.join(f"{c}=?" for c in CAMPOS if c != 'ods')
                vals = [row.get(c, '') for c in CAMPOS if c != 'ods'] + [ods]
                conn.execute(f"UPDATE registro_general SET {sets} WHERE ods=?", vals)
                actualizados += 1
        else:
            cols = ', '.join(CAMPOS)
            placeholders = ', '.join(['?'] * len(CAMPOS))
            vals = [row.get(c, '') for c in CAMPOS]
            conn.execute(f"INSERT INTO registro_general ({cols}) VALUES ({placeholders})", vals)
            agregados += 1

    conn.commit()
    conn.close()
    return jsonify({'agregados': agregados, 'actualizados': actualizados})

# ── API: actualizar un registro ───────────────────────────────────────────────
@app.route('/api/registros/<ods>', methods=['PUT'])
def update_registro(ods):
    row = request.get_json()
    conn = get_db()
    existe = conn.execute('SELECT id FROM registro_general WHERE ods=?', (ods,)).fetchone()
    if not existe:
        conn.close()
        return jsonify({'error': 'ODS no encontrada'}), 404

    sets = ', '.join(f"{c}=?" for c in CAMPOS if c != 'ods')
    vals = [row.get(c, '') for c in CAMPOS if c != 'ods'] + [ods]
    conn.execute(f"UPDATE registro_general SET {sets} WHERE ods=?", vals)
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

# ── API: eliminar un registro ─────────────────────────────────────────────────
@app.route('/api/registros/<ods>', methods=['DELETE'])
def delete_registro(ods):
    conn = get_db()
    existe = conn.execute('SELECT id FROM registro_general WHERE ods=?', (ods,)).fetchone()
    if not existe:
        conn.close()
        return jsonify({'error': 'ODS no encontrada'}), 404
    conn.execute('DELETE FROM registro_general WHERE ods=?', (ods,))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

# ── Arranque ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    import socket
    hostname = socket.gethostname()
    ip = socket.gethostbyname(hostname)
    print("=" * 55)
    print("  Servidor Reportes GU iniciado")
    print("=" * 55)
    print(f"  Acceso local:   http://localhost:8000")
    print(f"  Acceso en red:  http://{ip}:8000")
    print("  (Compartí esa dirección con los otros usuarios)")
    print("=" * 55)
    print("  Para detener: Ctrl+C")
    print()
    app.run(host='0.0.0.0', port=8000, debug=False)
