import os
import re
import json
import secrets
import sqlite3
import unicodedata
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "usuarios_hb.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def generar_clave_temporal():
    return "HB-" + secrets.token_urlsafe(6)


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            clave_temporal TEXT DEFAULT '',
            debe_cambiar_clave INTEGER DEFAULT 1,
            rol TEXT NOT NULL CHECK(rol IN ('admin', 'supervisor', 'asesor')),
            asesor_abako TEXT DEFAULT '',
            grupos_json TEXT DEFAULT '[]',
            asesores_json TEXT DEFAULT '[]',
            activo INTEGER DEFAULT 1
        )
    """)
    # Migración segura si ya existía la tabla v2.0
    cols_existentes = {r["name"] for r in cur.execute("PRAGMA table_info(usuarios)").fetchall()}
    if "clave_temporal" not in cols_existentes:
        cur.execute("ALTER TABLE usuarios ADD COLUMN clave_temporal TEXT DEFAULT ''")
    if "debe_cambiar_clave" not in cols_existentes:
        cur.execute("ALTER TABLE usuarios ADD COLUMN debe_cambiar_clave INTEGER DEFAULT 0")

    cur.execute("SELECT id FROM usuarios WHERE username = ?", ("admin",))
    if not cur.fetchone():
        clave_admin = generar_clave_temporal()
        cur.execute(
            """
            INSERT INTO usuarios (username, nombre, password_hash, clave_temporal, debe_cambiar_clave, rol)
            VALUES (?, ?, ?, ?, 1, 'admin')
            """,
            ("admin", "Administrador HB", generate_password_hash(clave_admin), clave_admin),
        )
        print("=" * 55)
        print(f"USUARIO ADMIN CREADO -> Usuario: admin | Clave temporal: {clave_admin}")
        print("=" * 55)
    conn.commit()
    conn.close()


def obtener_usuario_por_id(uid):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM usuarios WHERE id = ? AND activo = 1", (uid,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "nombre": row["nombre"],
        "rol": row["rol"],
        "debe_cambiar_clave": bool(row["debe_cambiar_clave"]),
        "asesor_abako": row["asesor_abako"] or "",
        "grupos": json.loads(row["grupos_json"] or "[]"),
        "asesores": json.loads(row["asesores_json"] or "[]"),
    }


def autenticar_usuario(username, password):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM usuarios WHERE LOWER(username) = LOWER(?) AND activo = 1",
        (username.strip(),),
    )
    row = cur.fetchone()
    conn.close()
    if row and check_password_hash(row["password_hash"], password):
        return obtener_usuario_por_id(row["id"])
    return None


def actualizar_contrasena_propia(uid, nueva_clave):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE usuarios
        SET password_hash = ?, clave_temporal = '', debe_cambiar_clave = 0
        WHERE id = ?
        """,
        (generate_password_hash(nueva_clave), uid),
    )
    conn.commit()
    conn.close()


def listar_usuarios():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, username, nombre, rol, clave_temporal, debe_cambiar_clave,
               asesor_abako, grupos_json, asesores_json, activo
        FROM usuarios ORDER BY rol, nombre
    """)
    rows = cur.fetchall()
    conn.close()
    return [
        {
            "id": r["id"],
            "username": r["username"],
            "nombre": r["nombre"],
            "rol": r["rol"],
            "clave_temporal": r["clave_temporal"] or "Ya cambiada por usuario",
            "debe_cambiar_clave": bool(r["debe_cambiar_clave"]),
            "asesor_abako": r["asesor_abako"] or "",
            "grupos": json.loads(r["grupos_json"] or "[]"),
            "asesores": json.loads(r["asesores_json"] or "[]"),
            "activo": bool(r["activo"]),
        }
        for r in rows
    ]


def crear_o_actualizar_usuario(data):
    conn = get_conn()
    cur = conn.cursor()
    uid = data.get("id")
    username = data.get("username", "").strip()
    nombre = data.get("nombre", "").strip()
    rol = data.get("rol", "asesor")
    password = data.get("password", "").strip()
    activo = 1 if data.get("activo", True) else 0
    asesor_abako = data.get("asesor_abako", "").strip()
    grupos_json = json.dumps(data.get("grupos", []), ensure_ascii=False)
    asesores_json = json.dumps(data.get("asesores", []), ensure_ascii=False)

    if uid:
        if password:
            cur.execute(
                """
                UPDATE usuarios
                SET username=?, nombre=?, password_hash=?, clave_temporal=?, debe_cambiar_clave=1,
                    rol=?, asesor_abako=?, grupos_json=?, asesores_json=?, activo=?
                WHERE id=?
                """,
                (username, nombre, generate_password_hash(password), password, rol, asesor_abako, grupos_json, asesores_json, activo, uid),
            )
        else:
            cur.execute(
                """
                UPDATE usuarios
                SET username=?, nombre=?, rol=?, asesor_abako=?, grupos_json=?, asesores_json=?, activo=?
                WHERE id=?
                """,
                (username, nombre, rol, asesor_abako, grupos_json, asesores_json, activo, uid),
            )
    else:
        clave_gen = password if password else generar_clave_temporal()
        cur.execute(
            """
            INSERT INTO usuarios (username, nombre, password_hash, clave_temporal, debe_cambiar_clave,
                                  rol, asesor_abako, grupos_json, asesores_json, activo)
            VALUES (?, ?, ?, ?, 1, ?, ?, ?, ?, ?)
            """,
            (username, nombre, generate_password_hash(clave_gen), clave_gen, rol, asesor_abako, grupos_json, asesores_json, activo),
        )
    conn.commit()
    conn.close()


def eliminar_usuario(uid):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM usuarios WHERE id=? AND username != 'admin'", (uid,))
    conn.commit()
    conn.close()


def limpiar_slug(texto):
    s = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-zA-Z0-9]+", ".", s.lower()).strip(".")


def generar_usuarios_masivos_asesores(lista_asesores_abako):
    conn = get_conn()
    cur = conn.cursor()
    creados = 0

    for ase_raw in lista_asesores_abako:
        ase = re.sub(r"\s+", " ", str(ase_raw or "")).strip().upper()
        if not ase or ase == "SIN ASIGNAR":
            continue

        cur.execute("SELECT id FROM usuarios WHERE UPPER(asesor_abako) = ?", (ase,))
        if cur.fetchone():
            continue

        partes = ase.split(" ")
        codigo = partes[-1] if partes[-1].isdigit() else ""
        primer_nombre = limpiar_slug(partes[0]) if len(partes) > 0 else "asesor"
        apellido = limpiar_slug(partes[-2]) if len(partes) > 2 else (limpiar_slug(partes[1]) if len(partes) > 1 else "")

        base_user = f"{primer_nombre}.{apellido}" if apellido else f"asesor.{codigo or primer_nombre}"
        username = base_user
        contador = 2
        while True:
            cur.execute("SELECT id FROM usuarios WHERE LOWER(username) = LOWER(?)", (username,))
            if not cur.fetchone():
                break
            username = f"{base_user}{contador}"
            contador += 1

        nombre_bonito = " ".join(p.capitalize() for p in partes)
        clave_unica = generar_clave_temporal()
        cur.execute(
            """
            INSERT INTO usuarios (username, nombre, password_hash, clave_temporal, debe_cambiar_clave,
                                  rol, asesor_abako, grupos_json, asesores_json, activo)
            VALUES (?, ?, ?, ?, 1, 'asesor', ?, '[]', '[]', 1)
            """,
            (username, nombre_bonito, generate_password_hash(clave_unica), clave_unica, ase),
        )
        creados += 1

    conn.commit()
    conn.close()
    return creados