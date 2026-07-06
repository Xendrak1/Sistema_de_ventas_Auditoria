import os
import sqlite3
from datetime import datetime

from flask import Flask, g, redirect, render_template, render_template_string, request, session, url_for


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "ventas.db")

app = Flask(__name__)

# Vulnerabilidad intencional: clave de sesion debil y fija para fines academicos.
app.config["SECRET_KEY"] = "ventas123"


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            rol TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS productos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sku TEXT,
            categoria TEXT,
            nombre TEXT NOT NULL,
            descripcion TEXT NOT NULL,
            precio REAL NOT NULL,
            stock INTEGER NOT NULL,
            imagen TEXT
        );

        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            telefono TEXT NOT NULL,
            correo TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS ventas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            producto_id INTEGER NOT NULL,
            cantidad INTEGER NOT NULL,
            total REAL NOT NULL,
            fecha TEXT NOT NULL,
            FOREIGN KEY(cliente_id) REFERENCES clientes(id),
            FOREIGN KEY(producto_id) REFERENCES productos(id)
        );
        """
    )
    ensure_product_columns(db)

    user_count = db.execute("SELECT COUNT(*) AS total FROM usuarios").fetchone()["total"]
    if user_count == 0:
        # Vulnerabilidad intencional: contrasenas en texto plano y credenciales debiles.
        db.executemany(
            "INSERT INTO usuarios (usuario, password, rol) VALUES (?, ?, ?)",
            [
                ("admin", "admin123", "Administrador"),
                ("ventas", "ventas123", "Vendedor"),
            ],
        )

    product_count = db.execute("SELECT COUNT(*) AS total FROM productos").fetchone()["total"]
    if product_count == 0:
        db.executemany(
            """
            INSERT INTO productos (sku, categoria, nombre, descripcion, precio, stock, imagen)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            [
                ("TEC-001", "Computacion", "Laptop oficina", "Equipo para facturacion, reportes y administracion.", 4200, 8, "img/laptop.png"),
                ("POS-002", "Punto de venta", "Impresora termica", "Impresora para comprobantes y tickets de venta.", 850, 12, "img/impresora-termica.png"),
                ("POS-003", "Inventario", "Lector codigo barras", "Lector USB para control de inventario y caja.", 320, 20, "img/scanner-barras.png"),
                ("POS-004", "Caja", "Cajon de dinero", "Cajon metalico para efectivo y monedas.", 690, 5, "img/cajon-dinero.png"),
                ("POS-005", "Punto de venta", "Monitor tactil POS", "Pantalla tactil compacta para mostrador.", 1850, 4, "img/monitor-pos.png"),
                ("POS-006", "Pagos", "Terminal de pago", "Equipo inalambrico para pagos con tarjeta.", 1150, 7, "img/terminal-pago.png"),
            ],
        )
    else:
        refresh_sample_products(db)

    client_count = db.execute("SELECT COUNT(*) AS total FROM clientes").fetchone()["total"]
    if client_count == 0:
        db.executemany(
            "INSERT INTO clientes (nombre, telefono, correo) VALUES (?, ?, ?)",
            [
                ("Cliente General", "70000000", "general@example.com"),
                ("Comercial Andina", "71111111", "ventas@andina.example"),
            ],
        )
    db.commit()


def ensure_product_columns(db):
    columns = [row["name"] for row in db.execute("PRAGMA table_info(productos)").fetchall()]
    migrations = {
        "sku": "ALTER TABLE productos ADD COLUMN sku TEXT",
        "categoria": "ALTER TABLE productos ADD COLUMN categoria TEXT",
        "imagen": "ALTER TABLE productos ADD COLUMN imagen TEXT",
    }
    for column, statement in migrations.items():
        if column not in columns:
            db.execute(statement)


def refresh_sample_products(db):
    updates = [
        ("TEC-001", "Computacion", "img/laptop.png", "Laptop oficina"),
        ("POS-002", "Punto de venta", "img/impresora-termica.png", "Impresora termica"),
        ("POS-003", "Inventario", "img/scanner-barras.png", "Lector codigo barras"),
    ]
    for sku, categoria, imagen, nombre in updates:
        db.execute(
            "UPDATE productos SET sku = COALESCE(sku, ?), categoria = COALESCE(categoria, ?), imagen = COALESCE(imagen, ?) WHERE nombre = ?",
            (sku, categoria, imagen, nombre),
        )

    extra_products = [
        ("POS-004", "Caja", "Cajon de dinero", "Cajon metalico para efectivo y monedas.", 690, 5, "img/cajon-dinero.png"),
        ("POS-005", "Punto de venta", "Monitor tactil POS", "Pantalla tactil compacta para mostrador.", 1850, 4, "img/monitor-pos.png"),
        ("POS-006", "Pagos", "Terminal de pago", "Equipo inalambrico para pagos con tarjeta.", 1150, 7, "img/terminal-pago.png"),
    ]
    for product in extra_products:
        exists = db.execute("SELECT id FROM productos WHERE sku = ?", (product[0],)).fetchone()
        if not exists:
            db.execute(
                """
                INSERT INTO productos (sku, categoria, nombre, descripcion, precio, stock, imagen)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                product,
            )


@app.before_request
def ensure_database():
    init_db()


def login_required():
    return "usuario" in session


@app.route("/")
def index():
    if login_required():
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario = request.form.get("usuario", "")
        password = request.form.get("password", "")

        # Vulnerabilidad intencional: consulta SQL concatenada.
        query = (
            "SELECT * FROM usuarios WHERE usuario = '"
            + usuario
            + "' AND password = '"
            + password
            + "'"
        )
        user = get_db().execute(query).fetchone()

        if user:
            session["usuario"] = user["usuario"]
            session["rol"] = user["rol"]
            return redirect(url_for("dashboard"))
        error = "Credenciales incorrectas"

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
def dashboard():
    if not login_required():
        return redirect(url_for("login"))

    db = get_db()
    stats = {
        "productos": db.execute("SELECT COUNT(*) AS total FROM productos").fetchone()["total"],
        "clientes": db.execute("SELECT COUNT(*) AS total FROM clientes").fetchone()["total"],
        "ventas": db.execute("SELECT COUNT(*) AS total FROM ventas").fetchone()["total"],
        "ingresos": db.execute("SELECT COALESCE(SUM(total), 0) AS total FROM ventas").fetchone()["total"],
    }
    ultimas_ventas = db.execute(
        """
        SELECT v.id, c.nombre AS cliente, p.nombre AS producto, v.cantidad, v.total, v.fecha
        FROM ventas v
        JOIN clientes c ON c.id = v.cliente_id
        JOIN productos p ON p.id = v.producto_id
        ORDER BY v.id DESC
        LIMIT 5
        """
    ).fetchall()
    bajo_stock = db.execute("SELECT * FROM productos WHERE stock <= 5 ORDER BY stock ASC").fetchall()
    destacados = db.execute("SELECT * FROM productos ORDER BY id DESC LIMIT 3").fetchall()
    return render_template(
        "dashboard.html",
        stats=stats,
        ventas=ultimas_ventas,
        bajo_stock=bajo_stock,
        destacados=destacados,
    )


@app.route("/productos")
def productos():
    if not login_required():
        return redirect(url_for("login"))
    rows = get_db().execute("SELECT * FROM productos ORDER BY id DESC").fetchall()
    return render_template("productos.html", productos=rows)


@app.route("/productos/nuevo", methods=["GET", "POST"])
def nuevo_producto():
    if not login_required():
        return redirect(url_for("login"))

    if request.method == "POST":
        nombre = request.form.get("nombre", "")
        descripcion = request.form.get("descripcion", "")
        precio = request.form.get("precio", "0")
        stock = request.form.get("stock", "0")
        sku = request.form.get("sku", "")
        categoria = request.form.get("categoria", "")
        imagen = request.form.get("imagen", "")

        get_db().execute(
            "INSERT INTO productos (sku, categoria, nombre, descripcion, precio, stock, imagen) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (sku, categoria, nombre, descripcion, precio, stock, imagen),
        )
        get_db().commit()
        return redirect(url_for("productos"))

    return render_template("producto_form.html")


@app.route("/clientes", methods=["GET", "POST"])
def clientes():
    if not login_required():
        return redirect(url_for("login"))

    db = get_db()
    if request.method == "POST":
        db.execute(
            "INSERT INTO clientes (nombre, telefono, correo) VALUES (?, ?, ?)",
            (
                request.form.get("nombre", ""),
                request.form.get("telefono", ""),
                request.form.get("correo", ""),
            ),
        )
        db.commit()
        return redirect(url_for("clientes"))

    rows = db.execute("SELECT * FROM clientes ORDER BY id DESC").fetchall()
    return render_template("clientes.html", clientes=rows)


@app.route("/ventas", methods=["GET", "POST"])
def ventas():
    if not login_required():
        return redirect(url_for("login"))

    db = get_db()
    if request.method == "POST":
        cliente_id = int(request.form.get("cliente_id", "0"))
        producto_id = int(request.form.get("producto_id", "0"))
        cantidad = int(request.form.get("cantidad", "1"))
        producto = db.execute("SELECT * FROM productos WHERE id = ?", (producto_id,)).fetchone()
        if producto:
            total = float(producto["precio"]) * cantidad
            db.execute(
                "INSERT INTO ventas (cliente_id, producto_id, cantidad, total, fecha) VALUES (?, ?, ?, ?, ?)",
                (cliente_id, producto_id, cantidad, total, datetime.now().strftime("%Y-%m-%d %H:%M")),
            )
            db.execute("UPDATE productos SET stock = stock - ? WHERE id = ?", (cantidad, producto_id))
            db.commit()
        return redirect(url_for("ventas"))

    rows = db.execute(
        """
        SELECT v.id, c.nombre AS cliente, p.nombre AS producto, v.cantidad, v.total, v.fecha
        FROM ventas v
        JOIN clientes c ON c.id = v.cliente_id
        JOIN productos p ON p.id = v.producto_id
        ORDER BY v.id DESC
        """
    ).fetchall()
    clientes_rows = db.execute("SELECT * FROM clientes ORDER BY nombre").fetchall()
    productos_rows = db.execute("SELECT * FROM productos ORDER BY nombre").fetchall()
    return render_template(
        "ventas.html",
        ventas=rows,
        clientes=clientes_rows,
        productos=productos_rows,
    )


@app.route("/reportes")
def reportes():
    if not login_required():
        return redirect(url_for("login"))

    db = get_db()
    ventas_por_producto = db.execute(
        """
        SELECT p.nombre, p.categoria, COALESCE(SUM(v.cantidad), 0) AS unidades,
               COALESCE(SUM(v.total), 0) AS total
        FROM productos p
        LEFT JOIN ventas v ON v.producto_id = p.id
        GROUP BY p.id
        ORDER BY total DESC, p.nombre
        """
    ).fetchall()
    inventario = db.execute("SELECT * FROM productos ORDER BY stock ASC").fetchall()
    return render_template("reportes.html", ventas_por_producto=ventas_por_producto, inventario=inventario)


@app.route("/buscar")
def buscar():
    if not login_required():
        return redirect(url_for("login"))

    q = request.args.get("q", "")

    # Vulnerabilidad intencional: SQL injection en busqueda.
    query = "SELECT * FROM productos WHERE nombre LIKE '%" + q + "%' OR descripcion LIKE '%" + q + "%'"
    productos_rows = get_db().execute(query).fetchall()

    # Vulnerabilidad intencional: XSS reflejado al renderizar el valor buscado.
    html = """
    {% extends "base.html" %}
    {% block content %}
    <section class="toolbar">
      <h1>Busqueda de productos</h1>
      <form action="/buscar" method="get" class="inline-form">
        <input name="q" value="{{ q }}" placeholder="Buscar producto">
        <button type="submit">Buscar</button>
      </form>
    </section>
    <p class="muted">Resultados para: """ + q + """</p>
    <div class="table-wrap">
      <table>
        <thead>
          <tr><th>ID</th><th>SKU</th><th>Nombre</th><th>Categoria</th><th>Descripcion</th><th>Precio</th><th>Stock</th></tr>
        </thead>
        <tbody>
        {% for producto in productos %}
          <tr>
            <td>{{ producto.id }}</td>
            <td>{{ producto.sku }}</td>
            <td>{{ producto.nombre }}</td>
            <td>{{ producto.categoria }}</td>
            <td>{{ producto.descripcion|safe }}</td>
            <td>{{ "%.2f"|format(producto.precio) }}</td>
            <td>{{ producto.stock }}</td>
          </tr>
        {% endfor %}
        </tbody>
      </table>
    </div>
    {% endblock %}
    """
    return render_template_string(html, q=q, productos=productos_rows)


if __name__ == "__main__":
    # Vulnerabilidad intencional: modo debug activo para entorno de laboratorio.
    app.run(host="0.0.0.0", port=5000, debug=True)
