from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import os
import sqlite3
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook

DATABASE_URL = os.environ.get("DATABASE_URL")

if DATABASE_URL:
    import psycopg2
    import psycopg2.extras

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "icmpc-inventory-secret-key")
DATABASE = "inventory.db"


def using_postgres():
    return bool(DATABASE_URL)


def get_db_connection():
    if using_postgres():
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)

    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def q(sql):
    if using_postgres():
        return sql.replace("?", "%s")
    return sql


def fetchone(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(q(sql), params)
    return cur.fetchone()


def fetchall(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(q(sql), params)
    return cur.fetchall()


def execute(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(q(sql), params)
    return cur


def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    if using_postgres():
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                full_name TEXT NOT NULL,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'staff',
                branch TEXT
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS suppliers (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                address TEXT,
                tin TEXT,
                vat_type TEXT,
                business_type TEXT,
                created_at TEXT
            )
        """)

        try:
            cur.execute("ALTER TABLE suppliers ADD COLUMN category TEXT")
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            cur.execute("ALTER TABLE suppliers ADD COLUMN contact_person TEXT")
            conn.commit()
        except Exception:
            conn.rollback()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS items (
                id SERIAL PRIMARY KEY,
                item_name TEXT NOT NULL,
                category TEXT,
                unit TEXT,
                quantity INTEGER NOT NULL DEFAULT 0,
                low_stock_limit INTEGER NOT NULL DEFAULT 5,
                supplier_id INTEGER REFERENCES suppliers(id),
                branch TEXT,
                created_at TEXT
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS stock_transactions (
                id SERIAL PRIMARY KEY,
                item_id INTEGER NOT NULL REFERENCES items(id),
                transaction_type TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                remarks TEXT,
                encoded_by TEXT,
                date_created TEXT NOT NULL
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS purchase_orders (
                id SERIAL PRIMARY KEY,
                po_number TEXT UNIQUE NOT NULL,
                supplier_id INTEGER REFERENCES suppliers(id),
                po_date TEXT NOT NULL,
                requested_by TEXT,
                status TEXT DEFAULT 'Pending',
                notes TEXT,
                created_at TEXT
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS purchase_order_items (
                id SERIAL PRIMARY KEY,
                po_id INTEGER NOT NULL REFERENCES purchase_orders(id),
                item_description TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                unit TEXT,
                unit_price NUMERIC DEFAULT 0
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS audit_logs (
                id SERIAL PRIMARY KEY,
                user_name TEXT,
                action TEXT NOT NULL,
                details TEXT,
                date_created TEXT NOT NULL
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS equipment_assignments (
                id SERIAL PRIMARY KEY,
                employee_name TEXT NOT NULL,
                position TEXT,
                branch TEXT,
                item_id INTEGER REFERENCES items(id),
                quantity INTEGER NOT NULL DEFAULT 1,
                remarks TEXT,
                assigned_by TEXT,
                date_assigned TEXT NOT NULL
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id SERIAL PRIMARY KEY,
                employee_name TEXT NOT NULL,
                position TEXT,
                branch TEXT,
                date_created TEXT
            )
        """)

    else:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'staff',
                branch TEXT
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                address TEXT,
                tin TEXT,
                vat_type TEXT,
                business_type TEXT,
                created_at TEXT
            )
        """)

        try:
            cur.execute("ALTER TABLE suppliers ADD COLUMN category TEXT")
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            cur.execute("ALTER TABLE suppliers ADD COLUMN contact_person TEXT")
            conn.commit()
        except Exception:
            conn.rollback()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_name TEXT NOT NULL,
                category TEXT,
                unit TEXT,
                quantity INTEGER NOT NULL DEFAULT 0,
                low_stock_limit INTEGER NOT NULL DEFAULT 5,
                supplier_id INTEGER,
                branch TEXT,
                created_at TEXT,
                FOREIGN KEY (supplier_id) REFERENCES suppliers(id)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS stock_transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                transaction_type TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                remarks TEXT,
                encoded_by TEXT,
                date_created TEXT NOT NULL,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS purchase_orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                po_number TEXT UNIQUE NOT NULL,
                supplier_id INTEGER,
                po_date TEXT NOT NULL,
                requested_by TEXT,
                status TEXT DEFAULT 'Pending',
                notes TEXT,
                created_at TEXT,
                FOREIGN KEY (supplier_id) REFERENCES suppliers(id)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS purchase_order_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                po_id INTEGER NOT NULL,
                item_description TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                unit TEXT,
                unit_price REAL DEFAULT 0,
                FOREIGN KEY (po_id) REFERENCES purchase_orders(id)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT,
                action TEXT NOT NULL,
                details TEXT,
                date_created TEXT NOT NULL
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS equipment_assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_name TEXT NOT NULL,
                position TEXT,
                branch TEXT,
                item_id INTEGER,
                quantity INTEGER NOT NULL DEFAULT 1,
                remarks TEXT,
                assigned_by TEXT,
                date_assigned TEXT NOT NULL,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_name TEXT NOT NULL,
                position TEXT,
                branch TEXT,
                date_created TEXT
            )
        """)

    # ALWAYS ENSURE ADMIN EXISTS
    cur.execute(q("DELETE FROM users WHERE username = ?"), ("admin",))

    cur.execute(q("""
        INSERT INTO users (full_name, username, password, role, branch)
        VALUES (?, ?, ?, ?, ?)
    """), ("Administrator", "admin", "admin123", "admin", "Head Office"))

    conn.commit()
    conn.close()


def logged_in():
    return "user_id" in session


def require_login():
    if not logged_in():
        return redirect(url_for("login"))
    return None


def require_admin():
    check = require_login()
    if check:
        return check

    if session.get("role") != "admin":
        flash("Admin access only.", "danger")
        return redirect(url_for("dashboard"))

    return None


def branch_where(alias="items"):
    role = session.get("role")
    branch = session.get("branch")

    if role == "branch_manager" and branch:
        return f" WHERE {alias}.branch = ? ", [branch]

    return "", []

def log_action(action, details=""):
    try:
        conn = get_db_connection()
        execute(conn, """
            INSERT INTO audit_logs (user_name, action, details, date_created)
            VALUES (?, ?, ?, ?)
        """, (
            session.get("full_name", "System"),
            action,
            details,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        conn.commit()
        conn.close()
    except Exception as e:
        print("Audit log error:", e)

@app.route("/")
def index():
    if logged_in():
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        conn = get_db_connection()
        user = fetchone(conn, """
            SELECT * FROM users
            WHERE username = ? AND password = ?
        """, (username, password))
        conn.close()

        if user:
            session["user_id"] = user["id"]
            session["full_name"] = user["full_name"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            session["branch"] = user["branch"]
            flash("Login successful.", "success")
            return redirect(url_for("dashboard"))

        flash("Invalid username or password.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have logged out.", "success")
    return redirect(url_for("login"))


@app.route("/dashboard")
def dashboard():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    where_sql, params = branch_where("items")

    total_items = fetchone(conn, f"SELECT COUNT(*) AS count FROM items {where_sql}", params)["count"]
    total_suppliers = fetchone(conn, "SELECT COUNT(*) AS count FROM suppliers")["count"]
    total_po = fetchone(conn, "SELECT COUNT(*) AS count FROM purchase_orders")["count"]
    total_transactions = fetchone(conn, "SELECT COUNT(*) AS count FROM stock_transactions")["count"]

    low_stock = fetchall(conn, f"""
        SELECT items.*, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        {where_sql}
        {"AND" if where_sql else "WHERE"} items.quantity <= items.low_stock_limit
        ORDER BY items.quantity ASC
        LIMIT 10
    """, params)

    recent_transactions = fetchall(conn, """
        SELECT stock_transactions.*, items.item_name
        FROM stock_transactions
        JOIN items ON items.id = stock_transactions.item_id
        ORDER BY stock_transactions.id DESC
        LIMIT 10
    """)

    category_rows = fetchall(conn, f"""
        SELECT COALESCE(category, 'Uncategorized') AS category, COUNT(*) AS total
        FROM items
        {where_sql}
        GROUP BY COALESCE(category, 'Uncategorized')
        ORDER BY total DESC
        LIMIT 8
    """, params)

    stock_rows = fetchall(conn, f"""
        SELECT item_name, quantity
        FROM items
        {where_sql}
        ORDER BY quantity DESC
        LIMIT 8
    """, params)

    conn.close()

    return render_template(
        "dashboard.html",
        total_items=total_items,
        total_suppliers=total_suppliers,
        total_po=total_po,
        total_transactions=total_transactions,
        low_stock=low_stock,
        recent_transactions=recent_transactions,
        category_labels=[r["category"] for r in category_rows],
        category_values=[r["total"] for r in category_rows],
        stock_labels=[r["item_name"] for r in stock_rows],
        stock_values=[r["quantity"] for r in stock_rows],
    )


@app.route("/users", methods=["GET", "POST"])
def users():
    check = require_admin()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        try:
            execute(conn, """
                INSERT INTO users (full_name, username, password, role, branch)
                VALUES (?, ?, ?, ?, ?)
            """, (
                request.form.get("full_name"),
                request.form.get("username"),
                request.form.get("password"),
                request.form.get("role"),
                request.form.get("branch")
            ))
            conn.commit()
            flash("User account created successfully.", "success")
        except Exception as e:
            conn.rollback()
            flash(f"Error creating user: {e}", "danger")

        conn.close()
        return redirect(url_for("users"))

    user_list = fetchall(conn, "SELECT * FROM users ORDER BY role, full_name")
    conn.close()
    return render_template("users.html", users=user_list)


@app.route("/users/delete/<int:user_id>")
def delete_user(user_id):
    check = require_admin()
    if check:
        return check

    if user_id == session.get("user_id"):
        flash("You cannot delete your own account.", "danger")
        return redirect(url_for("users"))

    conn = get_db_connection()
    execute(conn, "DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()

    flash("User deleted successfully.", "success")
    return redirect(url_for("users"))


@app.route("/suppliers", methods=["GET", "POST"])
def suppliers():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        execute(conn, """
            INSERT INTO suppliers 
            (name, email, phone, address, tin, vat_type, business_type, contact_person, category, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            request.form.get("name"),
            request.form.get("email"),
            request.form.get("phone"),
            request.form.get("address"),
            request.form.get("tin"),
            request.form.get("vat_type"),
            request.form.get("business_type"),
            request.form.get("contact_person"),
            request.form.get("category"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()

        flash("Supplier added successfully.", "success")
        return redirect(url_for("suppliers"))

    supplier_list = fetchall(conn, "SELECT * FROM suppliers ORDER BY name ASC")
    conn.close()

    return render_template("suppliers.html", suppliers=supplier_list)


@app.route("/items", methods=["GET", "POST"])
def items():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    # =============================
    # ADD ITEM
    # =============================
    if request.method == "POST":
        execute(conn, """
            INSERT INTO items
            (item_name, category, unit, quantity, low_stock_limit, supplier_id, branch, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            request.form.get("item_name"),
            request.form.get("category"),
            request.form.get("unit"),
            int(request.form.get("quantity") or 0),
            int(request.form.get("low_stock_limit") or 5),
            request.form.get("supplier_id") or None,
            request.form.get("branch") or session.get("branch"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        conn.commit()
        conn.close()

        flash("Item added successfully.", "success")
        return redirect(url_for("items"))

    # =============================
    # FILTER LOGIC (FIXED)
    # =============================
    search = request.args.get("search", "")
    category = request.args.get("category", "")

    query = """
        SELECT items.*, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        WHERE 1=1
    """

    params = []

    # branch filter
    role = session.get("role")
    branch = session.get("branch")

    if role == "branch_manager" and branch:
        query += " AND items.branch = ?"
        params.append(branch)

    # search filter
    if search:
        query += " AND items.item_name LIKE ?"
        params.append(f"%{search}%")

    # category filter
    if category:
        query += " AND items.category = ?"
        params.append(category)

    query += " ORDER BY items.item_name ASC"

    item_list = fetchall(conn, query, tuple(params))

    supplier_list = fetchall(conn, "SELECT * FROM suppliers ORDER BY name ASC")
    conn.close()

    return render_template(
        "items.html",
        items=item_list,
        suppliers=supplier_list
    )

# =============================
# DELETE ITEM
# =============================
@app.route("/items/delete/<int:item_id>", methods=["POST"])
def delete_item(item_id):
    check = require_login()
    if check:
        return check

    if session.get("role") != "admin":
        flash("Admin access only.", "danger")
        return redirect(url_for("items"))

    conn = get_db_connection()

    # delete related stock transactions (avoid FK issues)
    execute(conn, "DELETE FROM stock_transactions WHERE item_id = ?", (item_id,))

    # delete the item
    execute(conn, "DELETE FROM items WHERE id = ?", (item_id,))

    conn.commit()
    conn.close()

    flash("Item deleted successfully.", "success")
    return redirect(url_for("items"))


# =============================
# DELETE SUPPLIER
# =============================
@app.route("/suppliers/delete/<int:supplier_id>", methods=["POST"])
def delete_supplier(supplier_id):
    check = require_login()
    if check:
        return check

    if session.get("role") != "admin":
        flash("Admin access only.", "danger")
        return redirect(url_for("suppliers"))

    conn = get_db_connection()

    try:
        # Remove supplier reference from items
        execute(conn, """
            UPDATE items 
            SET supplier_id = NULL 
            WHERE supplier_id = ?
        """, (supplier_id,))

        # Remove supplier reference from purchase orders
        execute(conn, """
            UPDATE purchase_orders 
            SET supplier_id = NULL 
            WHERE supplier_id = ?
        """, (supplier_id,))

        # Delete supplier
        execute(conn, """
            DELETE FROM suppliers 
            WHERE id = ?
        """, (supplier_id,))

        conn.commit()
        flash("Supplier deleted successfully.", "success")

    except Exception as e:
        conn.rollback()
        flash(f"Error deleting supplier: {e}", "danger")

    finally:
        conn.close()

    return redirect(url_for("suppliers"))


@app.route("/stock-in", methods=["GET", "POST"])
def stock_in():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        item_id = int(request.form.get("item_id"))
        quantity = int(request.form.get("quantity") or 0)
        remarks = request.form.get("remarks")

        if quantity <= 0:
            flash("Quantity must be greater than zero.", "danger")
            conn.close()
            return redirect(url_for("stock_in"))

        execute(conn, "UPDATE items SET quantity = quantity + ? WHERE id = ?", (quantity, item_id))
        execute(conn, """
            INSERT INTO stock_transactions
            (item_id, transaction_type, quantity, remarks, encoded_by, date_created)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            item_id,
            "Stock In",
            quantity,
            remarks,
            session.get("full_name"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()

        flash("Stock added successfully.", "success")
        return redirect(url_for("stock_in"))

    where_sql, params = branch_where("items")
    item_list = fetchall(conn, f"SELECT * FROM items {where_sql} ORDER BY item_name ASC", params)
    conn.close()
    return render_template("stock_in.html", items=item_list)


@app.route("/stock-out", methods=["GET", "POST"])
def stock_out():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        item_id = int(request.form.get("item_id"))
        quantity = int(request.form.get("quantity") or 0)
        remarks = request.form.get("remarks")

        item = fetchone(conn, "SELECT * FROM items WHERE id = ?", (item_id,))

        if quantity <= 0:
            flash("Quantity must be greater than zero.", "danger")
            conn.close()
            return redirect(url_for("stock_out"))

        if not item or item["quantity"] < quantity:
            flash("Not enough stock available.", "danger")
            conn.close()
            return redirect(url_for("stock_out"))

        execute(conn, "UPDATE items SET quantity = quantity - ? WHERE id = ?", (quantity, item_id))
        execute(conn, """
            INSERT INTO stock_transactions
            (item_id, transaction_type, quantity, remarks, encoded_by, date_created)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            item_id,
            "Stock Out",
            quantity,
            remarks,
            session.get("full_name"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()

        flash("Stock released successfully.", "success")
        return redirect(url_for("stock_out"))

    where_sql, params = branch_where("items")
    item_list = fetchall(conn, f"SELECT * FROM items {where_sql} ORDER BY item_name ASC", params)
    conn.close()
    return render_template("stock_out.html", items=item_list)


@app.route("/transactions")
def transactions():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    rows = fetchall(conn, """
        SELECT stock_transactions.*, items.item_name, items.branch
        FROM stock_transactions
        JOIN items ON items.id = stock_transactions.item_id
        ORDER BY stock_transactions.id DESC
    """)
    conn.close()

    return render_template("transactions.html", transactions=rows)


@app.route("/purchase-orders", methods=["GET", "POST"])
def purchase_orders():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        po_number = "PO-" + datetime.now().strftime("%Y%m%d%H%M%S")
        supplier_id = request.form.get("supplier_id") or None
        po_date = request.form.get("po_date") or datetime.now().strftime("%Y-%m-%d")
        requested_by = request.form.get("requested_by")
        notes = request.form.get("notes")

        if using_postgres():
            cur = execute(conn, """
                INSERT INTO purchase_orders
                (po_number, supplier_id, po_date, requested_by, status, notes, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                RETURNING id
            """, (
                po_number,
                supplier_id,
                po_date,
                requested_by,
                "Pending",
                notes,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
            po_id = cur.fetchone()["id"]
        else:
            cur = execute(conn, """
                INSERT INTO purchase_orders
                (po_number, supplier_id, po_date, requested_by, status, notes, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                po_number,
                supplier_id,
                po_date,
                requested_by,
                "Pending",
                notes,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
            po_id = cur.lastrowid

        descriptions = request.form.getlist("description[]")
        quantities = request.form.getlist("po_quantity[]")
        units = request.form.getlist("po_unit[]")
        prices = request.form.getlist("unit_price[]")

        for desc, qty, unit, price in zip(descriptions, quantities, units, prices):
            if desc and desc.strip():
                execute(conn, """
                    INSERT INTO purchase_order_items
                    (po_id, item_description, quantity, unit, unit_price)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    po_id,
                    desc.strip(),
                    int(qty or 0),
                    unit,
                    float(price or 0)
                ))

        conn.commit()
        conn.close()

        flash("Purchase order created successfully.", "success")
        return redirect(url_for("purchase_orders"))

    po_list = fetchall(conn, """
        SELECT purchase_orders.*, suppliers.name AS supplier_name
        FROM purchase_orders
        LEFT JOIN suppliers ON suppliers.id = purchase_orders.supplier_id
        ORDER BY purchase_orders.id DESC
    """)

    supplier_list = fetchall(conn, "SELECT * FROM suppliers ORDER BY name ASC")
    conn.close()

    return render_template("purchase_orders.html", purchase_orders=po_list, suppliers=supplier_list)


@app.route("/purchase-orders/<int:po_id>/print")
def print_purchase_order(po_id):
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    po = fetchone(conn, """
        SELECT purchase_orders.*, suppliers.name AS supplier_name, suppliers.address,
               suppliers.tin, suppliers.phone, suppliers.email
        FROM purchase_orders
        LEFT JOIN suppliers ON suppliers.id = purchase_orders.supplier_id
        WHERE purchase_orders.id = ?
    """, (po_id,))

    po_items = fetchall(conn, """
        SELECT *, quantity * unit_price AS line_total
        FROM purchase_order_items
        WHERE po_id = ?
    """, (po_id,))
    conn.close()

    if not po:
        flash("Purchase order not found.", "danger")
        return redirect(url_for("purchase_orders"))

    grand_total = sum([item["line_total"] or 0 for item in po_items])

    return render_template(
        "purchase_order_print.html",
        po=po,
        po_items=po_items,
        grand_total=grand_total
    )


@app.route("/reports")
def reports():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    total_items = fetchone(conn, "SELECT COUNT(*) AS count FROM items")["count"]
    low_stock_count = fetchone(conn, "SELECT COUNT(*) AS count FROM items WHERE quantity <= low_stock_limit")["count"]
    stock_in_count = fetchone(conn, "SELECT COUNT(*) AS count FROM stock_transactions WHERE transaction_type = 'Stock In'")["count"]
    stock_out_count = fetchone(conn, "SELECT COUNT(*) AS count FROM stock_transactions WHERE transaction_type = 'Stock Out'")["count"]

    conn.close()

    return render_template(
        "reports.html",
        total_items=total_items,
        low_stock_count=low_stock_count,
        stock_in_count=stock_in_count,
        stock_out_count=stock_out_count
    )


@app.route("/export-items")
def export_items():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    search = request.args.get("search", "")
    category = request.args.get("category", "")

    query = """
        SELECT items.item_name, items.category, items.unit, items.quantity,
               items.low_stock_limit, items.branch, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        WHERE 1=1
    """

    params = []

    if search:
        query += " AND item_name LIKE ?"
        params.append(f"%{search}%")

    if category:
        query += " AND category = ?"
        params.append(category)

    query += " ORDER BY items.item_name ASC"

    rows = fetchall(conn, query, tuple(params))
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory Items"

    ws.append(["Iligan Cement Multi-Purpose Cooperative"])
    ws.append(["Filtered Inventory Report"])
    ws.append(["Generated:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])
    ws.append(["Item Name", "Category", "Unit", "Quantity", "Low Stock Limit", "Branch", "Supplier"])

    for row in rows:
        ws.append([
            row["item_name"],
            row["category"],
            row["unit"],
            row["quantity"],
            row["low_stock_limit"],
            row["branch"],
            row["supplier_name"] or ""
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="ICMPC_Filtered_Items.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/export-transactions")
def export_transactions():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    rows = fetchall(conn, """
        SELECT stock_transactions.*, items.item_name, items.branch
        FROM stock_transactions
        JOIN items ON items.id = stock_transactions.item_id
        ORDER BY stock_transactions.id DESC
    """)
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"

    ws.append(["Iligan Cement Multi-Purpose Cooperative"])
    ws.append(["Stock Transactions Report"])
    ws.append(["Generated:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])
    ws.append(["Date", "Item", "Branch", "Type", "Quantity", "Remarks", "Encoded By"])

    for row in rows:
        ws.append([
            row["date_created"],
            row["item_name"],
            row["branch"],
            row["transaction_type"],
            row["quantity"],
            row["remarks"],
            row["encoded_by"]
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="ICMPC_Stock_Transactions.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/audit-logs")
def audit_logs():
    check = require_admin()
    if check:
        return check

    conn = get_db_connection()
    logs = fetchall(conn, """
        SELECT *
        FROM audit_logs
        ORDER BY id DESC
        LIMIT 300
    """)
    conn.close()

    return render_template("audit_logs.html", logs=logs)


@app.route("/equipment-assignments", methods=["GET", "POST"])
def equipment_assignments():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        item_id = int(request.form.get("item_id"))
        quantity = int(request.form.get("quantity") or 1)

        item = fetchone(conn, "SELECT * FROM items WHERE id = ?", (item_id,))

        if not item or item["quantity"] < quantity:
            flash("Not enough stock available for assignment.", "danger")
            conn.close()
            return redirect(url_for("equipment_assignments"))

        execute(conn, """
            INSERT INTO equipment_assignments
            (employee_name, position, branch, item_id, quantity, remarks, assigned_by, date_assigned)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            request.form.get("employee_name"),
            request.form.get("position"),
            request.form.get("branch"),
            item_id,
            quantity,
            request.form.get("remarks"),
            session.get("full_name"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        execute(conn, "UPDATE items SET quantity = quantity - ? WHERE id = ?", (quantity, item_id))

        conn.commit()
        conn.close()

        log_action(
            "Equipment Assigned",
            f"{request.form.get('employee_name')} received item ID {item_id}, Qty {quantity}"
        )

        flash("Equipment assigned successfully.", "success")
        return redirect(url_for("equipment_assignments"))

    assignments = fetchall(conn, """
        SELECT equipment_assignments.*, items.item_name, items.unit
        FROM equipment_assignments
        LEFT JOIN items ON items.id = equipment_assignments.item_id
        ORDER BY equipment_assignments.id DESC
    """)

    items = fetchall(conn, "SELECT * FROM items ORDER BY item_name ASC")
    conn.close()

    return render_template("equipment_assignments.html", assignments=assignments, items=items)


@app.route("/reports/pdf/items")
def export_items_pdf():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    rows = fetchall(conn, """
        SELECT items.item_name, items.category, items.unit, items.quantity,
               items.low_stock_limit, items.branch, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        ORDER BY items.item_name ASC
    """)
    conn.close()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Iligan Cement Multi-Purpose Cooperative", styles["Title"]))
    elements.append(Paragraph("Inventory Items Report", styles["Heading2"]))
    elements.append(Paragraph("Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"), styles["Normal"]))
    elements.append(Spacer(1, 12))

    data = [["Item", "Category", "Unit", "Qty", "Limit", "Branch", "Supplier"]]

    for r in rows:
        data.append([
            r["item_name"] or "",
            r["category"] or "",
            r["unit"] or "",
            str(r["quantity"] or 0),
            str(r["low_stock_limit"] or 0),
            r["branch"] or "",
            r["supplier_name"] or ""
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#123b7a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="ICMPC_Inventory_Items_Report.pdf",
        mimetype="application/pdf"
    )


@app.route("/reports/pdf/assignments")
def export_assignments_pdf():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    rows = fetchall(conn, """
        SELECT equipment_assignments.*, items.item_name
        FROM equipment_assignments
        LEFT JOIN items ON items.id = equipment_assignments.item_id
        ORDER BY equipment_assignments.id DESC
    """)
    conn.close()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Iligan Cement Multi-Purpose Cooperative", styles["Title"]))
    elements.append(Paragraph("Equipment Assignment Report", styles["Heading2"]))
    elements.append(Paragraph("Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"), styles["Normal"]))
    elements.append(Spacer(1, 12))

    data = [["Date", "Employee", "Position", "Branch", "Item", "Qty", "Assigned By"]]

    for r in rows:
        data.append([
            r["date_assigned"] or "",
            r["employee_name"] or "",
            r["position"] or "",
            r["branch"] or "",
            r["item_name"] or "",
            str(r["quantity"] or 0),
            r["assigned_by"] or ""
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#123b7a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="ICMPC_Equipment_Assignment_Report.pdf",
        mimetype="application/pdf"
    )

@app.route("/items/edit/<int:item_id>", methods=["GET", "POST"])
def edit_item(item_id):
    if session.get("role") != "admin":
        return redirect(url_for("items"))

    conn = get_db_connection()

    if request.method == "POST":
        execute(conn, """
            UPDATE items
            SET item_name = ?, category = ?, quantity = ?, unit = ?, low_stock_limit = ?
            WHERE id = ?
        """, (
            request.form["item_name"],
            request.form["category"],
            request.form["quantity"],
            request.form["unit"],
            request.form["low_stock_limit"],
            item_id
        ))

        conn.commit()
        conn.close()

        flash("Item updated successfully!", "success")
        return redirect(url_for("items"))

    item = fetchone(conn, "SELECT * FROM items WHERE id = ?", (item_id,))
    conn.close()

    return render_template("edit_item.html", item=item)


@app.route("/suppliers/edit/<int:supplier_id>", methods=["GET", "POST"])
def edit_supplier(supplier_id):
    check = require_login()
    if check:
        return check

    if session.get("role") != "admin":
        return redirect(url_for("suppliers"))

    conn = get_db_connection()

    if request.method == "POST":
        execute(conn, """
            UPDATE suppliers
            SET name = ?, email = ?, phone = ?, address = ?, tin = ?, vat_type = ?, business_type = ?, contact_person = ?, category = ?
            WHERE id = ?
        """, (
            request.form.get("name"),
            request.form.get("email"),
            request.form.get("phone"),
            request.form.get("address"),
            request.form.get("tin"),
            request.form.get("vat_type"),
            request.form.get("business_type"),
            request.form.get("contact_person"),
            request.form.get("category"),
            supplier_id
        ))

        conn.commit()
        conn.close()

        flash("Supplier updated successfully.", "success")
        return redirect(url_for("suppliers"))

    supplier = fetchone(conn, "SELECT * FROM suppliers WHERE id = ?", (supplier_id,))
    conn.close()

    return render_template("edit_supplier.html", supplier=supplier)

@app.route("/employees", methods=["GET", "POST"])
def employees():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    # =============================
    # ADD EMPLOYEE (ADMIN ONLY)
    # =============================
    if request.method == "POST":
        if session.get("role") != "admin":
            flash("Admin only.", "danger")
            return redirect(url_for("employees"))

        execute(conn, """
            INSERT INTO employees (employee_name, position, branch, date_created)
            VALUES (?, ?, ?, ?)
        """, (
            request.form.get("employee_name"),
            request.form.get("position"),
            request.form.get("branch") or session.get("branch"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()

        flash("Employee added successfully.", "success")
        return redirect(url_for("employees"))

    employee_list = fetchall(conn, "SELECT * FROM employees ORDER BY employee_name ASC")
    conn.close()

    return render_template("employees.html", employees=employee_list)


@app.route("/health")
def health():
    return "OK - ICMPC Inventory System is running"


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
