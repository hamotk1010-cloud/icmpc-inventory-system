from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import os
import psycopg2
import sqlite3
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "icmpc-inventory-secret-key"
DATABASE = "inventory.db"


DATABASE_URL = os.environ.get("DATABASE_URL")

def get_db_connection():
    if DATABASE_URL:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    else:
        conn = sqlite3.connect(DATABASE)
        conn.row_factory = sqlite3.Row
        return conn


def column_exists(cur, table, column):
    cur.execute(f"PRAGMA table_info({table})")
    return column in [row[1] for row in cur.fetchall()]


def add_column_if_missing(cur, table, column, definition):
    if not column_exists(cur, table, column):
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

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

    cur.execute("SELECT id FROM users WHERE username = ?", ("admin",))
    if not cur.fetchone():
        cur.execute("""
            INSERT INTO users (full_name, username, password, role, branch)
            VALUES (?, ?, ?, ?, ?)
        """, ("Administrator", "admin", "admin123", "admin", "Head Office"))

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

if DATABASE_URL:
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM users
        WHERE username = %s AND password = %s
    """, (username, password))
    user = cur.fetchone()
else:
    user = conn.execute("""
        SELECT * FROM users
        WHERE username = ? AND password = ?
    """, (username, password)).fetchone()
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

    total_items = conn.execute(f"SELECT COUNT(*) FROM items {where_sql}", params).fetchone()[0]
    total_suppliers = conn.execute("SELECT COUNT(*) FROM suppliers").fetchone()[0]
    total_po = conn.execute("SELECT COUNT(*) FROM purchase_orders").fetchone()[0]
    total_transactions = conn.execute("SELECT COUNT(*) FROM stock_transactions").fetchone()[0]

    low_stock = conn.execute(f"""
        SELECT items.*, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        {where_sql}
        {"AND" if where_sql else "WHERE"} items.quantity <= items.low_stock_limit
        ORDER BY items.quantity ASC
        LIMIT 10
    """, params).fetchall()

    recent_transactions = conn.execute("""
        SELECT stock_transactions.*, items.item_name
        FROM stock_transactions
        JOIN items ON items.id = stock_transactions.item_id
        ORDER BY stock_transactions.id DESC
        LIMIT 10
    """).fetchall()

    category_rows = conn.execute(f"""
        SELECT COALESCE(category, 'Uncategorized') AS category, COUNT(*) AS total
        FROM items
        {where_sql}
        GROUP BY COALESCE(category, 'Uncategorized')
        ORDER BY total DESC
        LIMIT 8
    """, params).fetchall()

    stock_rows = conn.execute(f"""
        SELECT item_name, quantity
        FROM items
        {where_sql}
        ORDER BY quantity DESC
        LIMIT 8
    """, params).fetchall()

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
            conn.execute("""
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
        except sqlite3.IntegrityError:
            flash("Username already exists.", "danger")
        conn.close()
        return redirect(url_for("users"))

    user_list = conn.execute("SELECT * FROM users ORDER BY role, full_name").fetchall()
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
    conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
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
        conn.execute("""
            INSERT INTO suppliers (name, email, phone, address, tin, vat_type, business_type, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            request.form.get("name"),
            request.form.get("email"),
            request.form.get("phone"),
            request.form.get("address"),
            request.form.get("tin"),
            request.form.get("vat_type"),
            request.form.get("business_type"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        conn.commit()
        conn.close()
        flash("Supplier added successfully.", "success")
        return redirect(url_for("suppliers"))

    supplier_list = conn.execute("SELECT * FROM suppliers ORDER BY name ASC").fetchall()
    conn.close()
    return render_template("suppliers.html", suppliers=supplier_list)


@app.route("/items", methods=["GET", "POST"])
def items():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()

    if request.method == "POST":
        conn.execute("""
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

    where_sql, params = branch_where("items")
    item_list = conn.execute(f"""
        SELECT items.*, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        {where_sql}
        ORDER BY items.item_name ASC
    """, params).fetchall()

    supplier_list = conn.execute("SELECT * FROM suppliers ORDER BY name ASC").fetchall()
    conn.close()
    return render_template("items.html", items=item_list, suppliers=supplier_list)


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

        conn.execute("UPDATE items SET quantity = quantity + ? WHERE id = ?", (quantity, item_id))
        conn.execute("""
            INSERT INTO stock_transactions
            (item_id, transaction_type, quantity, remarks, encoded_by, date_created)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            item_id, "Stock In", quantity, remarks, session.get("full_name"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()
        flash("Stock added successfully.", "success")
        return redirect(url_for("stock_in"))

    where_sql, params = branch_where("items")
    item_list = conn.execute(f"SELECT * FROM items {where_sql} ORDER BY item_name ASC", params).fetchall()
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

        item = conn.execute("SELECT * FROM items WHERE id = ?", (item_id,)).fetchone()

        if quantity <= 0:
            flash("Quantity must be greater than zero.", "danger")
            conn.close()
            return redirect(url_for("stock_out"))

        if not item or item["quantity"] < quantity:
            flash("Not enough stock available.", "danger")
            conn.close()
            return redirect(url_for("stock_out"))

        conn.execute("UPDATE items SET quantity = quantity - ? WHERE id = ?", (quantity, item_id))
        conn.execute("""
            INSERT INTO stock_transactions
            (item_id, transaction_type, quantity, remarks, encoded_by, date_created)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            item_id, "Stock Out", quantity, remarks, session.get("full_name"),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()
        flash("Stock released successfully.", "success")
        return redirect(url_for("stock_out"))

    where_sql, params = branch_where("items")
    item_list = conn.execute(f"SELECT * FROM items {where_sql} ORDER BY item_name ASC", params).fetchall()
    conn.close()
    return render_template("stock_out.html", items=item_list)


@app.route("/transactions")
def transactions():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    rows = conn.execute("""
        SELECT stock_transactions.*, items.item_name, items.branch
        FROM stock_transactions
        JOIN items ON items.id = stock_transactions.item_id
        ORDER BY stock_transactions.id DESC
    """).fetchall()
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

        cur = conn.cursor()
        cur.execute("""
            INSERT INTO purchase_orders
            (po_number, supplier_id, po_date, requested_by, status, notes, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            po_number, supplier_id, po_date, requested_by, "Pending", notes,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        po_id = cur.lastrowid
        descriptions = request.form.getlist("description[]")
        quantities = request.form.getlist("po_quantity[]")
        units = request.form.getlist("po_unit[]")
        prices = request.form.getlist("unit_price[]")

        for desc, qty, unit, price in zip(descriptions, quantities, units, prices):
            if desc and desc.strip():
                conn.execute("""
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

    po_list = conn.execute("""
        SELECT purchase_orders.*, suppliers.name AS supplier_name
        FROM purchase_orders
        LEFT JOIN suppliers ON suppliers.id = purchase_orders.supplier_id
        ORDER BY purchase_orders.id DESC
    """).fetchall()

    supplier_list = conn.execute("SELECT * FROM suppliers ORDER BY name ASC").fetchall()
    conn.close()
    return render_template("purchase_orders.html", purchase_orders=po_list, suppliers=supplier_list)


@app.route("/purchase-orders/<int:po_id>/print")
def print_purchase_order(po_id):
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    po = conn.execute("""
        SELECT purchase_orders.*, suppliers.name AS supplier_name, suppliers.address,
               suppliers.tin, suppliers.phone, suppliers.email
        FROM purchase_orders
        LEFT JOIN suppliers ON suppliers.id = purchase_orders.supplier_id
        WHERE purchase_orders.id = ?
    """, (po_id,)).fetchone()

    po_items = conn.execute("""
        SELECT *, quantity * unit_price AS line_total
        FROM purchase_order_items
        WHERE po_id = ?
    """, (po_id,)).fetchall()
    conn.close()

    if not po:
        flash("Purchase order not found.", "danger")
        return redirect(url_for("purchase_orders"))

    grand_total = sum([item["line_total"] or 0 for item in po_items])
    return render_template("purchase_order_print.html", po=po, po_items=po_items, grand_total=grand_total)


@app.route("/reports")
def reports():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    total_items = conn.execute("SELECT COUNT(*) FROM items").fetchone()[0]
    low_stock_count = conn.execute("SELECT COUNT(*) FROM items WHERE quantity <= low_stock_limit").fetchone()[0]
    stock_in_count = conn.execute("SELECT COUNT(*) FROM stock_transactions WHERE transaction_type = 'Stock In'").fetchone()[0]
    stock_out_count = conn.execute("SELECT COUNT(*) FROM stock_transactions WHERE transaction_type = 'Stock Out'").fetchone()[0]
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
    rows = conn.execute("""
        SELECT items.item_name, items.category, items.unit, items.quantity,
               items.low_stock_limit, items.branch, suppliers.name AS supplier_name
        FROM items
        LEFT JOIN suppliers ON suppliers.id = items.supplier_id
        ORDER BY items.item_name ASC
    """).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory Items"
    ws.append(["Iligan Cement Multi-Purpose Cooperative"])
    ws.append(["Inventory Items Report"])
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
        download_name="ICMPC_Inventory_Items.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/export-transactions")
def export_transactions():
    check = require_login()
    if check:
        return check

    conn = get_db_connection()
    rows = conn.execute("""
        SELECT stock_transactions.*, items.item_name, items.branch
        FROM stock_transactions
        JOIN items ON items.id = stock_transactions.item_id
        ORDER BY stock_transactions.id DESC
    """).fetchall()
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


@app.route("/health")
def health():
    return "OK - ICMPC Inventory System is running"


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)