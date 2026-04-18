import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import re
import sqlite3
import os
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

# ==========================================
#           CACHE DATABASE SETUP
# ==========================================
# DB file stored in the same folder as the script
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rates_cache.db")
def init_db():
    """Create the cache table if it doesn't exist."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS rate_cache (
            from_pin  TEXT NOT NULL,
            to_pin    TEXT NOT NULL,
            weight    TEXT NOT NULL,
            service   TEXT NOT NULL,
            price     TEXT NOT NULL,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY (from_pin, to_pin, weight, service)
        )
    """)
    conn.commit()
    conn.close()


def cache_lookup(from_pin, to_pin, weight):
    """
    Return cached rows for this combination if they exist and are not expired.
    Returns a list of (from_pin, to_pin, weight, service, price) tuples, or None.
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        """SELECT from_pin, to_pin, weight, service, price
           FROM rate_cache
           WHERE from_pin=? AND to_pin=? AND weight=?""",
        (from_pin, to_pin, weight)
    )
    rows = cur.fetchall()
    conn.close()
    return rows if rows else None


def cache_save(from_pin, to_pin, weight, service_rows):
    """
    Upsert a batch of results into the cache.
    service_rows: list of (service, price) tuples.
    """
    now = datetime.now().isoformat()
    conn = sqlite3.connect(DB_PATH)
    conn.executemany(
        """INSERT OR REPLACE INTO rate_cache
           (from_pin, to_pin, weight, service, price, fetched_at)
           VALUES (?, ?, ?, ?, ?, ?)""",
        [(from_pin, to_pin, weight, svc, price, now) for svc, price in service_rows]
    )
    conn.commit()
    conn.close()


def cache_clear():
    """Delete all rows from the cache."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM rate_cache")
    conn.commit()
    conn.close()


def cache_count():
    """Return total number of cached route+weight combinations (not rows)."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT COUNT(DISTINCT from_pin || '|' || to_pin || '|' || weight) FROM rate_cache"
    )
    count = cur.fetchone()[0]
    conn.close()
    return count


# ==========================================
#              BACKEND WORKER
# ==========================================
def fetch_rates_backend(task_list, tree, status_label, buttons):
    for btn in buttons:
        btn.config(state=tk.DISABLED)

    for item in tree.get_children():
        tree.delete(item)

    total_tasks = len(task_list)
    current_task = 0
    cache_hits = 0

    # ---- STEP 1: Serve whatever is already in cache ----
    tasks_to_fetch = []
    for f_pin, t_pin, w in task_list:
        f_pin, t_pin, w = str(f_pin).strip(), str(t_pin).strip(), str(w).strip()
        cached = cache_lookup(f_pin, t_pin, w)
        if cached:
            cache_hits += 1
            for row in cached:
                tree.insert('', tk.END, values=row, tags=('cached',))
        else:
            tasks_to_fetch.append((f_pin, t_pin, w))

    tree.tag_configure('cached', foreground='#2e7d32')   # green = from cache
    tree.tag_configure('fresh',  foreground='#1565c0')   # blue  = freshly fetched

    if not tasks_to_fetch:
        status_label.config(
            text=f"Status: ✅ All {total_tasks} results served from cache (instant)! "
                 f"Cache has {cache_count()} stored combinations."
        )
        for btn in buttons:
            btn.config(state=tk.NORMAL)
        return

    # ---- STEP 2: Fetch only the uncached combinations via Playwright ----
    needs_fetch = len(tasks_to_fetch)
    status_label.config(
        text=f"Status: {cache_hits} from cache. Fetching {needs_fetch} new combinations... "
             f"Launching browser..."
    )

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            for f_pin, t_pin, w in tasks_to_fetch:
                current_task += 1
                status_label.config(
                    text=f"Status: [{cache_hits} cached | {current_task}/{needs_fetch} fetching] "
                         f"From {f_pin} → {t_pin} | {w}g ..."
                )

                url = (
                    f"https://www.indiapost.gov.in/calculate-postage"
                    f"?tab=domestic&fromPincode={f_pin}&toPincode={t_pin}"
                    f"&product=letter&weight={w}&POD=false&ACK=0&REG=0&COD=0"
                )

                try:
                    page.goto(url, wait_until="networkidle")
                    page.wait_for_selector("text='Speed Post'", timeout=15000)

                    rows = page.locator("tbody tr").all()
                    service_rows_to_cache = []

                    for row in rows:
                        row_text = row.inner_text()

                        service_name = "Unknown"
                        if "Speed Post" in row_text:
                            service_name = "Speed Post"
                        elif "Letter Card" in row_text:
                            service_name = "Letter Card"
                        elif "Letter" in row_text:
                            service_name = "Letter"
                        elif "Postcard" in row_text:
                            service_name = "Postcard"
                        elif "Parcel" in row_text:
                            service_name = "Parcel"

                        price_match = re.search(r'(\d+\.\d{2})', row_text)
                        price = f"₹ {price_match.group(1)}" if price_match else "N/A"

                        if service_name != "Unknown" and price != "N/A":
                            tree.insert('', tk.END,
                                        values=(f_pin, t_pin, f"{w}g", service_name, price),
                                        tags=('fresh',))
                            service_rows_to_cache.append((service_name, price))

                    # Save freshly fetched results to cache
                    if service_rows_to_cache:
                        cache_save(f_pin, t_pin, w, service_rows_to_cache)

                except Exception as inner_e:
                    tree.insert('', tk.END, values=(f_pin, t_pin, f"{w}g", "ERROR", "N/A"))
                    print(f"Failed on {f_pin} -> {t_pin} ({w}g): {inner_e}")

            browser.close()

        status_label.config(
            text=f"Status: ✅ Done! {cache_hits} from cache (🟢), "
                 f"{current_task} freshly fetched (🔵). "
                 f"Cache now has {cache_count()} stored combinations."
        )

    except Exception as e:
        status_label.config(text="Status: ❌ Fatal error occurred.")
        messagebox.showerror("Error", f"Failed to launch browser or perform fetch.\n\nDetails: {e}")

    finally:
        for btn in buttons:
            btn.config(state=tk.NORMAL)


# ==========================================
#           UI TRIGGER FUNCTIONS
# ==========================================
def run_manual_fetch():
    raw_from   = entry_from.get().split(',')
    raw_to     = entry_to.get().split(',')
    raw_weight = entry_weight.get().split(',')

    from_pins = [p.strip() for p in raw_from   if p.strip()]
    to_pins   = [p.strip() for p in raw_to     if p.strip()]
    weights   = [w.strip() for w in raw_weight if w.strip()]

    if not from_pins or not to_pins or not weights:
        messagebox.showwarning("Missing Input", "Please fill in all manual fields.")
        return

    task_list = [(f, t, w) for f in from_pins for t in to_pins for w in weights]
    start_thread(task_list)


def load_excel_and_run():
    filepath = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if not filepath:
        return

    try:
        df = pd.read_excel(filepath)
        for col in ['From', 'To', 'Weight']:
            if col not in df.columns:
                messagebox.showerror("Invalid Excel",
                                     f"Missing column: '{col}'. Headers must be 'From', 'To', 'Weight'.")
                return

        task_list = [
            (str(row['From']).replace('.0', ''),
             str(row['To']).replace('.0', ''),
             str(row['Weight']).replace('.0', ''))
            for _, row in df.iterrows()
            if pd.notna(row['From']) and pd.notna(row['To']) and pd.notna(row['Weight'])
        ]

        if not task_list:
            messagebox.showwarning("Empty Data", "No valid rows found in the Excel file.")
            return

        start_thread(task_list)

    except Exception as e:
        messagebox.showerror("Excel Error", f"Failed to read file:\n{e}")


def export_to_excel():
    if not tree.get_children():
        messagebox.showwarning("No Data", "There is no data to export!")
        return

    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        title="Save Results As",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if not filepath:
        return

    try:
        data = [tree.item(row_id)['values'] for row_id in tree.get_children()]
        df = pd.DataFrame(data, columns=['From Pincode', 'To Pincode', 'Weight', 'Service Type', 'Price'])
        df.to_excel(filepath, index=False)
        messagebox.showinfo("Success", f"Data exported successfully to:\n{filepath}")
    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to save file:\n{e}")


def clear_cache_prompt():
    count = cache_count()
    if count == 0:
        messagebox.showinfo("Cache Empty", "Cache is already empty — nothing to clear.")
        return
    confirm = messagebox.askyesno(
        "Clear Cache",
        f"This will delete {count} cached pincode combinations.\n\n"
        f"They will be re-fetched from India Post next time.\n\nProceed?"
    )
    if confirm:
        cache_clear()
        status_label.config(text="Status: Cache cleared. All results will be fetched fresh next run.")


def start_thread(task_list):
    threading.Thread(
        target=fetch_rates_backend,
        args=(task_list, tree, status_label, [btn_manual, btn_excel, btn_export, btn_clear_cache]),
        daemon=True
    ).start()


# ==========================================
#              GUI SETUP
# ==========================================
init_db()  # Ensure DB and table exist on startup

root = tk.Tk()
root.title("India Post Bulk Rate Calculator Pro")
root.geometry("800x580")
root.configure(padx=20, pady=20)

style = ttk.Style()
style.theme_use('clam')

# --- Input Area ---
frame_inputs = ttk.LabelFrame(root, text=" Manual Entry (Use commas for multiple values) ", padding=10)
frame_inputs.pack(fill="x", pady=(0, 10))

ttk.Label(frame_inputs, text="From Pincode:").grid(row=0, column=0, sticky="w", pady=5)
entry_from = ttk.Entry(frame_inputs, width=40)
entry_from.insert(0, "400710")
entry_from.grid(row=0, column=1, pady=5, padx=10)

ttk.Label(frame_inputs, text="To Pincodes:").grid(row=1, column=0, sticky="w", pady=5)
entry_to = ttk.Entry(frame_inputs, width=40)
entry_to.insert(0, "452001, 110001")
entry_to.grid(row=1, column=1, pady=5, padx=10)

ttk.Label(frame_inputs, text="Weights (grams):").grid(row=2, column=0, sticky="w", pady=5)
entry_weight = ttk.Entry(frame_inputs, width=40)
entry_weight.insert(0, "50, 100")
entry_weight.grid(row=2, column=1, pady=5, padx=10)

# Cache info label (right side of input frame)
lbl_cache_info = ttk.Label(frame_inputs, text="", foreground="#555", font=("Segoe UI", 8))
lbl_cache_info.grid(row=0, column=2, rowspan=3, padx=(20, 0), sticky="n")

def refresh_cache_label():
    lbl_cache_info.config(text=f"💾 Cache: {cache_count()} combinations stored\n(Use '🗑 Clear Cache' to reset)")

refresh_cache_label()

# --- Button Panel ---
frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=5)

btn_manual = ttk.Button(frame_buttons, text="🔍 Fetch Manual Input", command=run_manual_fetch)
btn_manual.grid(row=0, column=0, padx=5)

btn_excel = ttk.Button(frame_buttons, text="📂 Load Excel File", command=load_excel_and_run)
btn_excel.grid(row=0, column=1, padx=5)

btn_export = ttk.Button(frame_buttons, text="💾 Export Results to Excel", command=export_to_excel)
btn_export.grid(row=0, column=2, padx=5)

btn_clear_cache = ttk.Button(frame_buttons, text="🗑 Clear Cache", command=clear_cache_prompt)
btn_clear_cache.grid(row=0, column=3, padx=5)

# --- Status Bar ---
status_label = ttk.Label(root, text=f"Status: Ready | Cache has {cache_count()} stored combinations.", foreground="gray")
status_label.pack(pady=5)

# --- Legend ---
legend = ttk.Label(root, text="🟢 Green = served from cache (instant)   🔵 Blue = freshly fetched from India Post",
                   font=("Segoe UI", 8), foreground="#444")
legend.pack()

# --- Results Table ---
columns = ('From', 'To', 'Weight', 'Service', 'Price')
tree = ttk.Treeview(root, columns=columns, show='headings', height=12)

for col in columns:
    tree.heading(col, text=col)

tree.column('From',    width=90,  anchor='center')
tree.column('To',      width=90,  anchor='center')
tree.column('Weight',  width=80,  anchor='center')
tree.column('Service', width=160)
tree.column('Price',   width=90,  anchor='center')

# Scrollbar
scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True, pady=10)
tree.pack(in_=tree_frame, side="left", fill="both", expand=True)
scrollbar.pack(in_=tree_frame, side="right", fill="y")

root.mainloop()
