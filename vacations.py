import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import sqlite3
from datetime import datetime
import pandas as pd
import os
import hashlib
import logging

# ========================================================
# إعداد نظام السجلات (Logging)
# ========================================================
logging.basicConfig(
    filename='system_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

# ========================================================
# 1. إعداد قاعدة البيانات الكاملة
# ========================================================
def init_db():
    if not os.path.exists('backup'):
        os.makedirs('backup')
    try:
        conn = get_connection()
        cursor = conn.cursor()

        # جدول المستخدمين (نظام الصلاحيات)
        cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            username TEXT NOT NULL UNIQUE,
                            password_hash TEXT NOT NULL,
                            role TEXT NOT NULL DEFAULT 'viewer',
                            created_at TEXT)''')

        # جدول الموظفين
        cursor.execute('''CREATE TABLE IF NOT EXISTS employees (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            name TEXT NOT NULL,
                            initial_balance INTEGER,
                            total_deducted INTEGER DEFAULT 0,
                            last_deduct_info TEXT,
                            join_date TEXT)''')

        # جدول سجل الإجازات الكامل (Audit Log)
        cursor.execute('''CREATE TABLE IF NOT EXISTS vacation_history (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            employee_id INTEGER NOT NULL,
                            employee_name TEXT NOT NULL,
                            days INTEGER NOT NULL,
                            vacation_date TEXT NOT NULL,
                            notes TEXT,
                            created_by TEXT,
                            created_at TEXT NOT NULL,
                            is_cancelled INTEGER DEFAULT 0,
                            cancelled_by TEXT,
                            cancelled_at TEXT,
                            FOREIGN KEY (employee_id) REFERENCES employees(id))''')

        # إضافة مدير افتراضي إذا ما في مستخدمين
        cursor.execute("SELECT COUNT(*) FROM users")
        if cursor.fetchone()[0] == 0:
            admin_hash = hash_password("admin123")
            cursor.execute(
                "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)",
                ("admin", admin_hash, "admin", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )
            logging.info("تم إنشاء المستخدم الافتراضي admin")

        conn.commit()
        conn.close()
    except Exception as e:
        logging.error(f"خطأ في تهيئة قاعدة البيانات: {e}")
        messagebox.showerror("خطأ فادح", f"فشل تهيئة قاعدة البيانات:\n{e}")

# ========================================================
# 2. إدارة الاتصال بقاعدة البيانات
# ========================================================
def get_connection():
    try:
        conn = sqlite3.connect('vacation_pro_system.db', timeout=10)
        conn.execute("PRAGMA journal_mode=WAL")
        return conn
    except sqlite3.Error as e:
        logging.error(f"فشل الاتصال بقاعدة البيانات: {e}")
        messagebox.showerror("خطأ", f"فشل الاتصال بقاعدة البيانات:\n{e}")
        return None

# ========================================================
# 3. نظام كلمة المرور
# ========================================================
def hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def verify_login(username, password):
    conn = get_connection()
    if not conn:
        return None
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT role FROM users WHERE username=? AND password_hash=?",
                       (username, hash_password(password)))
        result = cursor.fetchone()
        return result[0] if result else None
    except Exception as e:
        logging.error(f"خطأ في التحقق من تسجيل الدخول: {e}")
        return None
    finally:
        conn.close()

# ========================================================
# 4. حساب الرصيد (المعادلة المحسّنة)
# ========================================================
def calculate_current_balance(initial_bal, join_date_str, total_deducted):
    try:
        join_date = datetime.strptime(join_date_str, "%Y-%m-%d")
        now = datetime.now()
        # حساب دقيق يأخذ اليوم بعين الاعتبار
        diff_months = (now.year - join_date.year) * 12 + (now.month - join_date.month)
        if now.day < join_date.day:
            diff_months -= 1
        if diff_months < 0:
            diff_months = 0
        return initial_bal + (diff_months * 3) - total_deducted
    except Exception as e:
        logging.warning(f"خطأ في حساب الرصيد: {e}")
        return initial_bal

# ========================================================
# 5. التحقق من المدخلات
# ========================================================
def validate_days(days_str):
    try:
        days = int(days_str)
        if days <= 0:
            return False, "عدد الأيام يجب أن يكون أكبر من صفر"
        if days > 365:
            return False, "عدد الأيام لا يمكن أن يتجاوز 365"
        return True, days
    except ValueError:
        return False, "عدد الأيام يجب أن يكون رقماً صحيحاً"

def validate_date(day_str, month_str, year_str):
    try:
        day, month, year = int(day_str), int(month_str), int(year_str)
        if len(year_str) != 4:
            return False, "السنة يجب أن تكون 4 أرقام"
        date_obj = datetime(year, month, day)
        if date_obj > datetime.now():
            return False, "لا يمكن تسجيل إجازة بتاريخ مستقبلي"
        return True, date_obj.strftime("%Y-%m-%d")
    except ValueError as e:
        return False, f"تاريخ غير صحيح: {e}"

# ========================================================
# 6. استيراد من Excel
# ========================================================
def import_from_excel():
    if current_role not in ("admin", "editor"):
        messagebox.showwarning("صلاحية", "ليس لديك صلاحية لاستيراد البيانات")
        return
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    try:
        df = pd.read_excel(file_path)
        conn = get_connection()
        if not conn:
            return
        cursor = conn.cursor()
        today = datetime.now().strftime("%Y-%m-%d")
        count = 0
        for _, row in df.iterrows():
            name = str(row.get('الاسم', '')).strip()
            if not name or name == 'nan':
                continue
            init_bal = int(row.get('الرصيد الابتدائي', row.get('الرصيد', 0)))
            total_ded = int(row.get('مجموع الخصومات', 0))
            last_info = str(row.get('آخر اجازة', "مستورد"))
            j_date = str(row.get('تاريخ الانضمام', today)).split()[0]
            cursor.execute(
                "INSERT INTO employees (name, initial_balance, total_deducted, last_deduct_info, join_date) VALUES (?, ?, ?, ?, ?)",
                (name, init_bal, total_ded, last_info, j_date)
            )
            count += 1
        conn.commit()
        conn.close()
        load_data()
        logging.info(f"تم استيراد {count} موظف من {file_path} بواسطة {current_user}")
        messagebox.showinfo("نجاح", f"تم استيراد {count} موظف بنجاح")
    except Exception as e:
        logging.error(f"فشل الاستيراد: {e}")
        messagebox.showerror("خطأ", f"فشل الاستيراد: تأكد من مسميات الأعمدة\nالخطأ: {str(e)}")

# ========================================================
# 7. تصدير إلى Excel
# ========================================================
def export_to_excel_manual():
    try:
        conn = get_connection()
        if not conn:
            return
        df = pd.read_sql_query(
            "SELECT name as 'الاسم', initial_balance as 'الرصيد الابتدائي', total_deducted as 'مجموع الخصومات', last_deduct_info as 'آخر اجازة', join_date as 'تاريخ الانضمام' FROM employees",
            conn
        )
        conn.close()
        if df.empty:
            messagebox.showwarning("تنبيه", "لا توجد بيانات لتصديرها")
            return
        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False, engine='openpyxl')
            logging.info(f"تم التصدير إلى {file_path} بواسطة {current_user}")
            messagebox.showinfo("نجاح", "تم تصدير الملف بنجاح")
    except Exception as e:
        logging.error(f"فشل التصدير: {e}")
        messagebox.showerror("خطأ", f"فشل التصدير: {str(e)}")

# ========================================================
# 8. إضافة موظف
# ========================================================
def add_employee():
    if current_role not in ("admin", "editor"):
        messagebox.showwarning("صلاحية", "ليس لديك صلاحية لإضافة موظفين")
        return
    name = entry_name.get().strip()
    bal = entry_initial.get().strip()
    if not name or not bal:
        messagebox.showwarning("تنبيه", "أكمل جميع الحقول")
        return
    try:
        bal_int = int(float(bal))
        if bal_int < 0:
            messagebox.showerror("خطأ", "الرصيد لا يمكن أن يكون سالباً")
            return
        today = datetime.now().strftime("%Y-%m-%d")
        conn = get_connection()
        if not conn:
            return
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO employees (name, initial_balance, total_deducted, last_deduct_info, join_date) VALUES (?, ?, ?, ?, ?)",
            (name, bal_int, 0, "لا يوجد", today)
        )
        conn.commit()
        conn.close()
        load_data()
        entry_name.delete(0, tk.END)
        entry_initial.delete(0, tk.END)
        logging.info(f"تمت إضافة الموظف '{name}' بواسطة {current_user}")
    except ValueError:
        messagebox.showerror("خطأ", "الرصيد الابتدائي يجب أن يكون رقماً صحيحاً")
    except Exception as e:
        logging.error(f"خطأ في إضافة موظف: {e}")
        messagebox.showerror("خطأ", f"فشل الإضافة: {e}")

# ========================================================
# 9. حذف موظف
# ========================================================
def delete_employee():
    if current_role != "admin":
        messagebox.showwarning("صلاحية", "فقط المدير يمكنه حذف الموظفين")
        return
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("تنبيه", "اختر موظفاً أولاً")
        return
    emp_id = tree.item(selected)['values'][0]
    emp_name = tree.item(selected)['values'][1]
    if messagebox.askyesno("تأكيد الحذف", f"هل أنت متأكد من حذف الموظف '{emp_name}'؟\nسيتم حذف جميع سجلات إجازاته أيضاً."):
        try:
            conn = get_connection()
            if not conn:
                return
            cursor = conn.cursor()
            cursor.execute("DELETE FROM vacation_history WHERE employee_id=?", (emp_id,))
            cursor.execute("DELETE FROM employees WHERE id=?", (emp_id,))
            conn.commit()
            conn.close()
            load_data()
            logging.warning(f"تم حذف الموظف '{emp_name}' (ID:{emp_id}) بواسطة {current_user}")
        except Exception as e:
            logging.error(f"خطأ في حذف موظف: {e}")
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")

# ========================================================
# 10. تحميل البيانات
# ========================================================
def load_data(filter_name=""):
    for row in tree.get_children():
        tree.delete(row)
    conn = get_connection()
    if not conn:
        return
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, name, initial_balance, total_deducted, last_deduct_info, join_date FROM employees WHERE name LIKE ?",
            ('%' + filter_name + '%',)
        )
        rows = cursor.fetchall()
        for row in rows:
            emp_id, name, init_bal, total_ded, last_info, join_date = row
            current_bal = calculate_current_balance(init_bal, join_date, total_ded)
            tag = 'low' if current_bal <= 5 else 'normal'
            tree.insert('', tk.END, values=(emp_id, name, current_bal, last_info), tags=(tag,))
        tree.tag_configure('low', background='#ffe0e0')
        tree.tag_configure('normal', background='white')
        update_status_bar(len(rows))
    except Exception as e:
        logging.error(f"خطأ في تحميل البيانات: {e}")
        messagebox.showerror("خطأ", f"فشل تحميل البيانات: {e}")
    finally:
        conn.close()

# ========================================================
# 11. خصم إجازة (مع التحقق الكامل)
# ========================================================
def deduct_vacation():
    if current_role not in ("admin", "editor"):
        messagebox.showwarning("صلاحية", "ليس لديك صلاحية لتسجيل الإجازات")
        return
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("تنبيه", "اختر موظفاً أولاً")
        return

    days_str = entry_deduct.get().strip()
    day_str = entry_day.get().strip()
    month_str = entry_month.get().strip()
    year_str = entry_year.get().strip()
    notes_str = entry_notes.get().strip()

    if not all([days_str, day_str, month_str, year_str]):
        messagebox.showwarning("تنبيه", "أكمل جميع بيانات الإجازة")
        return

    valid_days, days_result = validate_days(days_str)
    if not valid_days:
        messagebox.showerror("خطأ في البيانات", days_result)
        return

    valid_date, date_result = validate_date(day_str, month_str, year_str)
    if not valid_date:
        messagebox.showerror("خطأ في التاريخ", date_result)
        return

    emp_id = tree.item(selected)['values'][0]
    emp_name = tree.item(selected)['values'][1]
    current_bal = tree.item(selected)['values'][2]

    if days_result > current_bal:
        if not messagebox.askyesno("تحذير رصيد",
                                   f"الموظف '{emp_name}' رصيده الحالي {current_bal} يوم فقط!\n"
                                   f"هل تريد خصم {days_result} يوم وجعل الرصيد سالباً؟"):
            return

    try:
        info = f"خصم {days_result} يوم بتاريخ {day_str}/{month_str}/{year_str}"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = get_connection()
        if not conn:
            return
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE employees SET total_deducted = total_deducted + ?, last_deduct_info = ? WHERE id = ?",
            (days_result, info, emp_id)
        )
        cursor.execute(
            "INSERT INTO vacation_history (employee_id, employee_name, days, vacation_date, notes, created_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (emp_id, emp_name, days_result, date_result, notes_str, current_user, now_str)
        )
        conn.commit()
        conn.close()
        load_data()
        for e in [entry_deduct, entry_day, entry_month, entry_year, entry_notes]:
            e.delete(0, tk.END)
        logging.info(f"تم خصم {days_result} يوم من '{emp_name}' بواسطة {current_user}")
        messagebox.showinfo("تم", f"تم تسجيل إجازة {days_result} يوم للموظف {emp_name} ✓")
    except Exception as e:
        logging.error(f"خطأ في خصم الإجازة: {e}")
        messagebox.showerror("خطأ", f"فشل تسجيل الإجازة: {e}")

# ========================================================
# 12. نافذة سجل الإجازات (مع إمكانية الإلغاء)
# ========================================================
def open_history_window():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("تنبيه", "اختر موظفاً لعرض سجله")
        return
    emp_id = tree.item(selected)['values'][0]
    emp_name = tree.item(selected)['values'][1]

    hist_win = tk.Toplevel(root)
    hist_win.title(f"سجل إجازات: {emp_name}")
    hist_win.geometry("850x450")
    hist_win.grab_set()
    hist_win.configure(bg="#f5f5f5")

    tk.Label(hist_win, text=f"سجل إجازات الموظف: {emp_name}",
             font=("Arial", 14, "bold"), bg="#f5f5f5").pack(pady=10)

    cols = ('رقم', 'الأيام', 'تاريخ الإجازة', 'ملاحظات', 'سجّل بواسطة', 'وقت التسجيل', 'الحالة')
    hist_tree = ttk.Treeview(hist_win, columns=cols, show='headings', height=15)
    for col in cols:
        hist_tree.heading(col, text=col)
        hist_tree.column(col, anchor="center", width=110)
    hist_tree.pack(fill="both", expand=True, padx=10, pady=5)

    scrollbar = ttk.Scrollbar(hist_win, orient="vertical", command=hist_tree.yview)
    hist_tree.configure(yscrollcommand=scrollbar.set)

    def refresh_history():
        for row in hist_tree.get_children():
            hist_tree.delete(row)
        conn = get_connection()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, days, vacation_date, notes, created_by, created_at, is_cancelled FROM vacation_history WHERE employee_id=? ORDER BY created_at DESC",
                (emp_id,)
            )
            for row in cursor.fetchall():
                h_id, days, v_date, notes, by, at, cancelled = row
                status = "❌ ملغي" if cancelled else "✅ فعّال"
                tag = 'cancelled' if cancelled else 'active'
                hist_tree.insert('', tk.END,
                                 values=(h_id, days, v_date, notes or "-", by or "-", at, status),
                                 tags=(tag,))
            hist_tree.tag_configure('cancelled', background='#f0f0f0', foreground='gray')
            hist_tree.tag_configure('active', background='white')
        except Exception as e:
            logging.error(f"خطأ في تحميل السجل: {e}")
        finally:
            conn.close()

    def cancel_vacation():
        if current_role != "admin":
            messagebox.showwarning("صلاحية", "فقط المدير يمكنه إلغاء الإجازات")
            return
        sel = hist_tree.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر إجازة للإلغاء")
            return
        item = hist_tree.item(sel)
        h_id = item['values'][0]
        days = item['values'][1]
        status = item['values'][6]
        if "ملغي" in str(status):
            messagebox.showinfo("تنبيه", "هذه الإجازة ملغاة مسبقاً")
            return
        if messagebox.askyesno("تأكيد الإلغاء",
                               f"هل تريد إلغاء هذه الإجازة ({days} يوم)؟\nسيتم إعادة الأيام لرصيد الموظف."):
            try:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                conn = get_connection()
                if not conn:
                    return
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE vacation_history SET is_cancelled=1, cancelled_by=?, cancelled_at=? WHERE id=?",
                    (current_user, now_str, h_id)
                )
                cursor.execute(
                    "UPDATE employees SET total_deducted = total_deducted - ? WHERE id=?",
                    (days, emp_id)
                )
                conn.commit()
                conn.close()
                refresh_history()
                load_data()
                logging.warning(f"تم إلغاء إجازة ID:{h_id} للموظف '{emp_name}' بواسطة {current_user}")
                messagebox.showinfo("تم", "تم إلغاء الإجازة وإعادة الأيام للرصيد ✓")
            except Exception as e:
                logging.error(f"خطأ في إلغاء الإجازة: {e}")
                messagebox.showerror("خطأ", f"فشل الإلغاء: {e}")

    btn_frame = tk.Frame(hist_win, bg="#f5f5f5")
    btn_frame.pack(pady=5)
    tk.Button(btn_frame, text="إلغاء الإجازة المحددة", command=cancel_vacation,
              bg="#dc3545", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5).pack(side="left", padx=10)
    tk.Button(btn_frame, text="تحديث", command=refresh_history,
              bg="#6c757d", fg="white", font=("Arial", 11), padx=15, pady=5).pack(side="left", padx=10)

    refresh_history()

# ========================================================
# 13. نافذة تعديل بيانات الموظف
# ========================================================
def open_edit_window():
    if current_role not in ("admin", "editor"):
        messagebox.showwarning("صلاحية", "ليس لديك صلاحية للتعديل")
        return
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("تنبيه", "اختر موظفاً لتعديله")
        return
    item = tree.item(selected)
    emp_id = item['values'][0]
    old_name = item['values'][1]
    current_displayed_bal = item['values'][2]

    edit_win = tk.Toplevel(root)
    edit_win.title("تعديل بيانات الموظف")
    edit_win.geometry("350x280")
    edit_win.grab_set()
    edit_win.configure(bg="#f5f5f5")

    tk.Label(edit_win, text="تعديل بيانات الموظف", font=("Arial", 13, "bold"), bg="#f5f5f5").pack(pady=10)

    frame = tk.Frame(edit_win, bg="#f5f5f5")
    frame.pack(padx=20, fill="x")

    tk.Label(frame, text="الاسم الجديد:", bg="#f5f5f5").grid(row=0, column=0, sticky="e", pady=8)
    new_name_entry = tk.Entry(frame, width=25)
    new_name_entry.insert(0, old_name)
    new_name_entry.grid(row=0, column=1, padx=10)

    tk.Label(frame, text="الرصيد الحالي الجديد:", bg="#f5f5f5").grid(row=1, column=0, sticky="e", pady=8)
    new_bal_entry = tk.Entry(frame, width=25)
    new_bal_entry.insert(0, current_displayed_bal)
    new_bal_entry.grid(row=1, column=1, padx=10)

    tk.Label(edit_win,
             text="⚠ سيتم تصفير المعادلة واعتبار\nهذا الرقم رصيداً جديداً يبدأ من اليوم",
             fg="darkorange", bg="#f5f5f5", font=("Arial", 9)).pack(pady=5)

    def save_changes():
        n_name = new_name_entry.get().strip()
        if not n_name:
            messagebox.showerror("خطأ", "الاسم لا يمكن أن يكون فارغاً")
            return
        try:
            n_val = int(new_bal_entry.get())
            if n_val < 0:
                messagebox.showerror("خطأ", "الرصيد لا يمكن أن يكون سالباً")
                return
            if messagebox.askyesno("تأكيد", "هل تريد حفظ التعديلات؟"):
                today = datetime.now().strftime("%Y-%m-%d")
                conn = get_connection()
                if not conn:
                    return
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE employees SET name=?, initial_balance=?, total_deducted=0, join_date=? WHERE id=?",
                    (n_name, n_val, today, emp_id)
                )
                conn.commit()
                conn.close()
                edit_win.destroy()
                load_data()
                logging.info(f"تم تعديل بيانات الموظف ID:{emp_id} بواسطة {current_user}")
        except ValueError:
            messagebox.showerror("خطأ", "أدخل رقماً صحيحاً للرصيد")
        except Exception as e:
            logging.error(f"خطأ في تعديل الموظف: {e}")
            messagebox.showerror("خطأ", f"فشل التعديل: {e}")

    tk.Button(edit_win, text="حفظ التعديلات", command=save_changes,
              bg="darkorange", fg="white", font=("Arial", 11, "bold"), padx=20, pady=8).pack(pady=15)

# ========================================================
# 14. نافذة إدارة المستخدمين
# ========================================================
def open_users_window():
    if current_role != "admin":
        messagebox.showwarning("صلاحية", "فقط المدير يمكنه إدارة المستخدمين")
        return

    users_win = tk.Toplevel(root)
    users_win.title("إدارة المستخدمين")
    users_win.geometry("650x500")
    users_win.grab_set()
    users_win.configure(bg="#f5f5f5")

    tk.Label(users_win, text="إدارة المستخدمين والصلاحيات",
             font=("Arial", 14, "bold"), bg="#f5f5f5").pack(pady=10)

    cols = ('ID', 'اسم المستخدم', 'الدور', 'تاريخ الإنشاء')
    u_tree = ttk.Treeview(users_win, columns=cols, show='headings', height=10)
    for col in cols:
        u_tree.heading(col, text=col)
        u_tree.column(col, anchor="center")
    u_tree.column('ID', width=40)
    u_tree.pack(fill="both", expand=True, padx=10, pady=5)

    def refresh_users():
        for row in u_tree.get_children():
            u_tree.delete(row)
        conn = get_connection()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id, username, role, created_at FROM users")
            for row in cursor.fetchall():
                u_tree.insert('', tk.END, values=row)
        finally:
            conn.close()

    add_frame = tk.LabelFrame(users_win, text="إضافة مستخدم جديد", bg="#f5f5f5", padx=10, pady=10)
    add_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(add_frame, text="اسم المستخدم:", bg="#f5f5f5").grid(row=0, column=0)
    u_name = tk.Entry(add_frame, width=20)
    u_name.grid(row=0, column=1, padx=5)

    tk.Label(add_frame, text="كلمة المرور:", bg="#f5f5f5").grid(row=0, column=2)
    u_pass = tk.Entry(add_frame, width=20, show="*")
    u_pass.grid(row=0, column=3, padx=5)

    tk.Label(add_frame, text="الدور:", bg="#f5f5f5").grid(row=0, column=4)
    role_var = tk.StringVar(value="viewer")
    role_combo = ttk.Combobox(add_frame, textvariable=role_var,
                               values=["admin", "editor", "viewer"], width=10, state="readonly")
    role_combo.grid(row=0, column=5, padx=5)

    def add_user():
        uname = u_name.get().strip()
        upass = u_pass.get().strip()
        urole = role_var.get()
        if not uname or not upass:
            messagebox.showwarning("تنبيه", "أكمل الحقول")
            return
        if len(upass) < 6:
            messagebox.showerror("خطأ", "كلمة المرور يجب أن تكون 6 أحرف على الأقل")
            return
        try:
            conn = get_connection()
            if not conn:
                return
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)",
                (uname, hash_password(upass), urole, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )
            conn.commit()
            conn.close()
            refresh_users()
            u_name.delete(0, tk.END)
            u_pass.delete(0, tk.END)
            logging.info(f"تم إضافة مستخدم '{uname}' بواسطة {current_user}")
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "اسم المستخدم موجود مسبقاً")
        except Exception as e:
            logging.error(f"خطأ في إضافة مستخدم: {e}")
            messagebox.showerror("خطأ", str(e))

    def delete_user():
        sel = u_tree.selection()
        if not sel:
            return
        u_id = u_tree.item(sel)['values'][0]
        u_uname = u_tree.item(sel)['values'][1]
        if u_uname == current_user:
            messagebox.showerror("خطأ", "لا يمكنك حذف حسابك الحالي")
            return
        if messagebox.askyesno("تأكيد", f"حذف المستخدم '{u_uname}'؟"):
            try:
                conn = get_connection()
                if not conn:
                    return
                cursor = conn.cursor()
                cursor.execute("DELETE FROM users WHERE id=?", (u_id,))
                conn.commit()
                conn.close()
                refresh_users()
                logging.warning(f"تم حذف المستخدم '{u_uname}' بواسطة {current_user}")
            except Exception as e:
                logging.error(f"خطأ في حذف مستخدم: {e}")
                messagebox.showerror("خطأ", str(e))

    btn_f = tk.Frame(add_frame, bg="#f5f5f5")
    btn_f.grid(row=1, column=0, columnspan=6, pady=5)
    tk.Button(btn_f, text="إضافة", command=add_user, bg="darkgreen", fg="white", padx=15).pack(side="left", padx=5)
    tk.Button(btn_f, text="حذف المحدد", command=delete_user, bg="#dc3545", fg="white", padx=15).pack(side="left", padx=5)

    # شرح الأدوار
    roles_info = tk.Label(users_win,
                          text="admin: صلاحيات كاملة  |  editor: إضافة وتعديل  |  viewer: عرض فقط",
                          font=("Arial", 9), bg="#f5f5f5", fg="gray")
    roles_info.pack(pady=5)

    refresh_users()

# ========================================================
# 15. نافذة تغيير كلمة المرور
# ========================================================
def open_change_password():
    pw_win = tk.Toplevel(root)
    pw_win.title("تغيير كلمة المرور")
    pw_win.geometry("320x220")
    pw_win.grab_set()
    pw_win.configure(bg="#f5f5f5")

    tk.Label(pw_win, text="تغيير كلمة المرور", font=("Arial", 13, "bold"), bg="#f5f5f5").pack(pady=10)
    frame = tk.Frame(pw_win, bg="#f5f5f5")
    frame.pack(padx=20)

    tk.Label(frame, text="كلمة المرور الحالية:", bg="#f5f5f5").grid(row=0, column=0, sticky="e", pady=5)
    old_pw = tk.Entry(frame, show="*", width=20)
    old_pw.grid(row=0, column=1)

    tk.Label(frame, text="كلمة المرور الجديدة:", bg="#f5f5f5").grid(row=1, column=0, sticky="e", pady=5)
    new_pw = tk.Entry(frame, show="*", width=20)
    new_pw.grid(row=1, column=1)

    tk.Label(frame, text="تأكيد كلمة المرور:", bg="#f5f5f5").grid(row=2, column=0, sticky="e", pady=5)
    conf_pw = tk.Entry(frame, show="*", width=20)
    conf_pw.grid(row=2, column=1)

    def do_change():
        op = old_pw.get()
        np_ = new_pw.get()
        cp = conf_pw.get()
        if len(np_) < 6:
            messagebox.showerror("خطأ", "كلمة المرور الجديدة أقل من 6 أحرف")
            return
        if np_ != cp:
            messagebox.showerror("خطأ", "كلمة المرور الجديدة وتأكيدها غير متطابقين")
            return
        conn = get_connection()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM users WHERE username=? AND password_hash=?",
                           (current_user, hash_password(op)))
            if not cursor.fetchone():
                messagebox.showerror("خطأ", "كلمة المرور الحالية غير صحيحة")
                return
            cursor.execute("UPDATE users SET password_hash=? WHERE username=?",
                           (hash_password(np_), current_user))
            conn.commit()
            pw_win.destroy()
            logging.info(f"تم تغيير كلمة مرور المستخدم '{current_user}'")
            messagebox.showinfo("تم", "تم تغيير كلمة المرور بنجاح ✓")
        except Exception as e:
            logging.error(f"خطأ في تغيير كلمة المرور: {e}")
            messagebox.showerror("خطأ", str(e))
        finally:
            conn.close()

    tk.Button(pw_win, text="حفظ", command=do_change,
              bg="#007bff", fg="white", font=("Arial", 11, "bold"), padx=20, pady=6).pack(pady=15)

# ========================================================
# 16. شاشة تسجيل الدخول
# ========================================================
def show_login_screen():
    global current_user, current_role

    login_win = tk.Toplevel()
    login_win.title("تسجيل الدخول - نظام الإجازات")
    login_win.geometry("420x320")
    login_win.grab_set()
    login_win.resizable(False, False)
    login_win.configure(bg="#1a2a4a")

    # header
    header = tk.Frame(login_win, bg="#1a2a4a")
    header.pack(fill="x", pady=20)
    tk.Label(header, text="🏢 نظام إدارة الإجازات",
             font=("Arial", 16, "bold"), fg="white", bg="#1a2a4a").pack()
    tk.Label(header, text="الإصدار 2.0 - المطوّر",
             font=("Arial", 9), fg="#aaaacc", bg="#1a2a4a").pack()

    card = tk.Frame(login_win, bg="white", padx=30, pady=20)
    card.pack(padx=30, pady=10, fill="both", expand=True)

    tk.Label(card, text="اسم المستخدم:", font=("Arial", 11), bg="white").pack(anchor="e")
    u_entry = tk.Entry(card, font=("Arial", 12), width=25, justify="right")
    u_entry.pack(pady=5)
    u_entry.insert(0, "admin")

    tk.Label(card, text="كلمة المرور:", font=("Arial", 11), bg="white").pack(anchor="e")
    p_entry = tk.Entry(card, font=("Arial", 12), width=25, show="*", justify="right")
    p_entry.pack(pady=5)

    status_label = tk.Label(card, text="", fg="red", bg="white", font=("Arial", 10))
    status_label.pack()

    def do_login(event=None):
        global current_user, current_role
        uname = u_entry.get().strip()
        upass = p_entry.get().strip()
        if not uname or not upass:
            status_label.config(text="أكمل جميع الحقول")
            return
        role = verify_login(uname, upass)
        if role:
            current_user = uname
            current_role = role
            logging.info(f"تسجيل دخول ناجح: {uname} ({role})")
            login_win.destroy()
            build_main_ui()
        else:
            status_label.config(text="❌ اسم المستخدم أو كلمة المرور غير صحيحة")
            logging.warning(f"محاولة دخول فاشلة بالمستخدم: {uname}")
            p_entry.delete(0, tk.END)

    p_entry.bind('<Return>', do_login)
    tk.Button(card, text="دخول", command=do_login,
              bg="#1a2a4a", fg="white", font=("Arial", 12, "bold"),
              width=20, pady=6, cursor="hand2").pack(pady=5)

    login_win.protocol("WM_DELETE_WINDOW", root.destroy)
    login_win.wait_window()

# ========================================================
# 17. شريط الحالة
# ========================================================
def update_status_bar(count=0):
    status_var.set(f"  المستخدم: {current_user}  |  الصلاحية: {current_role}  |  عدد الموظفين: {count}")

# ========================================================
# 18. بناء الواجهة الرئيسية
# ========================================================
def build_main_ui():
    global tree, entry_name, entry_initial, entry_deduct
    global entry_day, entry_month, entry_year, entry_notes, entry_search

    root.deiconify()
    root.title(f"نظام إدارة الإجازات v2.0 - {current_user}")
    root.geometry("1200x800")
    root.configure(bg="#f0f2f5")

    # ---- شريط القوائم العلوي ----
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="ملف", menu=file_menu)
    file_menu.add_command(label="استيراد من Excel", command=import_from_excel)
    file_menu.add_command(label="تصدير إلى Excel", command=export_to_excel_manual)
    file_menu.add_separator()
    file_menu.add_command(label="خروج", command=root.destroy)

    account_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="الحساب", menu=account_menu)
    account_menu.add_command(label="إدارة المستخدمين", command=open_users_window)
    account_menu.add_command(label="تغيير كلمة المرور", command=open_change_password)

    # ---- شريط البحث والأزرار العلوية ----
    top_bar = tk.Frame(root, bg="#1a2a4a", padx=15, pady=10)
    top_bar.pack(fill="x")

    tk.Label(top_bar, text="🏢 نظام إدارة الإجازات",
             font=("Arial", 14, "bold"), fg="white", bg="#1a2a4a").pack(side="right", padx=10)

    btn_style = {"font": ("Arial", 10, "bold"), "padx": 12, "pady": 5, "cursor": "hand2", "bd": 0}
    tk.Button(top_bar, text="📊 تصدير Excel", command=export_to_excel_manual,
              bg="#17a2b8", fg="white", **btn_style).pack(side="left", padx=4)
    tk.Button(top_bar, text="📥 استيراد Excel", command=import_from_excel,
              bg="#28a745", fg="white", **btn_style).pack(side="left", padx=4)
    tk.Button(top_bar, text="📋 سجل الإجازات", command=open_history_window,
              bg="#6f42c1", fg="white", **btn_style).pack(side="left", padx=4)
    tk.Button(top_bar, text="👥 المستخدمين", command=open_users_window,
              bg="#fd7e14", fg="white", **btn_style).pack(side="left", padx=4)

    # ---- شريط البحث ----
    search_bar = tk.Frame(root, bg="#e9ecef", padx=10, pady=8)
    search_bar.pack(fill="x")
    tk.Label(search_bar, text="🔍 بحث بالاسم:", font=("Arial", 11), bg="#e9ecef").pack(side="right")
    entry_search = tk.Entry(search_bar, width=35, font=("Arial", 11), justify="right")
    entry_search.pack(side="right", padx=10)
    entry_search.bind('<KeyRelease>', lambda e: load_data(entry_search.get()))

    # ---- إطار إضافة الموظفين ----
    add_frame = tk.LabelFrame(root, text="  إدارة الموظفين  ",
                               font=("Arial", 11, "bold"), bg="#f0f2f5", padx=15, pady=10)
    add_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(add_frame, text="الاسم:", font=("Arial", 11), bg="#f0f2f5").grid(row=0, column=0, padx=5)
    entry_name = tk.Entry(add_frame, width=28, font=("Arial", 11), justify="right")
    entry_name.grid(row=0, column=1, padx=5)

    tk.Label(add_frame, text="الرصيد الابتدائي:", font=("Arial", 11), bg="#f0f2f5").grid(row=0, column=2, padx=5)
    entry_initial = tk.Entry(add_frame, width=12, font=("Arial", 11))
    entry_initial.grid(row=0, column=3, padx=5)

    btn_s = {"font": ("Arial", 10, "bold"), "width": 10, "pady": 4, "cursor": "hand2"}
    tk.Button(add_frame, text="➕ إضافة", command=add_employee, bg="#28a745", fg="white", **btn_s).grid(row=0, column=4, padx=8)
    tk.Button(add_frame, text="✏ تعديل", command=open_edit_window, bg="#fd7e14", fg="white", **btn_s).grid(row=0, column=5, padx=5)
    tk.Button(add_frame, text="🗑 حذف", command=delete_employee, bg="#dc3545", fg="white", **btn_s).grid(row=0, column=6, padx=5)

    # ---- جدول الموظفين ----
    tree_frame = tk.Frame(root, bg="#f0f2f5")
    tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

    style = ttk.Style()
    style.configure("Treeview", font=("Arial", 11), rowheight=28)
    style.configure("Treeview.Heading", font=("Arial", 11, "bold"))

    columns = ('ID', 'الاسم الكامل', 'الرصيد المحدث', 'آخر إجازة')
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='browse')

    tree.heading('ID', text='#')
    tree.column('ID', width=0, stretch=tk.NO)  # مخفي تقنياً لكن نستخدمه داخلياً

    col_widths = {'الاسم الكامل': 280, 'الرصيد المحدث': 130, 'آخر إجازة': 350}
    for col in columns[1:]:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=col_widths.get(col, 150))

    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)

    # ---- إطار تسجيل الإجازة ----
    large_font = ("Arial", 12, "bold")
    deduct_frame = tk.LabelFrame(root, text="  تسجيل إجازة جديدة  ",
                                  font=("Arial", 13, "bold"), bg="#f0f2f5", pady=18, padx=20)
    deduct_frame.pack(fill="x", padx=10, pady=8)

    tk.Label(deduct_frame, text="عدد الأيام:", font=large_font, bg="#f0f2f5").grid(row=0, column=0, padx=5)
    entry_deduct = tk.Entry(deduct_frame, width=8, font=large_font)
    entry_deduct.grid(row=0, column=1, padx=5)

    tk.Label(deduct_frame, text="يوم:", font=large_font, bg="#f0f2f5").grid(row=0, column=2, padx=5)
    entry_day = tk.Entry(deduct_frame, width=5, font=large_font)
    entry_day.grid(row=0, column=3, padx=5)

    tk.Label(deduct_frame, text="شهر:", font=large_font, bg="#f0f2f5").grid(row=0, column=4, padx=5)
    entry_month = tk.Entry(deduct_frame, width=5, font=large_font)
    entry_month.grid(row=0, column=5, padx=5)

    tk.Label(deduct_frame, text="سنة:", font=large_font, bg="#f0f2f5").grid(row=0, column=6, padx=5)
    entry_year = tk.Entry(deduct_frame, width=8, font=large_font)
    entry_year.grid(row=0, column=7, padx=5)

    tk.Label(deduct_frame, text="ملاحظات:", font=large_font, bg="#f0f2f5").grid(row=0, column=8, padx=5)
    entry_notes = tk.Entry(deduct_frame, width=20, font=("Arial", 11))
    entry_notes.grid(row=0, column=9, padx=5)

    tk.Button(deduct_frame, text="✅ تأكيد الإجازة", command=deduct_vacation,
              bg="#dc3545", fg="white", font=("Arial", 12, "bold"),
              padx=20, pady=8, cursor="hand2").grid(row=0, column=10, padx=20)

    # ---- شريط الحالة السفلي ----
    status_bar = tk.Label(root, textvariable=status_var, bd=1, relief="sunken",
                          anchor="w", font=("Arial", 10), bg="#1a2a4a", fg="white")
    status_bar.pack(side="bottom", fill="x")

    # ---- تلوين صفوف الرصيد المنخفض ----
    tk.Label(root,
             text="🔴 الموظفون ذوو الرصيد ≤ 5 أيام مُلوّنون باللون الأحمر الفاتح",
             font=("Arial", 9), bg="#f0f2f5", fg="#666").pack(side="bottom")

    load_data()

# ========================================================
# 19. نقطة الانطلاق
# ========================================================
current_user = ""
current_role = ""

root = tk.Tk()
root.withdraw()  # إخفاء النافذة الرئيسية حتى يتم تسجيل الدخول

status_var = tk.StringVar()

init_db()
show_login_screen()

# إذا تم إغلاق نافذة الدخول بدون دخول
if current_user:
    root.mainloop()
else:
    root.destroy()