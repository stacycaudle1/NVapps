import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import database

# Color palette
ACCENT = '#0078d4'       # primary accent (blue)
WIN_BG = '#f6f9fc'       # window background
HEADER_BG = '#2b579a'    # header background
HEADER_FG = 'white'      # header foreground


def get_risk_color(score):
    if score >= 8:
        return 'red'
    elif score >= 4:
        return 'yellow'
    else:
        return 'green'

class AppTracker(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Business Application Tracker')
        self.geometry('1000x700')
        self.minsize(900, 500)
        # apply a modern ttk style
        self.setup_style()
        # window background
        try:
            self.configure(bg=WIN_BG)
        except Exception:
            pass
        database.init_db()  # Ensure DB schema is correct before anything else
        self.create_widgets()
        self.refresh_table()

    def setup_style(self):
        style = ttk.Style(self)
        # prefer 'clam' for consistent Treeview heading styling; fall back to native themes
        try:
            if 'clam' in style.theme_names():
                style.theme_use('clam')
            else:
                for theme in ('vista', 'xpnative', 'clam'):
                    try:
                        style.theme_use(theme)
                        break
                    except Exception:
                        pass
        except Exception:
            pass
        default_font = ('Segoe UI', 10)
        heading_font = ('Segoe UI', 11, 'bold')
        style.configure('.', font=default_font)
        style.configure('Treeview', rowheight=26, font=default_font, background='white', fieldbackground='white')
        # Stronger heading style for visibility
        style.configure('Treeview.Heading', font=heading_font, background=HEADER_BG, foreground=HEADER_FG, relief='raised', borderwidth=1, padding=(8,4))
        # Ensure mapping works across themes
        try:
            style.map('Treeview.Heading', background=[('active', HEADER_BG), ('!disabled', HEADER_BG)], foreground=[('active', HEADER_FG), ('!disabled', HEADER_FG)])
        except Exception:
            pass
        style.configure('TButton', padding=6)
        style.configure('TEntry', padding=4)
        style.map('TButton', foreground=[('active', '!disabled', 'black')])
        # frame and window backgrounds
        style.configure('TFrame', background=WIN_BG)
        # Primary, Secondary and Danger button styles (uniform mapping)
        style.configure('Primary.TButton', background=ACCENT, foreground='white')
        style.map('Primary.TButton', background=[('active', '!disabled', '#005a9e')], foreground=[('disabled', '#d0d0d0')])
        style.configure('Secondary.TButton', background='#e1e1e1', foreground='black')
        style.map('Secondary.TButton', background=[('active', '!disabled', '#cfcfcf')], foreground=[('disabled', '#a0a0a0')])
        style.configure('Danger.TButton', background='#f8d7da', foreground='#8b0000')
        style.map('Danger.TButton', background=[('active', '!disabled', '#f5c6cb')])

    def get_departments(self):
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        c.execute('SELECT id, name FROM departments')
        departments = c.fetchall()
        conn.close()
        return departments

    def add_department_popup(self):
        popup = tk.Toplevel(self)
        popup.title('Add Department')
        ttk.Label(popup, text='Department Name:').pack(padx=10, pady=5)
        dept_entry = ttk.Entry(popup)
        dept_entry.pack(padx=10, pady=5)

        def add_dept():
            dept_name = dept_entry.get().strip()
            if not dept_name:
                return
            conn = database.sqlite3.connect(database.DB_NAME)
            c = conn.cursor()
            c.execute('INSERT OR IGNORE INTO departments (name) VALUES (?)', (dept_name,))
            conn.commit()
            conn.close()
            # Update department_listbox with new departments
            self.department_listbox.delete(0, 'end')
            for dept_id, dept_name in self.get_departments():
                self.department_listbox.insert('end', dept_name)
            # Also update filter_combo values
            self.filter_combo['values'] = ['All'] + [d[1] for d in self.get_departments()]
            self.filter_combo.current(0)
            popup.destroy()

        ttk.Button(popup, text='Add', command=add_dept, style='Primary.TButton').pack(padx=10, pady=10)

    def show_report(self):
        report_win = tk.Toplevel(self)
        report_win.title('Department Risk Report')
        tree = ttk.Treeview(report_win, columns=('Department', 'App Count', 'Avg Risk', 'Status'), show='headings')
        for col in tree['columns']:
            tree.heading(col, text=col)
        tree.pack(fill='both', expand=True)
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        # Aggregate by department using many-to-many relationship
        c.execute('''SELECT d.name, COUNT(ad.app_id), AVG(a.risk_score)
                     FROM departments d
                     LEFT JOIN application_departments ad ON d.id = ad.dept_id
                     LEFT JOIN applications a ON ad.app_id = a.id
                     GROUP BY d.id''')
        for dept, count, avg_risk in c.fetchall():
            if avg_risk is None:
                status = 'No Data'
                color = 'white'
            elif avg_risk > 50:
                status = 'Critical'
                color = '#ffcccc'
            elif avg_risk > 30:
                status = 'Warning'
                color = '#fff2cc'
            else:
                status = 'Safe'
                color = '#ccffcc'
            tree.insert('', 'end', values=(dept, count, f'{avg_risk:.1f}' if avg_risk else '0', status), tags=(status,))
            tree.tag_configure(status, background=color)
        conn.close()

    def manage_departments_popup(self):
        popup = tk.Toplevel(self)
        popup.title('Manage Departments')
        dept_listbox = tk.Listbox(popup, selectmode='multiple', exportselection=0)
        departments = self.get_departments()
        for dept_id, dept_name in departments:
            dept_listbox.insert('end', dept_name)
        dept_listbox.pack(padx=10, pady=10)

        def delete_selected():
            selected_indices = dept_listbox.curselection()
            if not selected_indices:
                return
            # Re-fetch departments to get ids in order
            departments_current = self.get_departments()
            to_delete = [departments_current[i][0] for i in selected_indices]
            conn = database.sqlite3.connect(database.DB_NAME)
            c = conn.cursor()
            for dept_id in to_delete:
                c.execute('DELETE FROM application_departments WHERE dept_id = ?', (dept_id,))
                c.execute('DELETE FROM departments WHERE id = ?', (dept_id,))
            conn.commit()
            conn.close()
            # Update department_listbox and filter_combo after deletion
            self.department_listbox.delete(0, 'end')
            for dept_id, dept_name in self.get_departments():
                self.department_listbox.insert('end', dept_name)
            self.filter_combo['values'] = ['All'] + [d[1] for d in self.get_departments()]
            self.filter_combo.current(0)
            popup.destroy()
            self.refresh_table()

        ttk.Button(popup, text='Delete Selected', command=delete_selected, style='Danger.TButton').pack(padx=10, pady=10)

    def purge_database_gui(self):
        if messagebox.askyesno('Confirm Purge', 'Are you sure you want to delete ALL data? This cannot be undone.'):
            database.purge_database()
            self.department_listbox.delete(0, 'end')
            for dept_id, dept_name in self.get_departments():
                self.department_listbox.insert('end', dept_name)
            self.filter_combo['values'] = ['All'] + [d[1] for d in self.get_departments()]
            self.filter_combo.current(0)
            self.refresh_table()
            messagebox.showinfo('Purge Complete', 'All data has been deleted.')

    def create_widgets(self):
        # use a PanedWindow to separate form and table
        paned = ttk.Panedwindow(self, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=12, pady=12)

        # Left frame: form
        form_frame = ttk.Frame(paned, width=360)
        paned.add(form_frame, weight=0)

        # Right frame: table and controls
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)

        # Form content with consistent padding
        padx = 6
        pady = 6
        ttk.Label(form_frame, text='App Name').grid(row=0, column=0, sticky='w', padx=padx, pady=pady)
        self.name_entry = ttk.Entry(form_frame)
        self.name_entry.grid(row=0, column=1, sticky='ew', padx=padx, pady=pady)

        # factor entries (Inserted 'Installed' under Criticality)
        factor_labels = [
            ('Stability', 1),
            ('Need', 2),
            ('Criticality', 3),
            ('Installed', 4),
            ('DisasterRecovery', 5),
            ('Safety', 6),
            ('Security', 7),
            ('Monetary', 8),
            ('CustomerService', 9)
        ]
        self.factor_entries = {}
        for label, row in factor_labels:
            ttk.Label(form_frame, text=label).grid(row=row, column=0, sticky='w', padx=padx, pady=pady)
            entry = ttk.Entry(form_frame, width=12)
            entry.grid(row=row, column=1, sticky='w', padx=padx, pady=pady)
            self.factor_entries[label] = entry
        ttk.Label(form_frame, text='Related Vendor').grid(row=10, column=0, sticky='w', padx=padx, pady=pady)
        self.vendor_entry = ttk.Entry(form_frame)
        self.vendor_entry.grid(row=10, column=1, sticky='ew', padx=padx, pady=pady)

        ttk.Label(form_frame, text='Departments').grid(row=11, column=0, sticky='nw', padx=padx, pady=pady)
        self.department_listbox = tk.Listbox(form_frame, selectmode='multiple', exportselection=0, height=6)
        self.department_listbox.grid(row=11, column=1, sticky='ew', padx=padx, pady=pady)
        ttk.Button(form_frame, text='Add Dept', command=self.add_department_popup, style='Primary.TButton').grid(row=12, column=0, padx=padx, pady=pady)
        ttk.Button(form_frame, text='Manage Depts', command=self.manage_departments_popup, style='Secondary.TButton').grid(row=12, column=1, padx=padx, pady=pady)

        ttk.Button(form_frame, text='Add Application', command=self.add_application, style='Primary.TButton').grid(row=13, column=0, columnspan=2, sticky='ew', pady=(12,6), padx=padx)
        ttk.Button(form_frame, text='Purge Database', command=self.purge_database_gui, style='Danger.TButton').grid(row=14, column=0, columnspan=2, sticky='ew', padx=padx)

        # fill listbox
        self.department_listbox.delete(0, 'end')
        for dept_id, dept_name in self.get_departments():
            self.department_listbox.insert('end', dept_name)
        # Rating scale and scoring explanation shown below the purge button
        rating_help = (
            'Rating Scale (1-10)\n'
            'For each attribute, rate the application:\n'
            '• 1 = Very Low Impact/Importance\n'
            '• 10 = Extremely High Impact/Importance\n'
            'Example for Criticality:\n'
            '  • 1-3 = Minor/non-core tool\n'
            '  • 4-6 = Useful but replaceable\n'
            '  • 7-10 = Mission critical\n'
            '\n'
            'For each application:\n'
            'Total Score = ∑ (Rating × Weight)\n'
            '• Ratings are on a 1-10 scale.\n'
            '• Weights are percentages converted to decimals.\n'
            '• Maximum possible score = 10 × 1.00 = 10.0.'
        )
        # Use a small frame with a Text widget and vertical scrollbar so the help text can scroll
        rating_frame = ttk.Frame(form_frame)
        rating_frame.grid(row=15, column=0, columnspan=2, sticky='nsew', padx=padx, pady=(8,0))
        rating_text = tk.Text(rating_frame, wrap='word', height=8, width=40)
        rating_vsb = ttk.Scrollbar(rating_frame, orient='vertical', command=rating_text.yview)
        rating_text.configure(yscrollcommand=rating_vsb.set)
        rating_text.grid(row=0, column=0, sticky='nsew')
        rating_vsb.grid(row=0, column=1, sticky='ns')
        # Insert with tags for better readability (headings, bullets, indentation)
        rating_text.tag_configure('heading', font=('Segoe UI', 10, 'bold'))
        rating_text.tag_configure('bullet', lmargin1=12, lmargin2=24)
        rating_text.tag_configure('normal', font=('Segoe UI', 10))

        rating_text.insert('end', 'Rating Scale (1–10)\n', 'heading')
        rating_text.insert('end', '\nFor each attribute, rate the application:\n', 'normal')
        rating_text.insert('end', '• 1 = Very Low Impact/Importance\n', 'bullet')
        rating_text.insert('end', '• 10 = Extremely High Impact/Importance\n\n', 'bullet')

        rating_text.insert('end', 'Example for Criticality:\n', 'heading')
        rating_text.insert('end', '  • 1–3 = Minor/non-core tool\n', 'bullet')
        rating_text.insert('end', '  • 4–6 = Useful but replaceable\n', 'bullet')
        rating_text.insert('end', '  • 7–10 = Mission critical\n\n', 'bullet')

        rating_text.insert('end', 'For each application:\n', 'heading')
        rating_text.insert('end', 'Total Score = ∑ (Rating × Weight)\n', 'normal')
        rating_text.insert('end', '• Ratings are on a 1–10 scale.\n', 'bullet')
        rating_text.insert('end', '• Weights are percentages converted to decimals.\n', 'bullet')
        rating_text.insert('end', '• Maximum possible score = 10 × 1.00 = 10.0.\n', 'bullet')
        rating_text.configure(state='disabled')

        rating_frame.rowconfigure(0, weight=1)
        rating_frame.columnconfigure(0, weight=1)
        # add edit controls above the table
        edit_frame = ttk.Frame(right_frame)
        edit_frame.pack(fill='x', pady=(0,6))
        ttk.Button(edit_frame, text='Edit Selected', command=self.load_selected_for_edit, style='Secondary.TButton').pack(side='left', padx=6)
        ttk.Button(edit_frame, text='Save Changes', command=self.save_edit, style='Primary.TButton').pack(side='left', padx=6)

        # Right frame: filter + table
        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill='x', pady=(0,6))
        ttk.Label(control_frame, text='Filter by Department:').pack(side='left', padx=(0,6))
        self.filter_combo = ttk.Combobox(control_frame, values=['All'] + [d[1] for d in self.get_departments()], width=24)
        self.filter_combo.current(0)
        self.filter_combo.pack(side='left')
        ttk.Button(control_frame, text='Apply Filter', command=self.refresh_table, style='Secondary.TButton').pack(side='left', padx=6)
        ttk.Button(control_frame, text='Show Report', command=self.show_report, style='Primary.TButton').pack(side='left', padx=6)

        # table area with vertical and horizontal scrollbars
        table_frame = ttk.Frame(right_frame)
        table_frame.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(
            table_frame,
            columns=(
                'Name', 'Vendor', 'Stability', 'Need', 'Criticality', 'Installed',
                'DisasterRecovery', 'Safety', 'Security', 'Monetary', 'CustomerService',
                'Departments', 'Risk', 'Last Modified'
            ),
            show='headings'
        )
        # column sizing: choose sensible base widths and allow stretching
        cols = list(self.tree['columns'])
        base_widths = {
            'Name': 220, 'Vendor': 160, 'Stability': 80, 'Need': 80, 'Criticality': 80,
            'Installed': 80, 'DisasterRecovery': 120, 'Safety': 80, 'Security': 80,
            'Monetary': 80, 'CustomerService': 120, 'Departments': 160, 'Risk': 100, 'Last Modified': 150
        }
        # apply initial widths and make columns stretchable
        for col in cols:
            w = base_widths.get(col, 100)
            anchor = 'w' if col in ('Name', 'Vendor', 'Departments') else 'center'
            heading_text = 'DR' if col == 'DisasterRecovery' else col
            self.tree.heading(col, text=heading_text)
            self.tree.column(col, width=w, anchor=anchor, stretch=True)

        # vertical scrollbar only; horizontal scrolling removed to keep table fit
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Details area below the table: read-only Text widget showing selected row details
        details_text = tk.Text(table_frame, wrap='word', height=6, width=80)
        details_vsb = ttk.Scrollbar(table_frame, orient='vertical', command=details_text.yview)
        details_text.configure(yscrollcommand=details_vsb.set, state='disabled')
        details_text.grid(row=1, column=0, sticky='nsew', pady=(6,0))
        details_vsb.grid(row=1, column=1, sticky='ns', pady=(6,0))
        table_frame.rowconfigure(1, weight=0)
        # expose as attribute so handlers can update it
        self.details_text = details_text

        # Resize handler: distribute available width across columns proportionally
        def _resize_columns(event):
            try:
                total = event.width - 4  # small padding
                # ensure scrollbar width not included
                weights = [3, 2, 1, 1, 1, 1, 2, 1, 1, 1, 2, 2, 1]
                # fallback if counts differ
                if len(weights) != len(cols):
                    weights[:] = [1] * len(cols)
                s = sum(weights) or 1
                for i, col in enumerate(cols):
                    w = max(60, int(total * weights[i] / s))
                    self.tree.column(col, width=w)
            except Exception:
                pass

        table_frame.bind('<Configure>', _resize_columns)

        # populate details area when a row is selected
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)


    def load_selected_for_edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo('Edit', 'Please select a row to edit.')
            return
        item = sel[0]
        vals = self.tree.item(item, 'values')
        try:
            self.name_entry.delete(0, 'end')
            self.name_entry.insert(0, vals[0])
            self.vendor_entry.delete(0, 'end')
            self.vendor_entry.insert(0, vals[1])
            keys = ['Stability', 'Need', 'Criticality', 'Installed', 'DisasterRecovery', 'Safety', 'Security', 'Monetary', 'CustomerService']
            for i, key in enumerate(keys, start=2):
                entry = self.factor_entries.get(key)
                if entry is not None:
                    entry.delete(0, 'end')
                    entry.insert(0, vals[i])
        except Exception:
            messagebox.showerror('Error', 'Failed to load selected row for editing.')

    def save_edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo('Save', 'Please select a row to save.')
            return
        item = sel[0]
        # get app id from iid
        try:
            app_id = int(item)
        except Exception:
            messagebox.showerror('Save', 'Could not determine application id for selected row.')
            return
        name = self.name_entry.get().strip()
        vendor = self.vendor_entry.get().strip()
        fields = {}
        try:
            for key, entry in self.factor_entries.items():
                fields[key.lower()] = int(entry.get())
        except ValueError:
            messagebox.showerror('Error', 'All factor fields must be integers.')
            return
        fields['name'] = name
        fields['vendor'] = vendor
        # update db
        database.update_application(app_id, fields)
        self.refresh_table()
        messagebox.showinfo('Saved', 'Changes saved.')

    def add_application(self):
        name = self.name_entry.get()
        vendor = self.vendor_entry.get()
        factors = {}
        for label, entry in self.factor_entries.items():
            try:
                factors[label] = int(entry.get())
            except ValueError:
                messagebox.showerror('Error', f'{label} must be an integer.')
                return
        selected_indices = self.department_listbox.curselection()
        if not selected_indices:
            messagebox.showerror('Error', 'Please select at least one department.')
            return
        departments = self.get_departments()
        dept_ids = [departments[i][0] for i in selected_indices]
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        last_mod = datetime.utcnow().isoformat()
        # calculate risk score using the same scoring function
        ratings = {
            'Stability': factors['Stability'], 'Need': factors['Need'], 'Criticality': factors['Criticality'],
            'Installed': factors['Installed'], 'DisasterRecovery': factors['DisasterRecovery'], 'Safety': factors['Safety'],
            'Security': factors['Security'], 'Monetary': factors['Monetary'], 'CustomerService': factors['CustomerService']
        }
        try:
            score_result = database.score_application(ratings)
            risk_score = score_result.total
        except Exception:
            risk_score = 0
        c.execute('''INSERT INTO applications (
            name, vendor, stability, need, criticality, installed, disasterrecovery, safety, security, monetary, customerservice, risk_score, last_modified
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (name, vendor, factors['Stability'], factors['Need'], factors['Criticality'], factors['Installed'], factors['DisasterRecovery'],
             factors['Safety'], factors['Security'], factors['Monetary'], factors['CustomerService'], risk_score, last_mod))
        app_id = c.lastrowid
        conn.commit()
        conn.close()
        database.link_app_to_departments(app_id, dept_ids)
        self.refresh_table()

    def refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        filter_dept = None
        if hasattr(self, 'filter_combo'):
            filter_dept = self.filter_combo.get()
        c.execute('''SELECT id, name, vendor, stability, need, criticality, installed, disasterrecovery, safety, security, monetary, customerservice, last_modified FROM applications''')
        for app_row in c.fetchall():
            app_id = app_row[0]
            depts = database.get_app_departments(app_id)
            dept_str = ', '.join(depts)
            risk_score, risk_level = database.calculate_business_risk(app_row)
            # app_row indices: 1:name,2:vendor,3:stability,4:need,5:criticality,6:installed,7:disasterrecovery,8:safety,9:security,10:monetary,11:customerservice
            # Format last_modified as date-only (YYYY-MM-DD) for display
            last_mod_raw = app_row[12]
            last_mod_display = ''
            if last_mod_raw:
                try:
                    # stored as ISO timestamp; convert to date
                    last_mod_display = datetime.fromisoformat(last_mod_raw).date().isoformat()
                except Exception:
                    # fallback: take date part before 'T' if present
                    last_mod_display = str(last_mod_raw).split('T')[0]

            row = (
                app_row[1], app_row[2], app_row[3], app_row[4], app_row[5], app_row[6], app_row[7], app_row[8], app_row[9], app_row[10], app_row[11], dept_str, f"{risk_score} ({risk_level})", last_mod_display
            )
            if filter_dept and filter_dept != 'All' and filter_dept not in depts:
                continue
            color = get_risk_color(risk_score)
            # store the app id as the item iid so we can identify it later
            self.tree.insert('', 'end', iid=str(app_id), values=row, tags=(color,))
        self.tree.tag_configure('red', background='#ffcccc')
        self.tree.tag_configure('yellow', background='#fff2cc')
        self.tree.tag_configure('green', background='#ccffcc')
        conn.close()

    def on_tree_select(self, event):
        # called when a row is selected; fill the details_text with a readable summary
        try:
            sel = self.tree.selection()
            if not sel:
                # clear details
                self.details_text.configure(state='normal')
                self.details_text.delete('1.0', 'end')
                self.details_text.configure(state='disabled')
                return
            item = sel[0]
            vals = self.tree.item(item, 'values')
            # build a simple labeled summary
            headers = list(self.tree['columns'])
            lines = []
            for h, v in zip(headers, vals):
                lines.append(f"{h}: {v}")
            text = '\n'.join(lines)
            self.details_text.configure(state='normal')
            self.details_text.delete('1.0', 'end')
            self.details_text.insert('end', text)
            self.details_text.configure(state='disabled')
        except Exception:
            pass

if __name__ == '__main__':
    app = AppTracker()
    app.mainloop()
