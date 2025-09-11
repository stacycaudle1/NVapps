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
        # Open the application in full-screen mode
        self.state('zoomed')
        # apply a modern ttk style
        self.setup_style()
        # window background
        try:
            self.configure(bg=WIN_BG)
        except Exception:
            pass
        database.init_db()  # Ensure DB schema is correct before anything else
        self.create_widgets()
        # Populate the department listbox after creating widgets
        print("DEBUG: Initializing departments")  # Debug log
        departments = self.get_departments()
        print(f"DEBUG: Got departments: {departments}")  # Debug log
        self.department_listbox.delete(0, 'end')
        for dept_id, dept_name in departments:
            print(f"DEBUG: Inserting department: {dept_name}")  # Debug log
            self.department_listbox.insert('end', dept_name)
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
        # Smaller font for table rows to ensure text fits in cells
        tree_font = ('Segoe UI', 8)
        heading_font = ('Segoe UI', 10, 'bold')
        style.configure('.', font=default_font)
        style.configure('Treeview', rowheight=28, font=tree_font, background='white', fieldbackground='white')
        # Stronger heading style for visibility
        style.configure('Treeview.Heading', font=heading_font, background=HEADER_BG, foreground=HEADER_FG, relief='flat', borderwidth=0, padding=(8,4))
        # Ensure mapping works across themes
        try:
            style.map('Treeview.Heading', background=[('active', HEADER_BG), ('!disabled', HEADER_BG)], foreground=[('active', HEADER_FG), ('!disabled', HEADER_FG)])
        except Exception:
            pass
        style.configure('TButton', padding=8, font=('Segoe UI', 10))
        style.configure('TEntry', padding=6, font=('Segoe UI', 10))
        style.map('TButton', foreground=[('active', '!disabled', 'black')])
        # frame and window backgrounds
        style.configure('TFrame', background=WIN_BG)
        # Primary, Secondary and Danger button styles (uniform mapping)
        style.configure('Primary.TButton', background=ACCENT, foreground='white', borderwidth=0)
        style.map('Primary.TButton', background=[('active', '!disabled', '#005a9e')], foreground=[('disabled', '#d0d0d0')])
        style.configure('Secondary.TButton', background='#e1e1e1', foreground='black', borderwidth=0)
        style.map('Secondary.TButton', background=[('active', '!disabled', '#cfcfcf')], foreground=[('disabled', '#a0a0a0')])
        style.configure('Danger.TButton', background='#f8d7da', foreground='#8b0000', borderwidth=0)
        style.map('Danger.TButton', background=[('active', '!disabled', '#f5c6cb')])
        # Green button style for note functions
        style.configure('Success.TButton', background='#90EE90', foreground='black', borderwidth=0)  # Light green
        style.map('Success.TButton', background=[('active', '!disabled', '#7CCD7C')], foreground=[('disabled', '#a0a0a0')])

    def get_departments(self):
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        c.execute('SELECT id, name FROM business_units ORDER BY name ASC')
        departments = c.fetchall()
        conn.close()
        print(f"DEBUG: Retrieved departments: {departments}")  # Debugging log
        return departments

    def add_department_popup(self):
        popup = tk.Toplevel(self)
        popup.title('Add Department')
        ttk.Label(popup, text='Business Unit:').pack(padx=10, pady=5)
        dept_entry = ttk.Entry(popup)
        dept_entry.pack(padx=10, pady=5)

        def add_dept():
            dept_name = dept_entry.get().strip()
            if not dept_name:
                return
            conn = database.sqlite3.connect(database.DB_NAME)
            c = conn.cursor()
            c.execute('INSERT OR IGNORE INTO business_units (name) VALUES (?)', (dept_name,))
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

    def report_sort_table(self, tree, col, reverse):
        # Sort function specifically for the report window
        try:
            # Get data for sorting
            data = [(tree.set(k, col), k) for k in tree.get_children('')]
            
            # Sort the data based on appropriate type conversion
            data.sort(reverse=reverse, key=lambda t: self._convert_report_value(t[0]))
            
            # Rearrange items in the tree
            for index, (val, k) in enumerate(data):
                tree.move(k, '', index)
                
            # Update the heading to show sort direction and set command for next sort
            sort_direction = "▼" if reverse else "▲"  # Down arrow for descending, up for ascending
            
            # Get original heading text without any arrows
            heading_text = tree.heading(col, 'text').replace("▲", "").replace("▼", "").strip()
            
            # Reset all column headings to remove any previous sort indicators
            for column in tree['columns']:
                if column != col:  # Skip the column we're currently sorting
                    current_text = tree.heading(column, 'text').replace("▲", "").replace("▼", "").strip()
                    tree.heading(column, text=current_text)
            
            # Set new heading with sort indicator
            tree.heading(col, text=f"{heading_text} {sort_direction}", 
                        command=lambda: self.report_sort_table(tree, col, not reverse))
                
        except Exception as e:
            messagebox.showerror('Error', f'Failed to sort report table: {e}')
    
    def _convert_report_value(self, value):
        # Helper to convert values for sorting in reports
        if not value:
            return ""  # Empty values sort first
        
        # Try to convert to number
        try:
            return float(value)
        except ValueError:
            pass
        
        # Default: return lowercase string for case-insensitive sorting
        return value.lower() if isinstance(value, str) else value
    
    def show_report(self):
        report_win = tk.Toplevel(self)
        report_win.title('Business Unit Risk Report')
        tree = ttk.Treeview(report_win, columns=('Business Unit', 'App Count', 'Avg Risk', 'Status'), show='headings')
        for col in tree['columns']:
            tree.heading(col, text=col, command=lambda c=col: self.report_sort_table(tree, c, False))
        tree.pack(fill='both', expand=True)
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        # Aggregate by business unit using the correct business_units relationship
        c.execute('''SELECT bu.name, COUNT(abu.app_id), AVG(a.risk_score)
                     FROM business_units bu
                     LEFT JOIN application_business_units abu ON bu.id = abu.unit_id
                     LEFT JOIN applications a ON abu.app_id = a.id
                     GROUP BY bu.id''')
        for bu_name, count, avg_risk in c.fetchall():
            if avg_risk is None:
                status = 'No Data'
                color = 'white'
            elif avg_risk > 8:  # Adjusted threshold to match risk_color function
                status = 'Critical'
                color = '#ffcccc'
            elif avg_risk > 4:  # Adjusted threshold to match risk_color function
                status = 'Warning'
                color = '#fff2cc'
            else:
                status = 'Safe'
                color = '#ccffcc'
            tree.insert('', 'end', values=(bu_name, count, f'{avg_risk:.1f}' if avg_risk else '0', status), tags=(status,))
            tree.tag_configure(status, background=color)
            
        # Add styling to the report tree
        tree_style = ttk.Style()
        tree_style.configure('Treeview', rowheight=28, font=('Segoe UI', 9))
        tree_style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        
        # Add a scrollbar to the report window
        vsb = ttk.Scrollbar(report_win, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        
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
                c.execute('DELETE FROM application_business_units WHERE unit_id = ?', (dept_id,))
                c.execute('DELETE FROM business_units WHERE id = ?', (dept_id,))
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
        # Initialize filter_combo early with a default ttk.Combobox instance
        self.filter_combo = ttk.Combobox(values=['All'], width=24)
        self.filter_combo.current(0)

        # use a PanedWindow to separate form and table
        paned = ttk.Panedwindow(self, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=16, pady=16)

        form_frame = ttk.Frame(paned, width=360, padding=12)
        paned.add(form_frame, weight=0)

        right_frame = ttk.Frame(paned, padding=12)
        paned.add(right_frame, weight=1)

        # Form content with consistent padding
        padx = 8
        pady = 8
        ttk.Label(form_frame, text='App Name', font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=padx, pady=pady)
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

        ttk.Label(form_frame, text='Business Unit').grid(row=11, column=0, sticky='nw', padx=padx, pady=pady)
        
        # Create a frame to hold both buttons vertically
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=11, column=0, sticky='nw', padx=padx, pady=(40,0))
        
        # Add Business Unit button - with left alignment (anchor='w')
        ttk.Button(buttons_frame, text='Add Business Unit', command=self.add_department_popup, style='Primary.TButton').pack(pady=5, anchor='w', fill='x')
        
        # Manage Business Units button
        ttk.Button(buttons_frame, text='Manage Business Units', command=self.manage_departments_popup, style='Primary.TButton').pack(pady=5)
        
        dept_frame = ttk.Frame(form_frame)
        dept_frame.grid(row=11, column=1, sticky='nsew', padx=padx, pady=pady)
        
        # Set minimum size for the department frame
        dept_frame.grid_propagate(False)
        dept_frame.configure(width=300, height=200)  # Set explicit size

        # Create department listbox with white background and increased height
        self.department_listbox = tk.Listbox(dept_frame, selectmode='multiple', exportselection=0, 
                                           height=15, width=40, bg='white', font=('Segoe UI', 10))
        self.department_listbox.pack(side='left', fill='both', expand=True)

        dept_scrollbar = ttk.Scrollbar(dept_frame, orient='vertical', command=self.department_listbox.yview)
        dept_scrollbar.pack(side='right', fill='y')
        self.department_listbox.configure(yscrollcommand=dept_scrollbar.set)
        
        # Ensure the department frame expands properly
        form_frame.grid_columnconfigure(1, weight=1)
        dept_frame.grid_configure(pady=(pady, pady*2))  # Add extra padding below
        dept_frame.pack_propagate(False)

        # Add Application button (now in row 13 since we removed the separate Manage Business Units button)
        ttk.Button(form_frame, text='Add Application', command=self.add_application, style='Primary.TButton').grid(row=13, column=0, columnspan=2, sticky='ew', pady=(12,6), padx=padx)

        # Create a frame for top-right buttons
        top_buttons_frame = ttk.Frame(right_frame)
        top_buttons_frame.pack(side='top', fill='x', pady=6)
        
        # Edit Selected and Save Changes buttons at the top with Primary style
        ttk.Button(top_buttons_frame, text='Edit Selected', command=self.load_selected_for_edit, style='Primary.TButton').pack(side='left', padx=6)
        ttk.Button(top_buttons_frame, text='Save Changes', command=self.save_edit, style='Primary.TButton').pack(side='left', padx=6)
        
        # Move 'Purge Database' button to the top right-hand side of the screen
        purge_button = ttk.Button(top_buttons_frame, text='Purge Database', command=self.purge_database_gui, style='Danger.TButton')
        purge_button.pack(side='right', padx=6)
        
        # Create Note and Save Note buttons in a separate frame below
        edit_frame = ttk.Frame(right_frame)
        edit_frame.pack(fill='x', pady=(0,6))
        ttk.Button(edit_frame, text='Create Note', command=self.create_note, style='Success.TButton').pack(side='left', padx=6)
        ttk.Button(edit_frame, text='Save Note', command=self.save_notes, style='Success.TButton').pack(side='left', padx=6)
        # Add 'Delete Selected' button 
        ttk.Button(edit_frame, text='Delete Selected', command=self.delete_selected_app, style='Danger.TButton').pack(side='left', padx=6)

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
                'Business Unit', 'Risk', 'Last Modified'
            ),
            show='headings'
        )
        # column sizing: adjusted widths to ensure text fits better in cells
        cols = list(self.tree['columns'])
        base_widths = {
            'Name': 220, 'Vendor': 160, 'Stability': 80, 'Need': 80, 'Criticality': 90,
            'Installed': 80, 'DisasterRecovery': 100, 'Safety': 80, 'Security': 80,
            'Monetary': 80, 'CustomerService': 140, 'Business Unit': 180, 'Risk': 120, 'Last Modified': 150
        }
        # apply initial widths and make columns stretchable with improved text display
        for col in cols:
            w = base_widths.get(col, 100)
            # Left-align text columns, center-align numeric columns
            anchor = 'w' if col in ('Name', 'Vendor', 'Business Unit') else 'center'
            # Use abbreviations for wider column names to help with spacing
            if col == 'DisasterRecovery':
                heading_text = 'DR'
            elif col == 'CustomerService':
                heading_text = 'Cust Service'
            else:
                heading_text = col
            # Fixed command lambda to properly capture column variable
            self.tree.heading(col, text=heading_text, command=lambda c=col: self.sort_table(c, False))
            self.tree.column(col, width=w, anchor=anchor, stretch=True, minwidth=w)

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
            # Clear all current selections
            self.department_listbox.selection_clear(0, tk.END)
            
            # Get the departments for this application
            app_id = int(item)
            app_departments = database.get_app_departments(app_id)
            
            # Select departments in the listbox
            for i in range(self.department_listbox.size()):
                dept_name = self.department_listbox.get(i)
                if dept_name in app_departments:
                    self.department_listbox.selection_set(i)
            
            # Load other fields
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
        except Exception as e:
            messagebox.showerror('Error', f'Failed to load selected row for editing: {e}')

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
        
        # Clear form fields after saving
        self.name_entry.delete(0, 'end')
        self.vendor_entry.delete(0, 'end')
        for entry in self.factor_entries.values():
            entry.delete(0, 'end')
        self.department_listbox.selection_clear(0, tk.END)
        
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
        last_mod = datetime.utcnow().isoformat()

        ratings = {
            'Stability': factors['Stability'], 'Need': factors['Need'], 'Criticality': factors['Criticality'],
            'Installed': factors['Installed'], 'DisasterRecovery': factors['DisasterRecovery'], 'Safety': factors['Safety'],
            'Security': factors['Security'], 'Monetary': factors['Monetary'], 'CustomerService': factors['CustomerService']
        }

        conn = None
        app_id = None
        try:
            conn = database.sqlite3.connect(database.DB_NAME)
            c = conn.cursor()

            try:
                score_result = database.score_application(ratings)
                risk_score = score_result.total
            except Exception:
                risk_score = 0

            # Create a separate application entry for each department
            app_ids = []
            for dept_id in dept_ids:
                # Create new application entry
                c.execute('''INSERT INTO applications (
                    name, vendor, stability, need, criticality, installed, disasterrecovery, safety, security, monetary, customerservice, risk_score, last_modified
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                    (name, vendor, factors['Stability'], factors['Need'], factors['Criticality'], factors['Installed'], factors['DisasterRecovery'],
                     factors['Safety'], factors['Security'], factors['Monetary'], factors['CustomerService'], risk_score, last_mod))
                
                app_id = c.lastrowid
                app_ids.append(app_id)
                
                # Create business unit link for this application
                c.execute('INSERT INTO application_business_units (app_id, unit_id) VALUES (?, ?)', 
                         (app_id, dept_id))
            
            conn.commit()
        except Exception as e:
            messagebox.showerror('Error', f'Failed to add application: {e}')
            return
        finally:
            if conn:
                conn.close()

        if not app_ids:
            messagebox.showerror('Error', 'Failed to create application entries.')

        self.name_entry.delete(0, tk.END)
        self.vendor_entry.delete(0, tk.END)
        for entry in self.factor_entries.values():
            entry.delete(0, tk.END)
        self.department_listbox.selection_clear(0, tk.END)
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
        rows = []
        for app_row in c.fetchall():
            app_id = app_row[0]
            depts = database.get_app_departments(app_id)
            dept_str = ', '.join(depts)
            risk_score, risk_level = 0, 'Unknown'  # Initialize risk_score and risk_level
            try:
                risk_score, risk_level = database.calculate_business_risk(app_row)
            except Exception:
                pass
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
            rows.append((row, app_id, risk_score))
        # Sort rows by the 'Name' column (index 0)
        rows.sort(key=lambda x: x[0][0].lower())
        for row, app_id, risk_score in rows:
            color = get_risk_color(risk_score)
            # store the app id as the item iid so we can identify it later
            self.tree.insert('', 'end', iid=str(app_id), values=row, tags=(color,))
        self.tree.tag_configure('red', background='#ffcccc')
        self.tree.tag_configure('yellow', background='#fff2cc')
        self.tree.tag_configure('green', background='#ccffcc')
        conn.close()

    def on_tree_select(self, event):
        # called when a row is selected; fill the details_text with notes only
        try:
            sel = self.tree.selection()
            if not sel:
                # clear details
                self.details_text.configure(state='normal')
                self.details_text.delete('1.0', 'end')
                self.details_text.configure(state='disabled')
                return
            item = sel[0]
            # store selected app id (we set iid to app id when inserting)
            try:
                self.selected_app_id = int(item)
            except Exception:
                self.selected_app_id = None
            # fetch notes from DB if possible
            notes = None
            if self.selected_app_id is not None:
                app_row = database.get_application(self.selected_app_id)
                if app_row is not None and len(app_row) > 12:
                    notes = app_row[12]
            # display notes or default text
            self.details_text.configure(state='normal')
            self.details_text.delete('1.0', 'end')
            if notes:
                self.details_text.insert('end', notes)
            else:
                self.details_text.insert('end', 'No notes')
            self.details_text.configure(state='disabled')
        except Exception:
            pass

    def create_note(self):
        # Allow appending a note for the currently selected app
        if not hasattr(self, 'selected_app_id') or self.selected_app_id is None:
            messagebox.showinfo('Create Note', 'Please select an application to create a note for.')
            return
        try:
            self.details_text.configure(state='normal')
            existing_notes = self.details_text.get('1.0', 'end').strip()
            if existing_notes == 'No notes':
                existing_notes = ''
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            new_note = f"[{timestamp}]\n"
            self.details_text.delete('1.0', 'end')  # Clear existing content
            self.details_text.insert('1.0', f"{new_note}\n{existing_notes}")  # Append new note at the top
            self.details_text.mark_set('insert', f'1.{len(new_note)}')  # Place cursor after the timestamp
            self.details_text.focus_set()
        except Exception:
            messagebox.showerror('Error', 'Failed to prepare note creation.')

    def save_notes(self):
        # Save notes text into DB for selected app
        if not hasattr(self, 'selected_app_id') or self.selected_app_id is None:
            messagebox.showinfo('Save Note', 'Please select an application to save notes for.')
            return
        try:
            notes_text = self.details_text.get('1.0', 'end').strip()
            app_id = self.selected_app_id  # Store the app_id before refresh
            
            # Update notes via database helper
            database.update_application(app_id, {'notes': notes_text})
            
            # Make read-only again
            self.details_text.configure(state='disabled')
            
            # Refresh table to reflect last_modified change
            self.refresh_table()
            
            # Re-select the previously selected application
            for item in self.tree.get_children():
                if int(item) == app_id:
                    self.tree.selection_set(item)
                    self.tree.see(item)  # Ensure the selected item is visible
                    break
                    
            messagebox.showinfo('Saved', 'Notes saved.')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save notes: {e}')

    def sort_table(self, col, reverse):
        # Sort the table by the given column
        try:
            # Store current selection to restore after sorting
            current_selection = self.tree.selection()
            selected_ids = [int(item) for item in current_selection] if current_selection else []
            
            # Get data for sorting
            data = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
            
            # Sort the data based on appropriate type conversion
            data.sort(reverse=reverse, key=lambda t: self._convert_sort_value(t[0]))
            
            # Rearrange items in the tree
            for index, (val, k) in enumerate(data):
                self.tree.move(k, '', index)
            
            # Update the heading to show sort direction and set command for next sort
            sort_direction = "▼" if reverse else "▲"  # Down arrow for descending, up for ascending
            
            # Get original heading text without any arrows
            heading_text = self.tree.heading(col, 'text').replace("▲", "").replace("▼", "").strip()
            
            # Reset all column headings to remove any previous sort indicators
            for column in self.tree['columns']:
                if column != col:  # Skip the column we're currently sorting
                    current_text = self.tree.heading(column, 'text').replace("▲", "").replace("▼", "").strip()
                    self.tree.heading(column, text=current_text)
            
            # Set new heading with sort indicator
            self.tree.heading(col, text=f"{heading_text} {sort_direction}", 
                              command=lambda: self.sort_table(col, not reverse))
            
            # Restore selection if there was one
            if selected_ids:
                for item in self.tree.get_children():
                    if int(item) in selected_ids:
                        self.tree.selection_add(item)
                        self.tree.see(item)  # Make the first selected item visible
                        break
                
        except Exception as e:
            messagebox.showerror('Error', f'Failed to sort table: {e}')

    def _convert_sort_value(self, value):
        # Helper to convert values for sorting (e.g., numeric, date, or string)
        if not value:
            return ""  # Empty values sort first
            
        # Handle risk score format like "8 (High)"
        if isinstance(value, str) and "(" in value and ")" in value:
            try:
                # Extract the numeric part before the parentheses
                numeric_part = value.split('(')[0].strip()
                return float(numeric_part)
            except ValueError:
                pass
        
        # Try to convert to number
        try:
            return float(value)
        except ValueError:
            pass
            
        # Try to convert to date
        try:
            if isinstance(value, str) and '-' in value:
                return datetime.fromisoformat(value)
        except ValueError:
            pass
            
        # Default: return lowercase string for case-insensitive sorting
        return value.lower() if isinstance(value, str) else value

    def refresh_departments(self):
        # Refresh the department listbox with current data
        self.department_listbox.delete(0, 'end')
        departments = self.get_departments()
        for dept_id, dept_name in departments:
            self.department_listbox.insert('end', dept_name)

    def delete_selected_app(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo('Delete', 'Please select an application to delete.')
            return
        item = sel[0]
        # get app id from iid
        try:
            app_id = int(item)
        except Exception:
            messagebox.showerror('Delete', 'Could not determine application id for selected row.')
            return
        if messagebox.askyesno('Confirm Delete', 'Are you sure you want to delete the selected application? This cannot be undone.'):
            conn = database.sqlite3.connect(database.DB_NAME)
            c = conn.cursor()
            # Remove from application_departments first to avoid foreign key constraint
            c.execute('DELETE FROM application_departments WHERE app_id = ?', (app_id,))
            # Now remove the application
            c.execute('DELETE FROM applications WHERE id = ?', (app_id,))
            conn.commit()
            conn.close()
            self.refresh_table()
            messagebox.showinfo('Deleted', 'Application deleted.')

if __name__ == '__main__':
    app = AppTracker()
    app.mainloop()
