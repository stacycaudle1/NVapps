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
        # Initialize search variables
        self.search_entry = None
        self.search_type_var = None
        self.search_after_id = None  # For delayed search
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
            
        # Configure notebook (tabs) styling with blue colors
        style.configure('TNotebook', background=ACCENT)
        style.configure('TNotebook.Tab', background=HEADER_BG, foreground='white', padding=[10, 4])
        # Configure active tab
        style.map('TNotebook.Tab', background=[('selected', ACCENT)], foreground=[('selected', 'white')])
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
            # No need to update filter_combo anymore as we're using search functionality
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
        report_win.configure(bg=WIN_BG)
        
        # Add a modern header to the report
        header_frame = tk.Frame(report_win, bg=WIN_BG)
        header_frame.pack(fill='x', pady=(10, 15))
        header_label = tk.Label(header_frame, text='Business Unit Risk Overview', 
                              font=('Segoe UI', 14, 'bold'), fg=ACCENT, bg=WIN_BG)
        header_label.pack(side='left', padx=15)
        
        tree = ttk.Treeview(report_win, columns=('Business Unit', 'App Count', 'Avg Risk', 'Status'), show='headings')
        for col in tree['columns']:
            tree.heading(col, text=col, command=lambda c=col: self.report_sort_table(tree, c, False))
        tree.pack(fill='both', expand=True, padx=15, pady=(0, 15))
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
            # Update department_listbox after deletion
            self.department_listbox.delete(0, 'end')
            for dept_id, dept_name in self.get_departments():
                self.department_listbox.insert('end', dept_name)
            popup.destroy()
            self.refresh_table()

        ttk.Button(popup, text='Delete Selected', command=delete_selected, style='Danger.TButton').pack(padx=10, pady=10)

    def on_tab_changed(self, event):
        # Get the selected tab index
        tab_id = self.tab_control.select()
        tab_index = self.tab_control.index(tab_id)
        
        # Handle tab-specific actions (can be expanded in the future)
        if tab_index == 0:  # Application Risk Assessment tab
            # Refresh data when returning to the main tab
            self.refresh_table()
        elif tab_index == 1:  # Reports tab
            # Future functionality for reports tab can be added here
            pass

    # Placeholder methods are now directly in create_widgets
            
    def delayed_search(self):
        """Perform the search after a short delay to prevent excessive refreshes while typing"""
        # Cancel any existing delayed search
        if hasattr(self, 'search_after_id') and self.search_after_id is not None:
            self.after_cancel(self.search_after_id)
            self.search_after_id = None
            
        # Only search if not showing the placeholder text
        if hasattr(self, 'search_entry') and self.search_entry is not None:
            search_text = self.search_entry.get()
            if search_text != "Type to search...":
                # Update suggestions as user types
                self.update_search_suggestions()
                # Then refresh the table with current search
                self.refresh_table()
    
    def clear_search(self):
        """Clear the search entry and refresh the table to show all results"""
        if hasattr(self, 'search_entry') and self.search_entry is not None:
            self.search_entry.set("")  # For Combobox, use set instead of delete/insert
            self.search_entry.insert(0, "Type to search...")
            self.search_entry.config(foreground='gray')
        self.refresh_table()
            
    def purge_database_gui(self):
        if messagebox.askyesno('Confirm Purge', 'Are you sure you want to delete ALL data? This cannot be undone.'):
            database.purge_database()
            self.department_listbox.delete(0, 'end')
            for dept_id, dept_name in self.get_departments():
                self.department_listbox.insert('end', dept_name)
            self.refresh_table()
            messagebox.showinfo('Purge Complete', 'All data has been deleted.')

    def create_widgets(self):
        # No need to initialize filter_combo early anymore as we're using search functionality
        
        # Create the tab control
        self.tab_control = ttk.Notebook(self)
        self.tab_control.pack(fill='both', expand=True, padx=8, pady=8)
        
        # Bind tab change event
        self.tab_control.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Create the first tab - Application Risk Assessment
        self.tab_risk = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.tab_risk, text="Application Risk Assessment")
        
        # Create a second tab as a placeholder
        self.tab_reports = ttk.Frame(self.tab_control, style='TFrame')
        self.tab_control.add(self.tab_reports, text="Reports")
        
        # Create a container for the top section of the Reports tab
        reports_top_frame = ttk.Frame(self.tab_reports)
        reports_top_frame.pack(fill='x', pady=10, padx=15)
        
        # Add Show Report button to the top left of the Reports tab
        show_report_btn = ttk.Button(reports_top_frame, text='Show Report', command=self.show_report, style='Primary.TButton')
        show_report_btn.pack(side='left', padx=0)
        
        # Add some placeholder content to the Reports tab
        reports_label = tk.Label(self.tab_reports, text="Reports Dashboard", 
                               font=('Segoe UI', 16, 'bold'), fg=ACCENT, bg=WIN_BG)
        reports_label.pack(pady=20)
        
        info_label = tk.Label(self.tab_reports, text="Click 'Show Report' to view the Business Unit Risk Report.",
                            font=('Segoe UI', 12), fg='#555555', bg=WIN_BG)
        info_label.pack(pady=10)
        
        # use a PanedWindow to separate form and table inside the first tab
        paned = ttk.Panedwindow(self.tab_risk, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=16, pady=16)

        form_frame = ttk.Frame(paned, width=360, padding=12)
        paned.add(form_frame, weight=0)

        right_frame = ttk.Frame(paned, padding=12)
        paned.add(right_frame, weight=1)

        # Define label style for modern field labels
        self.form_label_style = {'font': ('Segoe UI', 10, 'bold'), 'fg': ACCENT, 'bg': WIN_BG, 'padx': 10}
        
        # Form content with consistent padding
        padx = 8
        pady = 8
        # Define category headers style for grouping
        self.category_style = {'font': ('Segoe UI', 11, 'bold'), 'fg': '#333333', 'bg': WIN_BG, 'pady': 5}

        # Create a category header for Rating Factors - moved above system label
        ratings_header = tk.Label(form_frame, text="Rating Factors", **self.category_style)
        ratings_header.grid(row=0, column=0, sticky='w', pady=(0, 10))
        
        app_name_label = tk.Label(form_frame, text='System', **self.form_label_style)
        app_name_label.grid(row=1, column=0, sticky='w', padx=padx, pady=pady)
        self.name_entry = ttk.Entry(form_frame)
        self.name_entry.grid(row=1, column=1, sticky='ew', padx=padx, pady=pady)
        
        # factor entries with modern styling
        factor_labels = [
            ('Stability', 2),
            ('Need', 3),
            ('Criticality', 4),
            ('Installed', 5),
            ('DisasterRecovery', 6),
            ('Safety', 7),
            ('Security', 8),
            ('Monetary', 9),
            ('CustomerService', 10)
        ]
        
        # Add a subtle separator below the Rating Factors header
        ttk.Separator(form_frame, orient='horizontal').grid(row=0, column=0, columnspan=2, sticky='ew', pady=(30, 5))
        
        # Add the factor entries with modern labels
        self.factor_entries = {}
        for label, row in factor_labels:
            factor_label = tk.Label(form_frame, text=label, **self.form_label_style)
            factor_label.grid(row=row, column=0, sticky='w', padx=padx, pady=pady)
            entry = ttk.Entry(form_frame, width=12)
            entry.grid(row=row, column=1, sticky='w', padx=padx, pady=pady)
            self.factor_entries[label] = entry
            
        vendor_label = tk.Label(form_frame, text='Related Vendor', **self.form_label_style)
        vendor_label.grid(row=11, column=0, sticky='w', padx=padx, pady=pady)
        self.vendor_entry = ttk.Entry(form_frame)
        self.vendor_entry.grid(row=11, column=1, sticky='ew', padx=padx, pady=pady)

        # Create a section header for Business Units - moved to left column
        bu_header = tk.Label(form_frame, text="Select Business Units", **self.category_style)
        bu_header.grid(row=12, column=0, sticky='nw', padx=padx, pady=(0, 2))
        
        # Add a subtle separator - moved to match the header position
        ttk.Separator(form_frame, orient='horizontal').grid(row=12, column=0, columnspan=2, sticky='ew', pady=(25, 5))
        
        # Create a frame to hold both buttons vertically - moved to row 13, left of business units textbox
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=13, column=0, sticky='nw', padx=padx, pady=pady)
        
        # Add Business Unit button - with left alignment (anchor='w')
        ttk.Button(buttons_frame, text='Add Business Unit', command=self.add_department_popup, style='Primary.TButton').pack(pady=5, anchor='w', fill='x')
        
        # Manage Business Units button
        ttk.Button(buttons_frame, text='Manage Business Units', command=self.manage_departments_popup, style='Primary.TButton').pack(pady=5)
        
        dept_frame = ttk.Frame(form_frame)
        dept_frame.grid(row=13, column=1, sticky='nsew', padx=padx, pady=pady)
        
        # Set minimum size for the department frame
        dept_frame.grid_propagate(False)
        dept_frame.configure(width=300, height=200)  # Set explicit size

        # Create department listbox with white background and increased height
        self.department_listbox = tk.Listbox(dept_frame, selectmode='multiple', exportselection=0, 
                                           height=15, width=40, bg='white', font=('Segoe UI', 10),
                                           highlightthickness=1, highlightcolor=ACCENT, borderwidth=1)
        self.department_listbox.pack(side='left', fill='both', expand=True)

        dept_scrollbar = ttk.Scrollbar(dept_frame, orient='vertical', command=self.department_listbox.yview)
        dept_scrollbar.pack(side='right', fill='y')
        self.department_listbox.configure(yscrollcommand=dept_scrollbar.set)
        
        # Ensure the department frame expands properly
        form_frame.grid_columnconfigure(1, weight=1)
        dept_frame.grid_configure(pady=(pady, pady*2))  # Add extra padding below
        dept_frame.pack_propagate(False)

        # Add Application button - moved to row 14 since we've adjusted the layout
        ttk.Button(form_frame, text='Add Application', command=self.add_application, style='Primary.TButton').grid(row=14, column=0, columnspan=2, sticky='ew', pady=(20,6), padx=padx)

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

        # Right frame: search + filter + table with modern styling
        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill='x', pady=(0,6))
        
        # Create an improved search frame with better layout
        search_frame = ttk.Frame(control_frame)
        search_frame.pack(side='left', fill='x', expand=True)
        
        filter_label_style = {'font': ('Segoe UI', 10), 'fg': ACCENT, 'bg': WIN_BG}
        
        # Create search bar with icon-like prefix
        search_container = ttk.Frame(search_frame)
        search_container.pack(side='top', fill='x', pady=2)
        
        search_icon = tk.Label(search_container, text="🔍", font=('Segoe UI', 12), bg=WIN_BG)
        search_icon.pack(side='left', padx=(0,2))
        
        search_label = tk.Label(search_container, text='Quick Search:', **filter_label_style)
        search_label.pack(side='left', padx=(0,6))
        
        # Create search as a combobox with autocomplete
        self.search_entry = ttk.Combobox(search_container, width=20)
        self.search_entry.pack(side='left', padx=(0,6))
        self.search_entry.insert(0, "Type to search...")
        self.search_entry.config(foreground='gray')
        # Set up focus events for placeholder behavior for combobox
        def on_combobox_click(event):
            if self.search_entry.get() == "Type to search...":
                self.search_entry.delete(0, "end")
                self.search_entry.config(foreground='black')
            # When the combobox is clicked, update the suggestions
            self.update_search_suggestions()
                
        def on_combobox_leave(event):
            if self.search_entry.get() == "":
                self.search_entry.set("")  # Clear first
                self.search_entry.insert(0, "Type to search...")
                self.search_entry.config(foreground='gray')
                
        self.search_entry.bind("<FocusIn>", on_combobox_click)
        self.search_entry.bind("<FocusOut>", on_combobox_leave)
        
        # Bind selection event to trigger search immediately
        self.search_entry.bind("<<ComboboxSelected>>", lambda event: self.refresh_table())
        
        # Bind key release event to trigger search with small delay
        self.search_entry.bind("<KeyRelease>", lambda event: self.after(300, self.delayed_search))
        
        # Initialize the suggestions list
        self.update_search_suggestions()
        
        # Create search options in a separate row for clarity - radio buttons first
        options_frame = ttk.Frame(search_frame)
        options_frame.pack(side='top', fill='x', pady=2)
        
        # Add radio buttons for search type
        self.search_type_var = tk.StringVar(value="Name")
        ttk.Radiobutton(options_frame, text="Name", variable=self.search_type_var, 
                       value="Name", command=self.update_search_selections).pack(side='left', padx=(5,5))
        ttk.Radiobutton(options_frame, text="Business Unit", variable=self.search_type_var, 
                       value="Business Unit", command=self.update_search_selections).pack(side='left', padx=5)
        
        # Move search container below the radio buttons
        search_container.pack_forget()  # Remove the current packing of the search container
        search_container.pack(side='top', anchor='w', pady=(5, 0), padx=(20, 0))  # Position below and indented
        
        # Add search button and clear button in a nicer layout
        button_frame = ttk.Frame(search_container)
        button_frame.pack(side='right', padx=(6,0))
        
        ttk.Button(search_container, text='Clear', command=self.clear_search, 
                 style='Secondary.TButton').pack(side='right', padx=(6,0))

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
            'Monetary': 80, 'CustomerService': 140, 'Business Unit': 180, 'Risk': 80, 'Last Modified': 150
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

    def update_search_selections(self):
        """Handle search type change and update dropdown values"""
        self.update_search_suggestions()
        self.refresh_table()
        
    def update_search_suggestions(self, event=None):
        """Update the dropdown suggestions based on the search type selected"""
        if not hasattr(self, 'search_entry') or self.search_entry is None:
            return
            
        search_type = "Name"
        if hasattr(self, 'search_type_var') and self.search_type_var is not None:
            search_type = self.search_type_var.get()
            
        # Get appropriate values from the database
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        
        if search_type == "Name":
            c.execute("SELECT DISTINCT name FROM applications ORDER BY name")
            values = [row[0] for row in c.fetchall()]
        else:  # Business Unit
            c.execute("SELECT DISTINCT name FROM business_units ORDER BY name")
            values = [row[0] for row in c.fetchall()]
            
        conn.close()
        
        # Update the combobox values
        self.search_entry['values'] = values
    
    def refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        conn = database.sqlite3.connect(database.DB_NAME)
        c = conn.cursor()
        
        # Get search parameters if available
        search_text = ""
        search_type = "Name"
        if hasattr(self, 'search_entry') and self.search_entry is not None:
            entered_text = self.search_entry.get().strip().lower()
            # Don't search if the placeholder text is showing
            if entered_text != "type to search...":
                search_text = entered_text
        if hasattr(self, 'search_type_var') and self.search_type_var is not None:
            search_type = self.search_type_var.get()
            
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
            
            # Apply search filtering
            if search_text:
                if search_type == "Name" and search_text.lower() not in app_row[1].lower():
                    continue
                elif search_type == "Business Unit" and not any(search_text.lower() in dept.lower() for dept in depts):
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
