# project imports/dependencies/libraries
import pyodbc
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from win32com.client import Dispatch
import pandas as pd

# create a class for mergining reporting databases for ease of use
class ReportingMerge():
    # Constructor function meaning this runs on startup. __init__ is called when the class is instantiated and is built into python
    def __init__(self):
        
        # define table names in a list for easy reference
        self.all_tables = [
            "ACTIVITY_CODES",
            "ACTUALWELL",
            "ADDRESSES",
            "BID_SHEETS",
            "COST_CODES",
            "CRS",
            "DAILY_ACTIVITY",
            "DAILY_REPORTS",
            "DATABASE_INFO",
            "DEPTH_DATUMS",
            "DRILL_STRING",
            "ENG_FLUIDS",
            "ENGINEERING",
            "EXPORT_CHECK_TABLE",
            "EXTERNAL_FILES",
            "FACILITY",
            "FIELD",
            "FIELD_COST",
            "FLOW_RATES",
            "FRICTION_FACTOR",
            "GEOSTEERING",
            "HYD_PARAMS",
            "INVENTORY",
            "IPM",
            "MOTOR_REPORT",
            "OFFSET_WELL_SELECTOR",
            "OPERATOR",
            "OUTER_STRING",
            "PHASE_CODES",
            "PIPE_TALLY",
            "PLANNEDWELL",
            "PLANPROJECTS",
            "PLANPROJECTS_COSTS",
            "PLANPROJECTS_DEPTHTIME",
            "PLANPROJECTS_PHASES",
            "PLANPROJECTS_RISKS",
            "PUMP_DATA",
            "SHIPPING_TICKET",
            "SLIDE_RECORDS",
            "SURVEY",
            "TARGETS",
            "TEMPLATE",
            "TND_PARAMS",
            "TRIP_SPEEDS",
            "WELL",
            "WELLBORE",
            "WITSML"
        ]
        
        # define table names that will be merged in a list for easy reference
        self.merge_tables = ['DAILY_REPORTS', 'DRILL_STRING', 'DAILY_ACTIVITY', 'SLIDE_RECORDS', 'MOTOR_REPORT', 'FIELD_COST']
    
        # Initialize the root window for tkinter
        self.root = tk.Tk()
        self.root.withdraw()
        
        # tuple of database paths
        db_paths = filedialog.askopenfilename(title="Select Databases to Merge", filetypes=[("WellSeeker DB", "*.mdb")], multiple=True)
        
        # create the database connections
        self._create_database_connections(db_paths=db_paths)
        
        # create the result database connection
        self._create_result_database()
        
    def _check_export_table(self, db_connections: list):
        # Check that the export tables exist and that each has a record_id column with value 1
        for conn in db_connections:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM EXPORT_CHECK_TABLE")
            count = cursor.fetchone()
            if count[0] == 0:
                messagebox.showerror("Error", "The selected database is not a wellseeker export table.")
                exit()
            else:
                cursor.execute("SELECT RECORD_ID FROM EXPORT_CHECK_TABLE")
                record_id = cursor.fetchone()
                if record_id[0] != 1:
                    messagebox.showerror("Error", "The selected database is not a REPORTING table export.")
                    exit()
              
    def _determine_db_order(self, db_connections: list) -> list:
        # This function sorts the database connections based on the earliest date in the DAILY_REPORTS table
        
        # Function to fetch the earliest date from the database
        def get_earliest_date(conn):
            cursor = conn.cursor()
            cursor.execute("SELECT MIN(YEAR), MIN(MONTH), MIN(DAY) FROM DAILY_REPORTS WHERE YEAR=(SELECT MIN(YEAR) FROM DAILY_REPORTS)")
            year, month, day = cursor.fetchone()
            return year, month, day
        
        # Get earliest date for each connection and sort accordingly
        sorted_db_connections = sorted(db_connections, key=lambda conn: get_earliest_date(conn))
        
        # verify databases are reporting exports
        self._check_export_table(db_connections=db_connections)
        
        return sorted_db_connections

    def _create_database_connections(self, db_paths: tuple = None):
        if db_paths is not None:
            if len(db_paths) > 1:
                # List to store the database connections
                db_connections = []
                
                for db in db_paths:
                    # Create a new database connection
                    connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db
                    conn = pyodbc.connect(connection_string)
                    db_connections.append(conn)
                
                # Sort the connections from oldest to newest using the determine_db_order function
                self.db_connections = self._determine_db_order(db_connections)
                print(self.db_connections)
            else:
                # Not Enough Databases
                retry = messagebox.askretrycancel("Not Enough Databases", "Not Enough Databases. Please select at least two databases")
                
                # If the user chooses to retry
                if retry:
                    db_paths = filedialog.askopenfilename(title="Select Databases to Merge", filetypes=[("WellSeeker DB", "*.mdb")], multiple=True)
                    # Call the function recursively with the new paths
                    self._create_database_connections(db_paths=db_paths)
                else:
                    exit()
                
    def _create_result_database(self):
        # Ask the user where to save the merged database
        result_db_path = filedialog.asksaveasfilename(title="Save Merged Database As", filetypes=[("WellSeeker DB", "*.mdb")], defaultextension=".mdb")
        self.result_db_name = os.path.basename(result_db_path)
        
        # check if path already exists
        if os.path.exists(result_db_path):
            messagebox.showerror("Error", "The path provided already exists.")
            exit()
        else:
            try:
                # create the new access database
                accApp = Dispatch("Access.Application")
                dbEngine = accApp.DBEngine
                workspace = dbEngine.Workspaces(0)
                dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
                newdb = workspace.CreateDatabase(result_db_path, dbLangGeneral, 64)
            except Exception as e:
                messagebox.showerror("Error", str(e))
                exit()

            finally:
                accApp.DoCmd.CloseDatabase
                accApp.Quit
                newdb = None
                workspace = None
                dbEngine = None
                accApp = None
            
        # If a path was provided (i.e., the user didn't cancel the save dialog)
        if result_db_path:
            # Create a new Access database using pyodbc (empty for now)
            conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + result_db_path

            # Store a connection to the newly created database in self.result_db
            self.result_db = pyodbc.connect(conn_str)
        else:
            messagebox.showerror("Error", "No path provided for the result database.")
            exit()
            
    def select_actualwell_name(self):
        # Create an empty set to store unique actual well names
        actualwell_set = set()

        # Loop through each database connection and fetch unique actualwell names
        for db_conn in self.db_connections:
            cursor = db_conn.cursor()
            cursor.execute("SELECT DISTINCT ACTUALWELL FROM DAILY_REPORTS")
            actualwells = [row.ACTUALWELL for row in cursor.fetchall()]
            cursor.close()
            
            # Add the actualwell names to the set
            actualwell_set.update(actualwells)

        # Convert the set to a list
        actualwell_list = list(actualwell_set)

        # If no actual wells found
        if not actualwell_list:
            messagebox.showinfo("Info", "No actual wells found.")
            return

        # Create a new top-level window
        select_window = tk.Toplevel(self.root)
        select_window.title("Select an Actual Well")

        # Label
        label = tk.Label(select_window, text="Select an Actual Well:")
        label.pack(pady=20)
        
        def scroll_x(*args):
            listbox.xview(*args)

        # Create a listbox to list the actual wells
        listbox = tk.Listbox(select_window, height=10, width=40, font=("Arial", 12))
        for actualwell in actualwell_list:
            listbox.insert(tk.END, actualwell)
        listbox.pack(pady=20)
        
        # Create a horizontal scrollbar
        scrollbar = tk.Scrollbar(select_window, orient=tk.HORIZONTAL)
        scrollbar.pack(fill=tk.X)

        # Configure listbox and scrollbar to work together
        listbox.config(xscrollcommand=scrollbar.set)
        scrollbar.config(command=scroll_x)
                
        # Button to confirm selection
        confirm_button = tk.Button(select_window, text="Confirm", command=lambda: self._set_selected_actualwell(listbox, select_window))
        confirm_button.pack(pady=20)
        
        # Start the Tkinter main loop
        self.root.mainloop()
        
    def _set_selected_actualwell(self, listbox, window):
        selected_actualwell = listbox.get(listbox.curselection())
        self.selected_actualwell = selected_actualwell        
        window.destroy()
        self.root.destroy()
            
    def _generate_create_table_sql(self, conn, table_name):
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {table_name} WHERE 1=0")  # This will fetch the table structure without any data
        columns = []

        for desc in cursor.description:
            field_name = desc[0]
            field_type = desc[1]
            
            # Map pyodbc type to Access data type
            if field_type == str:
                type_str = 'TEXT'
            elif field_type in [int, float]:
                type_str = 'NUMBER'
            elif field_type == bytes:
                type_str = 'BINARY'
            else:
                type_str = 'TEXT'  # Default, but you may need to adjust based on your actual data types
            
            columns.append(f"{field_name} {type_str}")

        columns_str = ', '.join(columns)
        create_table_sql = f"CREATE TABLE {table_name} ({columns_str})"
        return create_table_sql

    def table_exists(self, table_name):
        cursor = self.result_db.cursor()
        try:
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            return True
        except:
            return False
        
    def _move_all_table_data(self):
        # filter through all non merge tables to get the data
        for table in self.all_tables:
            if table not in self.merge_tables:
                # Collect data from all source databases into a list of DataFrames
                dfs = []

                for db_conn in self.db_connections:
                    query = f"SELECT * FROM {table}"
                    df = pd.read_sql(query, db_conn)
                    dfs.append(df)

                # Concatenate all data and remove duplicates
                combined_df = pd.concat(dfs, ignore_index=True).drop_duplicates()

                # If the table doesn't exist in result_db, create it
                cursor = self.result_db.cursor()
                if not self.table_exists(table):
                    create_table_sql = self._generate_create_table_sql(self.db_connections[0], table)
                    cursor.execute(create_table_sql)

                # Insert deduplicated data into the result_db
                for index, row in combined_df.iterrows():
                    placeholders = ', '.join(['?'] * len(row))
                    sql = f"INSERT INTO {table} VALUES ({placeholders})"
                    cursor.execute(sql, tuple(row))

                self.result_db.commit()
                
    def _move_merge_table_data(self):
        # move the data that needs to be merged and numbered
        daily_report_record_id = 0
        daily_report_rep_num = 0
        daily_report_uid = 0
        
        daily_activity_record_id = 0
        daily_activity_bha = 0
        daily_activity_uid = 0
        
        drill_string_bha_num = 0
        drill_string_bha_num_rep = 0
        drill_string_mwd_run_num = 0
        drill_string_uid = 0
        
        field_cost_record_id = 0
        field_cost_bha_num = 0
        field_cost_uid = 0
        
        motor_report_record_id = 0
        motor_report_uid = 0
        
        slide_records_bha_num = 0
        slide_records_bha_link = 0
        slide_records_uid = 0
        
        for db_conn in self.db_connections:
            # Collect data from all source databases into a list of DataFrames
            for table in self.merge_tables:
                # get table data    
                query = f"SELECT * FROM {table}"
                df = pd.read_sql(query, db_conn)
                
                # check if actualwell exists in df
                if 'ACTUALWELL' in df.columns:
                    # update actualwell
                    df['ACTUALWELL'] = self.selected_actualwell

                #######################################################################
                # handle the primary tables drill_string and daily_reports
                
                # handle Daily Reports data
                if table == "DAILY_REPORTS":
                    # merge data calculations
                    df['RECORD_ID'] = df['RECORD_ID'].astype(int) + daily_report_record_id
                    df['REP_NUM'] = df['REP_NUM'].astype(int) + daily_report_rep_num
                    df['UID'] = df['UID'].astype(int) + daily_report_uid
                    
                    # update merge numbers
                    daily_report_record_id = df['RECORD_ID'].max()
                    daily_report_rep_num = df['REP_NUM'].max()
                    daily_report_uid = df['UID'].max()
                    
                # handle Drill String data
                elif table == "DRILL_STRING":
                    # merge data calculations
                    df['BHA_NUM'] = df['BHA_NUM'].astype(int) + drill_string_bha_num
                    df['BHA_NUM_REP'] = df['BHA_NUM_REP'].astype(int) + drill_string_bha_num_rep
                    
                    # Handle empty strings in MWD_RUN_NUM
                    df['MWD_RUN_NUM'] = df['MWD_RUN_NUM'].replace('', '0')
                    df['MWD_RUN_NUM'] = df['MWD_RUN_NUM'].astype(int) + drill_string_mwd_run_num
                    
                    df['UID'] = df['UID'] + drill_string_uid

                    # update merge numbers
                    drill_string_bha_num = df['BHA_NUM'].max()
                    drill_string_bha_num_rep = df['BHA_NUM_REP'].max()
                    drill_string_mwd_run_num = df['MWD_RUN_NUM'].max()
                    drill_string_uid = df['UID'].max()
                    
                ##########################################################################
                # following tables are dependent on the following values from above tables
                # daily_report_record_id
                # drill_string_bha_num
                    
                # handle Daily Activities data
                elif table == "DAILY_ACTIVITY":
                    # merge data calculations
                    df['RECORD_ID'] = df['RECORD_ID'].astype(int) + daily_activity_record_id
                    df['BHA'] = df['BHA'].astype(int) + daily_activity_bha
                    df['UID'] = df['UID'].astype(int) + daily_activity_uid
                    
                    # update merge numbers
                    daily_activity_record_id = daily_report_record_id
                    daily_activity_bha = drill_string_bha_num
                    daily_activity_uid = df['UID'].max()
                    
                # handle Slide Records data
                elif table == "SLIDE_RECORDS":
                    # merge data calculations
                    df['BHA_NUM'] = df['BHA_NUM'].astype(int) + slide_records_bha_num
                    
                    # mask so that only non-null values are merged
                    mask = (df['BHA_LINK'].notna()) & (df['BHA_LINK'] != '')
                    df.loc[mask, 'BHA_LINK'] = df.loc[mask, 'BHA_LINK'].astype(int) + slide_records_bha_link
                    
                    df['UID'] = df['UID'].astype(int) + slide_records_uid
                    
                    # update merge numbers
                    slide_records_bha_num = drill_string_bha_num
                    slide_records_bha_link = drill_string_bha_num
                    slide_records_uid = df['UID'].max()
                    
                    # Replace NaN values with 0 for the entire DataFrame
                    df.fillna(0, inplace=True)
                    
                # handle Field Costs data
                elif table == "FIELD_COST":
                    # merge data calculations
                    df['RECORD_ID'] = df['RECORD_ID'].astype(int) + field_cost_record_id
                    
                    # update bha num where not nan or -1
                    condition = (df['BHA_NUM'].notna()) & (df['BHA_NUM'] != -1) & (df['BHA_NUM'] != '') & (df['BHA_NUM'] != '-1')
                    df.loc[condition, 'BHA_NUM'] = df.loc[condition, 'BHA_NUM'].astype(int) + field_cost_bha_num
                    
                    df['UID'] = df['UID'].astype(int) + field_cost_uid
                    
                    # update merge numbers
                    field_cost_record_id = daily_report_record_id
                    field_cost_bha_num = drill_string_bha_num
                    field_cost_uid = df['UID'].max()
                
                # handle Motor Reports data
                elif table == "MOTOR_REPORT":
                    #merge data calculations
                    df['RECORD_ID'] = df['RECORD_ID'].astype(int) + motor_report_record_id
                    df['UID'] = df['UID'].astype(int) + motor_report_uid
                    
                    # update merge numbers
                    motor_report_record_id = drill_string_bha_num
                    motor_report_uid = df['UID'].max()
                
                # write data to the result_db and combine it with the existing data, also create table if it does not exist.
                # If the table doesn't exist in result_db, create it
                cursor = self.result_db.cursor()
                if not self.table_exists(table):
                    create_table_sql = self._generate_create_table_sql(self.db_connections[0], table)
                    cursor.execute(create_table_sql)
                
                # Insert merged data into the result_db
                for index, row in df.iterrows():
                    placeholders = ', '.join(['?'] * len(row))
                    sql = f"INSERT INTO {table} VALUES ({placeholders})"
                    cursor.execute(sql, tuple(row))
                    
                self.result_db.commit() 
            
    def merge_dbs(self):
        self._move_all_table_data()  
        self._move_merge_table_data()        
            
    def _close_database_connections(self):
        for db in self.db_connections:
            db.close()
        try:
            self.result_db.close()
        except:
            pass
        
    def close(self):
        try:
            self._close_database_connections()
        except:
            pass
        try:
            self.root.destroy()
        except:
            # force kill command
            os.system("taskkill /im WSReportingMerge.exe")
        
# MAIN CALL FUNCTION
if __name__ == '__main__':
    Merge = None
    try:
        # INITIALIZE THE CLASS
        Merge = ReportingMerge()
        
        Merge.select_actualwell_name()
        
        # MERGE THE DATABASES
        Merge.merge_dbs()
    finally:
        if Merge is not None:
            # CLOSE OUT THE CLASS
            Merge.close()  
            messagebox.showinfo("Done", "Successfully create {result_db_name} database".format(result_db_name=Merge.result_db_name))