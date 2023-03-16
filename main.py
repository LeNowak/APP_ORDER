"""
Data structure for SQLite db
order_tables: ID, name, description, date_created,comments, is_active
articles: ID, name,	finish,	code_ID, order_table_ID, drawing_no, revision,  units,	description	,comments,	qty_factor,	qty_on_assembly, made_of,weight, created,	internal_comments,is_active
profiles: ID, name,	initial_name, finish,	code_ID, order_table_ID, bar_length, description	,comments,	width,	end_cut,painting_area, weight,	created, internal_comments,is_active
codes: ID, name, description, date_created, is_active
finishes: ID, name, description, date_created, is_active
projects: ID, name, description, date_created, is_active
orders: ID, project_ID, name, description,job, date_created,verify_date, status,is_active
records: ID, orderID, article_ID,profile_ID, description, date_created, is_sent,is_active
"""
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as msb
from tkinter import filedialog as fd
import db_order_manager as db
import pandas as pd
import builtins
import os
from PIL import Image, ImageTk
from datetime import datetime
import openpyxl
import datetime
###########################################
import config 

###########################################


if not os.path.exists('C:/Users/HP/OneDrive/APP_ORDERS'):
    builtins.main_folder = 'C:/OneDrive/APP_ORDERS/'
else:
    builtins.main_folder = 'C:/Users/HP/OneDrive/APP_ORDERS/'
builtins.images_folder = builtins.main_folder + 'images/'

class PopupWindow(tk.Toplevel):
    def __init__(self, master, options):
        tk.Toplevel.__init__(self, master)
        self.combobox = ttk.Combobox(self, values=options)
        self.combobox.current(0)
        self.combobox.pack()
        ttk.Button(self, text="OK", command=self.ok).pack()

    def ok(self):
        self.result = self.combobox.get()
        self.destroy()


class order_app(tk.Tk):
    def __init__(self):
        builtins.current_project_ID = config.readCFGval("DEFAULT", "currentproject")
        #check if current_project_ID is in the database in project table and get its name
        builtins.project_name = db.check_if_id_exists(current_project_ID,'project')
        if len(project_name)==0:
            msb.showinfo("Project ID", "Project ID: " + current_project_ID + " is not in the database")
            exit()


        self.skip_list = ['ID','date_created','verify_date','is_active']


        tk.Tk.__init__(self)
        app = tk.Tk()
        app.withdraw()
        self.geometry("900x600")
        self.title('MAS-TASK ORDER MANAGER / ' + project_name)

        self.main_menu()
        self.orders_list()
        self.mainloop()

    def main_menu(self):
        try:
            for widget in self.left_frame.winfo_children():
                widget.destroy()
            self.left_frame.pack_forget()
            self.left_frame.destroy()
        except: pass
        self.orders_list()
        # Create left frame for buttons
        self.left_frame = tk.Frame(self, width=200, height=600)
        self.left_frame.grid(row=0, column=0, sticky="nsew",padx=10,pady=10)

        # Create buttons and add to left frame
        button0 = tk.Button(self.left_frame, text="ORDERS", command=self.menu_order,width=30)
        button0.pack()
        button1 = tk.Button(self.left_frame, text="ARTICLES", command=self.empty,width=30)
        button1.pack()
        button2 = tk.Button(self.left_frame, text="PROJECTS", command=self.menu_project,width=30)
        button2.pack()
        button2 = tk.Button(self.left_frame, text="CONFIGURATION", command=self.menu_config,width=30)
        button2.pack()   

        #show image
        image = Image.open(main_folder + "logo.png")
        image = image.resize((200,150), Image.ANTIALIAS)
        photo = ImageTk.PhotoImage(image)
        label = tk.Label(self.left_frame, image=photo)
        label.image = photo
        label.pack()

    def menu_order(self):
        try:
            for widget in self.left_frame.winfo_children():
                widget.destroy()
            self.left_frame.pack_forget()
            self.left_frame.destroy()
        except: pass
        self.orders_list()
  
        # Create left frame for buttons
        self.left_frame = tk.Frame(self, width=200, height=600)
        self.left_frame.grid(row=0, column=0, sticky="nsew",padx=10,pady=10)

        # Create buttons and add to left frame
        button0 = tk.Button(self.left_frame, text="<<<      BACK", command=self.main_menu,width=30)
        button0.pack()
        button1 = tk.Button(self.left_frame, text="SELECT ORDER TO UPDATE", command=self.empty,width=30)
        button1.pack()
        button2 = tk.Button(self.left_frame, text="NEW ORDER", command=lambda:self.add_new('orders'),width=30)
        button2.pack()
        button3 = tk.Button(self.left_frame, text="NEW ORDER FROM EXCEL", command=self.add_order_from_excel,width=30)
        button3.pack()
        button4 = tk.Button(self.left_frame, text="ORDER CLOSE / SEND", command=self.empty,width=30)
        button4.pack()
        button5 = tk.Button(self.left_frame, text="TO EXCEL", command=self.empty,width=30)
        button5.pack()

    def menu_project(self):
        try:
            for widget in self.left_frame.winfo_children():
                widget.destroy()
            self.left_frame.pack_forget()
            self.left_frame.destroy()
        except: pass

        self.projects_list()
        # Create left frame for buttons
        self.left_frame = tk.Frame(self, width=400, height=800)
        self.left_frame.grid(row=0, column=0, sticky="nsew",padx=10,pady=10)

        # Create buttons and add to left frame
        button0 = tk.Button(self.left_frame, text="<<<      BACK", command=self.main_menu,width=30)
        button0.pack()
        button1 = tk.Button(self.left_frame, text="CHANGE ACTIVE PROJECT", command=self.change_active_project,width=30)
        button1.pack()

        button1 = tk.Button(self.left_frame, text="ADD PROJECT", command=lambda:self.add_new('project'),width=30)
        button1.pack()
        button2 = tk.Button(self.left_frame, text="UPDATE PROJECT", command=self.empty,width=30)
        button2.pack()
        button1 = tk.Button(self.left_frame, text="ADD JOB", command=lambda:self.add_new('job'),width=30)
        button1.pack()
        button2 = tk.Button(self.left_frame, text="UPDATE JOB", command=self.empty,width=30)
        button2.pack()


        button3 = tk.Button(self.left_frame, text="ORDERS FOR PROJECT", command=self.empty,width=30)
        button3.pack()
        b5 = tk.Button(self.left_frame, text="LIST PROJECTS", command=self.show_projects_list,width=30)
        b5.pack()
        b5 = tk.Button(self.left_frame, text="LIST JOBS FOR PROJECT", command=self.show_projects_list,width=30)
        b5.pack()
        b6 = tk.Button(self.left_frame, text="ARCHIVE PROJECT", command=self.empty,width=30)
        b6.pack()

    def menu_config(self):
        try:
            for widget in self.left_frame.winfo_children():
                widget.destroy()
            self.left_frame.pack_forget()
            self.left_frame.destroy()
        except: pass

        self.clear_right_framet()
        # Create left frame for buttons
        self.left_frame = tk.Frame(self, width=400, height=800)
        self.left_frame.grid(row=0, column=0, sticky="nsew",padx=10,pady=10)

        # Create buttons and add to left frame
        button0 = tk.Button(self.left_frame, text="<<<      BACK", command=self.main_menu,width=30)
        button0.pack()
        button1 = tk.Button(self.left_frame, text="ADD FROM EXCEL", command=self.add_from_excel,width=30)
        button1.pack()
        button2 = tk.Button(self.left_frame, text="ADD ARTICLE CODE", command=lambda:self.add_new('code'),width=30)
        button2.pack()
        button3 = tk.Button(self.left_frame, text="ADD FINISH", command=lambda:self.add_new('finish'),width=30)
        button3.pack()
        button4 = tk.Button(self.left_frame, text="ADD ORDER_TABLE", command=lambda:self.add_new('order_table'),width=30)
        button4.pack()
        button5 = tk.Button(self.left_frame, text="ADD UNIT", command=lambda:self.add_new('unit'),width=30)
        button5.pack()
        button6 = tk.Button(self.left_frame, text="ADD STATUS", command=lambda:self.add_new('status'),width=30)
        button6.pack()

    def change_active_project(self):
        #show popup window with list of projects
        projects_dict = db.get_active('project')
        project_name,project_ID = self.show_list(projects_dict)
        #change active project
        config.updateCFGval("DEFAULT","currentproject",project_ID)
        #update active project in main window
        self.active_project.set(project_name)


    #function creates popup window with combobox based on given list
    #function returns selected item to main window
    def show_list(master, options_dict):
        popup = PopupWindow(master, list(options_dict.keys()))
        master.wait_window(popup)
        return popup.result, options_dict[popup.result]

    def empty(self):
        print("EMPTY")

    def widget_to_dictionary(self,table,widget_dict):
        #function add data from widget_dict to database
        #create list of values from widget_dict
        #print(widget_dict)
        values = {}
        #check if all entries are proper type according to db table
        for column,typ in db.get_table_columns(table).items():
            if column in self.skip_list:
                continue
            if column.endswith('_ID'):
                continue
            #check if datatypes are correct
            if typ == 'INTEGER':
                try: 
                    int(widget_dict[column].get())
                except:
                    msb.showerror("Error","Not integer value in "+column)
                    self.window.focus_set()
                    return
            elif typ == 'REAL':
                try: 
                    float(widget_dict[column].get())
                except:
                    msb.showerror("Error","Not float value in "+column)
                    self.window.focus_set()
                    return
            elif typ == 'TEXT':
                if len(widget_dict[column].get()) > 100:
                    msb.showerror("Error","Maximum text length in "+column)
                    self.window.focus_set()
                    return

            #print(widget_dict[column].get(),typ)
        for column,value in widget_dict.items():
            if column.endswith('_ID'):
                dict = db.get_ID_from_column_as_dict('name',column[:-3])
                #print('###', column, dict, value.get())
                values[column] = dict[value.get()]
            else:
                values[column] = value.get()
        #print(values)

        #add values to database
        result = db.add_data(table,values)
        #destroy window
        self.orders_list()
        self.window.destroy()
        return result

    def add_new(self,table,suggested_name=''):

        #function add new item to database
        #create new window
        self.window = tk.Toplevel()
        self.window.title("Add new item to "+table+" table")
        #self.window.geometry("300x300")
        self.window.resizable(False, False)
        #Add desctiptio from table comment
        label = tk.Label(self.window, text=db.get_comment_form_table(table), width=30)   
        label.pack(padx=20,pady=20)

        #create entry for each db table column. If column is on skip_list skip it
        #if column ends with '_ID' create combobox with values from other table with same name without '_ID'

        widget_dict = {}
        values = {}
        for column,typ in db.get_table_columns(table).items():
            if column in self.skip_list:
                continue
            if column.endswith('_ID'):
                label = tk.Label(self.window, text=column, width=20)
                label.pack(padx=20)
                values = db.get_active_from_column('name' , column[:-3])

                widget_dict[column] = ttk.Combobox(self.window, values=values, width=30)
                widget_dict[column].pack(padx=20)
            else:
                label = tk.Label(self.window, text=column)
                label.pack()
                widget_dict[column] = tk.Entry(self.window, width=30)
                widget_dict[column].pack()


        #create button
        button = tk.Button(self.window, text="Add", command=lambda:self.widget_to_dictionary(table,widget_dict))
        button.pack()
        button_cancel = tk.Button(self.window, text="Cancel", command=self.window.destroy)
        button_cancel.pack()

    ############################################################ ORDERS ########################################
    def add_order_from_excel(self):


        init_dir = main_folder + '/excel_data/'
        file = fd.askopenfilename(initialdir = init_dir,title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
        #check if file is selected
        if file == '':
            return
        #check if file update is max 5 days old
        if (datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getmtime(file))).days > 5:
            ask = msb.askyesno("Warning","File is older than 5 days. Do you want to continue?")
            if ask == False:
                return

        wb = openpyxl.load_workbook(file, data_only=True)
        sheet = wb.active
        #value from excel , cell A1
        A1 = sheet['A1'].value
        order_data = {}
        order_data['order_ID'] = sheet['B1'].value
        order_data['project'] = sheet['B2'].value

        order_data['name'] = sheet['B5'].value
        order_data['job'] = sheet['E1'].value
        order_data['issue_date'] = sheet['E2'].value
        order_data['order_type'] = sheet['E3'].value


        #check if order already exist
        if db.check_id_exists(order_data['order_ID'],'orders'):
            ask = msb.askyesnocancel("Warning","Order already exist. Do you want to continue and overwrite it (Yes), Create new (No)?")
            if ask == None:
                return
            elif ask == False:
                order_data['order_ID'] = db.get_new_id('orders')
                #print(order_data['order_ID'].value)
                sheet['B1'] = order_data['order_ID']
                wb.save(file)
            else:
                pass

        #check if project exist
        projects_dict = db.get_active('project')

        order_types_dict = db.get_active('order_type')   

        #print(str(order_data['project']))
        project_list = [str(key) for key in list(projects_dict.keys())]
        #print(project_list)
        if str(order_data['project']) not in project_list:
            ask = msb.askyesnocancel("Warning",f"Project {order_data['project']} not exist. Do you want to add new (Yes)? or select from list? (No)")
            if ask == None:
                return
            elif ask == False:
                project_name,project_ID = self.show_list(projects_dict)
            else:
                project_ID = db.add_new('project',order_data['project'].value)
                job_ID = db.add_new('job',order_data['job'].value)
        else:
            project_name = str(order_data['project'])
            #print(project_name)
            project_ID = projects_dict[project_name]
            #print(project_ID)
        #select job based on db job list where project_ID = project_ID
        widget_dict = {}

        jobs_df = db.select_all_where('job',"project_ID = " + str(project_ID))
        #print(jobs_df)
        jobs_dict = jobs_df.set_index('ID')['name'].to_dict()
        #print(jobs_dict)

        ###########################################################################################

        self.window = tk.Toplevel()
        self.window.title("New order from excel")
        self.project_ID = project_ID
        self.project_name = project_name
        #self.window.geometry("300x300")
        #self.window.resizable(False, False)

        self.frame_info = tk.Frame(self.window)
        self.frame_info.grid(row=10,column=0,padx=5,pady=10)

        label = tk.Label(self.frame_info, text="Read file : ", width=10)   
        label.grid(row=1,column=0,padx=5,pady=10)
        file_no_path = os.path.basename(file)
        label = tk.Label(self.frame_info, text=file_no_path, width=30)
        label.grid(row=1,column=1,padx=5,pady=10)

        #create entry for each db table column. If column is on skip_list skip it
        label = tk.Label(self.frame_info, text="Project: " , width=10)
        label.grid(row=3,column=0,padx=5,pady=5)
        label = tk.Label(self.frame_info, text=project_name, width=20)   
        label.grid(row=3,column=1,padx=5,pady=5)


        label = tk.Label(self.frame_info, text='JOB', width=10)   
        label.grid(row=5,column=0,padx=5,pady=5)
        job = tk.StringVar()
        if order_data['job'] in jobs_dict.values():
            print(order_data['job'])
            job.set(order_data['job'])
        else:
            job.set('')
        widget_dict['job'] = ttk.Combobox(self.frame_info, values=list(jobs_dict.values()), width=30, textvariable=job)
        widget_dict['job'].grid(row=5,column=1,padx=5,pady=5)

        label = tk.Label(self.frame_info, text="Order Type", width=10)   
        label.grid(row=7,column=0,padx=5,pady=5)
        order_type = tk.StringVar()
        if order_data['order_type'] in order_types_dict.keys():
            #print(order_data['order_type'])
            order_type.set(order_data['order_type'])
        else:
            order_type.set('')
        widget_dict['order_type'] = ttk.Combobox(self.frame_info, values=list(order_types_dict.keys()), width=30, textvariable=order_type)
        widget_dict['order_type'].grid(row=7,column=1,padx=5,pady=5)


        label = tk.Label(self.frame_info, text="Order name", width=10)
        label.grid(row=9,column=0,padx=5,pady=5)
        field = tk.StringVar()
        field.set(order_data['name'])
        widget_dict['name'] = tk.Entry(self.frame_info, textvariable=field, width=40)
        widget_dict['name'].grid(row=9,column=1,padx=5,pady=5)


        label = tk.Label(self.frame_info, text="Order description", width=25)
        label.grid(row=10,column=0,padx=5,pady=5)
        field = tk.StringVar()
        field.set('')
        widget_dict['description'] = tk.Entry(self.frame_info, textvariable=field, width=40)
        widget_dict['description'].grid(row=10,column=1,padx=5,pady=5)


        label = tk.Label(self.frame_info, text="Issue date: " , width=10)
        label.grid(row=11,column=0,padx=5,pady=5)
        issue_date = tk.StringVar()
        issue_date.set(str(order_data['issue_date'])[:10])
        widget_dict['issue_date'] = tk.Entry(self.frame_info, textvariable=issue_date, width=20)
        widget_dict['issue_date'].grid(row=11,column=1,padx=5,pady=5)



        button = tk.Button(self.frame_info, text="   Add  order  ", command=lambda:self.add_order_from_excel_2(widget_dict))
        button.grid(row=20,column=0,padx=10,pady=20)
        
        button_c = tk.Button(self.frame_info, text="   Cancel   ", command=self.window.destroy)
        button_c.grid(row=20,column=1,padx=20,pady=20)

        ##################################################################################################################
        #read list of articles from db table articles
        articles_list = db.get_active_as_list('article')
        articles_list_in_project = db.get_articles_for_project(project_ID)

        #read excel file
        df = pd.read_excel(file,skiprows=6)
        df['in_db'] = ''
        df.loc[~df.name.isin(articles_list),'in_db'] = 'X'
        df['in_project'] = ''
        df.loc[~df.name.isin(articles_list_in_project),'in_project'] = 'X'
        self.df = df
        label = tk.Label(self.window, text="Articles in DB: " )
        label.grid(row=41,column=0,padx=5,pady=5)
        self.list_to_treeview(42)
        self.window.mainloop()

    def add_order_from_excel_2(self,widgets):
        order_data = {k:v.get() for k,v in widgets.items()}
        order_data['project'] = self.project_name
        print(order_data)

        order_ID = db.add_order_to_db(order_data)
        print(order_ID)
        pass

    def list_to_treeview(self, row_in_frame):    
        try:
            for widget in self.frame_TVb.winfo_children():
                widget.destroy()
            self.frame_TVb.pack_forget()
            self.frame_TVb.destroy()
        except: pass
        self.orders_list()

        self.frame_TVb = tk.Frame(self.window)
        self.frame_TVb.grid(row=row_in_frame,column=0,padx=5,pady=10)

            #create treeview
        tree = ttk.Treeview(self.frame_TVb, columns=('name','quantity','in_db','in_project'), show='headings')
        #create scrollbar
        scroll = ttk.Scrollbar(self.frame_TVb, orient=tk.VERTICAL, command=tree.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree.configure(yscrollcommand=scroll.set)
        #create columns
        tree.heading('#0', text='ID')
        tree.heading('name', text='Name')
        tree.heading('quantity', text='Quantity')
        tree.heading('in_db', text='DB')
        tree.heading('in_project', text='PRJ')


        #create rows
        tree.column('#0', anchor='center', width=50)
        tree.column('name', anchor='center', width=150)
        tree.column('quantity', anchor='center', width=100)
        tree.column('in_db', anchor='center', width=30)
        tree.column('in_project', anchor='center', width=30)
        #insert rows

        for i,row in self.df.iterrows():
            tree.insert('', 'end', text=str(i), \
                values=(row['name'],row['quantity'],row['in_db'],row['in_project']))
        #pack treeview
        tree.pack()

        button2 = tk.Button(self.frame_TVb, text="Add missing articles to db and project", command=lambda:self.add_missing_articles(self.df,project_ID))
        button2.pack(padx=5,pady=20)

##############################################################################################
###############################################################################################
################################################################################################`
    def add_missing_articles(self,project_ID):
        df_to_db = self.df[self.df['in_db'] == 'X']
        df_in_db = self.df[self.df['in_db'] != 'X']
        #df_to_project = df[df['in_project'] == 'X']

        file = main_folder + '/missing_values/articles_to_db_' + str(datetime.datetime.now())[:10] + '.xlsx'

        #repeat until all articles are added to db   
        if len(df_to_db) > 0:
            df_to_db.to_excel(file)
            file_save_time = datetime.datetime.now()
            print(file_save_time)
            msb.showinfo("Info", "Missing articles saved to excel file\n" \
                    + file + "\n" \
                    + "Please update information , save and press 'OK' to add to db") 
            file_save_time_2 = datetime.datetime.now()           
            while (file_save_time_2 - file_save_time).seconds < 3:
                msb.showinfo("Info", "File is not updated\n" \
                        + file + "\n" \
                        + "Please update information , save and press 'OK' to add to db")

            articles_list = db.get_active_as_list('article')
            #read excel file
            df = pd.read_excel(file)
            df.loc[len(df.description)==0,'description'] = df.name
            db.add_data_from_df_to_table(df,'article')
            self.df = pd.concat([df,df_in_db])

        #add articles to project
        articles_name_ID = db.get_active('article')
        #create ID column from dictionary
        df['ID'] = df.name.map(articles_name_ID)
        #create list from id column
        articles_IDs = df['ID'].tolist()
        #add articles to project
        db.add_articles_to_project(project_ID,articles_IDs)
        self.df = df
        self.list_to_treeview(42)      


    ############################################################################################################     

    def add_from_excel(self,table=''):
        #function append data from excel to database
        #open file dialog to select source file
        init_dir = main_folder + '/excel_data/'
        file = fd.askopenfilename(initialdir = init_dir,title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
        #show popup list of db tables to select
        self.window = tk.Toplevel()
        self.window.title("Select a table")
        self.window.geometry("300x300")
        self.window.resizable(False, False)
        #create listbox
        listbox = tk.Listbox(self.window, width=30, height=15)
        listbox.pack()
        #create scrollbar
        scrollbar = tk.Scrollbar(self.window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        #attach listbox to scrollbar
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        #add items to listbox
        list = db.get_tables_names_in_db()
        for item in list:
            listbox.insert(tk.END, item)
        #create button
        button = tk.Button(self.window, text="Select", command=lambda: self.add_from_excel_2(file,listbox.get(tk.ACTIVE)))
        button.pack()
        #cancel button
        button2 = tk.Button(self.window, text="Cancel", command=self.window.destroy)
        button2.pack()

    def add_from_excel_2(self,file,table):
        self.window.destroy()
        #db_columns = db.get_table_columns_to_list(table)
        df = pd.read_excel(file)

        list_from_df = df['name'].to_list()
        names_list_in_db = db.get_active_from_column('name',table)
        duplicates = list(set(list_from_df).intersection(names_list_in_db))

        #print('#################',names_list_in_db, duplicates)
        if len(list_from_df) == len(duplicates):
            msb.showerror("Error","All items are in database")
            return
        if len(duplicates) > 0:
            df_copy = df.copy()
            df_copy.loc[df_copy['name'].isin(duplicates), 'status'] = 'duplicate'
            df_copy.loc[~df_copy['name'].isin(duplicates), 'status'] = 'added to db'
            df_copy.to_excel(file, index=False)

            df.to_excel(file, index=False)                      
            ask = msb.askyesno("Error","Duplicated items found in db. .\n " + 
                            "Do you want to skip duplicates and continue?\n" + 
                            "if (Yes) - status will be added to excel file\n ")
            if ask :
                df = df[~df['name'].isin(duplicates)]
            else:
                return

        #replace relatives columns to ID (if column name ends with _ID)
        #example if column name is 'finish_ID' replace it with 'ID' from 'finish' table
        missing_items = []
        for column in df.columns:
            if column.endswith('_ID'):
                tab = column[:-3]
                #verify table name exists in db
                if tab not in db.get_tables_names_in_db():
                    msb.showerror('Error', 'Table ' + column + ' not found in db. Verify column names in excel.')
                    return
                # print('###################',tab)
                values_list_in_db = db.get_active_from_column('name',tab)
                values_not_in_db = list(set(df[column].to_list()).difference(values_list_in_db))
                if 0 < len(values_not_in_db):
                    #msb.showwarning("Warning",f"{len(values_not_in_db)} items not in database\n Fullfill missing data in excel file and try again ")
                    #create dataframe with columns from tab table and values from values_not_in_db in name column

                    df2 = pd.DataFrame(columns=db.get_table_columns_to_list(tab))
                    df2['name'] = values_not_in_db
                    df2_columns = [column for column in df2.columns if column not in self.skip_list]
                    #Save in file with timestamp in name
                    df2[df2_columns].to_excel(main_folder + '/missing_values/' + tab + '_' + datetime.now().strftime("%d-%m-%Y_%H-%M-%S") + '.xlsx', index=False)
                    missing_items.append(tab)
                    continue
                print('#################',values_list_in_db, values_not_in_db)
                #dictionary of names and ID from table
                id_from_name_dict = db.get_ID_from_column_as_dict('name',tab)
                #replace names with ID
                df[column] = df[column].map(id_from_name_dict)
            else :
                continue
        if len(missing_items) > 0:
            str_missing_items = ', '.join(missing_items)
            msb.showerror("Error",f"Missing items in database: {str_missing_items}")
            os.startfile(main_folder + '/missing_values/')
            return

        #df.to_excel(main_folder + '/test.xlsx')
        #exit()
        db.add_data_from_df_to_table(df,table)
        

    def table_row_action(self,event):
        item = self.table.item(self.table.selection())
        print(f"Row {item['values'][0]} clicked")

    def orders_list(self):

        try:
            for widget in self.right_frame.winfo_children():
                widget.destroy()    
            self.right_frame.pack_forget()
            self.right_frame.destroy()
        except: pass

        self.right_frame = tk.Frame(self, width=800, height=800)
        self.right_frame.grid(row=0, column=1, sticky="nsew")
        row = tk.Label(self.right_frame, text="orders")
        row.pack()

        df = db.get_active_orders_view(project_name)
        print(df)
        #exit()
        df = df[df.status != 'closed']
        #create date column 'created' with format 'DD-MM-YYY' from 'date_created' that is unix timestamp
        df['date_created'] = pd.to_datetime(df['date_created'], unit='s').dt.strftime('%d-%m-%Y')
        columns = ['ID','name', 'project','job', 'date_created', 'status']
        df = df[columns]
        
        #Create table from dataframe, add to right frame and center columns, add edit button
        self.table = ttk.Treeview(self.right_frame)
        self.table["columns"] = df.columns.to_list() 
        self.table["show"] = "headings"
        narrow_columns = ['ID']
        wide_columns = ['name']
        for column in self.table["columns"]:
            w = 70
            if column in narrow_columns:
                w = 30
            if column in wide_columns:
                w = 300
            self.table.heading(column, text=column)
            self.table.column(column, anchor="center", width=w)
        #edit = {}
        for index, row in df.iterrows():
            self.table.insert("", "end", values=list(row))  

        self.table.pack()
        self.table.bind("<ButtonRelease-1>", self.table_row_action)

    def clear_right_framet(self):
    # Create right frame for table
        try:
            for widget in self.right_frame.winfo_children():
                widget.destroy()
            self.right_frame.pack_forget()
            self.right_frame.destroy()
        except: pass

        self.right_frame = tk.Frame(self, width=800, height=800)
        self.right_frame.grid(row=0, column=1, sticky="nsew")
        row = tk.Label(self.right_frame, text="")
        row.pack()

    def projects_list(self):
    # Create right frame for table
        try:
            for widget in self.right_frame.winfo_children():
                widget.destroy()
            self.right_frame.pack_forget()
            self.right_frame.destroy()
        except: pass

        self.right_frame = tk.Frame(self, width=800, height=800)
        self.right_frame.grid(row=0, column=1, sticky="nsew")
        row = tk.Label(self.right_frame, text="projects list")
        row.pack()

        df = db.get_active_full('project')

        #create date column 'created' with format 'DD-MM-YYY' from 'date_created' that is unix timestamp
        df['date_created'] = pd.to_datetime(df['date_created'], unit='s').dt.strftime('%d-%m-%Y')
        columns = ['name', 'date_created']
        df = df[columns]
        
        # Create table from dataframe, add to right frame and center columns
        table = ttk.Treeview(self.right_frame)
        table["columns"] = df.columns.to_list()
        table["show"] = "headings"
        for column in table["columns"]:
            table.heading(column, text=column)
            table.column(column, anchor="center", width=100)
        for index, row in df.iterrows():
            table.insert("", "end", values=list(row))  
        table.pack()

        # Bind the table to the 'Row Selected' event
        table.bind("<ButtonRelease-1>", self.table_row_action)

    def add_order(self):
        # Initialize tkinter window
        def submit():
            project_ID = list(dict_of_projects.values())[combobox.current()]
            data = {'project_ID': project_ID,'name': name.get(), 'status': 'open'}
            db.add_data(data, table = 'orders')
            window.destroy()

        window = tk.Toplevel()
        window.title("New order form")

        empty_row = tk.Label(window, text="Select project")
        empty_row.pack()

        dict_of_projects = db.get_active('projects')
        print('dict_of_projects',dict_of_projects)
        combobox = ttk.Combobox(window)
        combobox.pack()
        combobox['values'] = [v for v in dict_of_projects.keys()]
        combobox.current(0)
        empty_row = tk.Label(window, text="")
        empty_row.pack()
        empty_row = tk.Label(window, text="Name")
        empty_row.pack()
        #combobox.bind("<<ComboboxSelected>>", show_selection)
        name = tk.Entry(window)
        name.pack()

        submit_button = tk.Button(window, text="Submit", command=submit)
        #submit_button.bind("<Button-1>", submit)
        submit_button.pack()

        cancel_button = tk.Button(window, text="Cancel", command=window.destroy)
        cancel_button.pack()

    def add_project(self):
        def submit_p():
            data = {'name': name.get(), 'description': description.get()}
            db.add_data(data, table = 'projects')
            window.destroy()
        window = tk.Toplevel()
        window.title("New project form")
        empty_row = tk.Label(window, text = "Project name:")
        empty_row.pack()   
        name = tk.Entry(window)
        name.pack()

        empty_row = tk.Label(window, text = "Project description:")
        empty_row.pack() 
        description = tk.Entry(window)
        description.pack()

        submit_button = tk.Button(window, text="Submit", command=submit_p)
        #submit_button.bind("<Button-1>", submit)
        submit_button.pack()

        cancel_button = tk.Button(window, text="Cancel", command=window.destroy)
        cancel_button.pack()

    def show_projects_list(self):
        window = tk.Toplevel()
        window.title("Projects list")
        projects = db.get_active('project')
        for k,v in projects.items():
            empty_row = tk.Label(window, text = v)
            empty_row.pack()

    def show_orders_list(self):
        window = tk.Toplevel()
        window.title("Orders list")
        projects = db.get_active('orders')
        for k,v in projects.items():
            empty_row = tk.Label(window, text = v)
            empty_row.pack()


if __name__ == "__main__":
    #db.init_db()
    
    app = order_app()
