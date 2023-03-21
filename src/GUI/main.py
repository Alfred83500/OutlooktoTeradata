import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from . _tools import retrieve_csv as retr
from . _tools import data_format as dformat
from . _tools import create_fastload as cfast
import win32com.client
import csv
import os
import pathlib


class retrieveMail(ttk.Frame):
    def __init__(self, parent,controller):
        ttk.Frame.__init__(self)
        self.controller = controller
        
        #make the app responsive
        for index in [0, 1, 2]:
            self.columnconfigure(index=index, weight=1)
            self.rowconfigure(index=index, weight=1)


        self.setup_widgets()
        
    
    def actionButtonAttached(self):
        self.controller.AppData["messages"] = retr.retreive_mail_tool(self.entrySender.get())
        self.controller.show_frame(selectAttachedFile)

    def setup_widgets(self):

        # Create a Frame for the inputs
        self.input_frame = ttk.LabelFrame(self, text="Récupération Mail", padding=(200, 100))
        self.input_frame.grid(
            row=0, column=0, padx=(20, 10), pady=(20, 10), sticky="nsew"
        )
        #make the app responsive
        for index in [0, 1,3]:
            self.input_frame.columnconfigure(index=0, weight=1)
            self.input_frame.rowconfigure(index=index, weight=1)

        # Entry Sender
        self.entrySender = ttk.Entry(self.input_frame)
        self.entrySender.insert(0, "TLECORNE@bouyguestelecom.fr")
        self.entrySender.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="nsew")

        # Separator
        self.separator = ttk.Separator(self.input_frame)
        self.separator.grid(row=2, column=0, padx=(20, 10), pady=10, sticky="nsew")
        # Button
        self.button = ttk.Button(self.input_frame, text="Récupérer les mails",
        command = lambda : self.actionButtonAttached())
        self.button.grid(row=3, column=0, padx=5, pady=10, sticky="nsew")

class selectAttachedFile(ttk.Frame):
    def __init__(self,parent,controller):
        ttk.Frame.__init__(self)
        self.controller = controller
        self.attachmentsData = {}
        self.mails_dict = self.controller.AppData['messages'] 
        self.setup_widgets()
    def setup_widgets(self):

        # Panedwindow
        self.paned = ttk.PanedWindow(self)
        self.paned.grid(row=0, column=0, pady=(25, 5), sticky="nsew")



    #TREEVIEW

        # Tree 1 Pane 
        self.treeGivePan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.treeGivePan, weight=1)
        self.treeGivePan.grid(row=1, column=1)
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.treeGivePan)
        self.scrollbar.pack(side="right", fill="y")
        self.treeviewGive = ttk.Treeview(
            self.treeGivePan,
            selectmode="browse",
            yscrollcommand=self.scrollbar.set,
            columns=(1,2),
            height=10,
        )
        self.treeviewGive.pack(expand=True, fill="both")
        self.scrollbar.config(command=self.treeviewGive.yview)
        # treeviewGive columns
        self.treeviewGive.column("#0", anchor="w", width=120)
        self.treeviewGive.column(1, anchor="w", width=120)
        self.treeviewGive.column(2, anchor="w", width=120)
        # treeviewGive headings
        self.treeviewGive.heading("#0", text="Emetteur", anchor="center")
        self.treeviewGive.heading("1", text="Nom de la pièce jointe", anchor="center")
        self.treeviewGive.heading("2", text="Date", anchor="center")

        item = list(self.mails_dict)
        len_item = len(item)
        for i in range(len(list(self.mails_dict))-1,-1,-1):
            self.treeviewGive.insert(
                "", iid=len_item -i, index="end", text=str(item[i].SenderName).replace(" ","_"), values= (str(item[i].Subject).replace(" ", "_"), str(item[i].ReceivedTime))
                )
           
            j = 1
            
            for attachment in list(item[i].Attachments):
                
                if str(attachment).split('.')[-1] == 'csv' or str(attachment).split('.')[-1] == 'txt':
                    
                    self.attachmentsData[str(str(attachment.Filename)+str(item[i].ReceivedTime)).replace(" ","")] = attachment 
                    self.treeviewGive.insert(len_item-i, iid=f'{len_item -i}.{j}', index = "end", text=str(attachment.Filename), values =(str(item[i].ReceivedTime),"") )
                    j+=1
                
        

        # Tree 2 Pane 
        self.treePan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.treePan, weight=1)
        self.treePan.grid(row=1, column=3, padx=20)
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.treePan)
        self.scrollbar.pack(side="right", fill="y")
        self.treeviewReceive = ttk.Treeview(
            self.treePan,
            selectmode="browse",
            yscrollcommand=self.scrollbar.set,
            columns=(1,2),
            height=10,
        )
        self.treeviewReceive.pack(expand=True, fill="both")
        self.scrollbar.config(command=self.treeviewReceive.yview)
        # treeviewReceive columns
        self.treeviewReceive.column("#0", anchor="w", width=120)
        self.treeviewReceive.column(1, anchor="w", width=120)
        
        # treeviewReceive headings
        self.treeviewReceive.heading("#0", text="Nom de la pièce jointe", anchor="center")
        self.treeviewReceive.heading("1", text="Date", anchor="center")
        
        
      
            
        
            #Pane 2
        self.pane_2 = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.pane_2, weight=3)
        self.pane_2.grid(row=1, column=2)
        def get_selection():     
            self.treeviewReceive.insert(
                "", iid=i, index="end", text=(self.treeviewGive.item(self.treeviewGive.selection())['text']), values =(str(self.treeviewGive.item(self.treeviewGive.selection())['values'][0]))
                )
            pass

        self.buttonAdd = ttk.Button(self.pane_2, text="Ajouter", command=get_selection)
        self.buttonAdd.grid(row=1,column=0)



        

        def next_page():
            attachFile = self.attachmentsData.get(str(str(self.treeviewGive.item(self.treeviewGive.selection())['text'])+ 
                str(self.treeviewGive.item(self.treeviewGive.selection())['values'][0])).replace(" ",""))
            retr.save_Attachement(attachFile)
            print(f"src/GUI/data/{attachFile.FileName}")
            self.controller.AppData["attachement_name_selected"] = os.path.join(pathlib.Path().resolve(), f"src/GUI/data/{attachFile.FileName}")
            self.controller.show_frame(showData)
            pass



        #Button Next Pan
        self.nextButtonPan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.nextButtonPan, weight=3)
        self.nextButtonPan.grid(row=3, column=2)
        self.buttonNext = ttk.Button(self.nextButtonPan, text="Valider", command=next_page)
        self.buttonNext.grid(row=0,column=0)


        def previous_page():
            self.controller.show_frame(retrieveMail)


        #Button Previous Pan
        self.previousButtonPan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.previousButtonPan, weight=3)
        self.previousButtonPan.grid(row=3, column=1)
        self.buttonPrevious = ttk.Button(self.previousButtonPan, text="Précedent", command=previous_page)
        self.buttonPrevious.grid(row=0,column=0)

        def deleteRow():
            print(self.treeviewReceive.item(self.treeviewReceive.selection()))
            self.treeviewReceive.delete(self.treeviewReceive.selection())

        #Button Delete
        
        self.buttonDel = ttk.Button(self.pane_2, text="Effacer", command=deleteRow)
        self.buttonDel.grid(row=3,column=0, pady=10)

class showData(ttk.Frame):
    def __init__(self, parent,controller):
        ttk.Frame.__init__(self)
        self.controller = controller
        self.parsed_file =""
        
        #make the app responsive
        for index in [0, 1, 2]:
            self.columnconfigure(index=index, weight=1)
            self.rowconfigure(index=index, weight=1)

        self.setup_widgets()


    def select_file(self):
        filename = fd.askopenfilename()
        self.controller.AppData["log_on_file"] = filename
    
    def setup_widgets(self):

        def next_page():
            print()
            cfast.create_fastload_execute(self.parsed_file,self.entryTableName.get(),self.controller.AppData["log_on_file"])
            self.controller.show_frame(validation)
            pass

         # Panedwindow
        self.paned = ttk.PanedWindow(self)
        self.paned.grid(row=0, column=0, pady=(25, 5), sticky="nsew")
        self.parsed_file = dformat.format_data(self.controller.AppData["attachement_name_selected"])


        with open(self.parsed_file) as csv_file:
            reader = csv.reader(csv_file, delimiter=';')
            data = list(reader)

        
        
        ################### SHOWING CSV ########################


        # # TreeView Pan 
        self.treePan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.treePan, weight=1)
        self.treePan.grid(row=0, column=0,columnspan=3)
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.treePan)
        self.scrollbar.pack(side="right", fill="y")
        self.treeShowData = ttk.Treeview(
            self.treePan,
            selectmode="browse",
            yscrollcommand=self.scrollbar.set,
            height=10,
        )
        self.treeShowData.pack(expand=True, fill="both")
        self.scrollbar.config(command=self.treeShowData.yview)


        headers = data[0]
        self.treeShowData["columns"] = headers
    
        self.treeShowData.column("#0", width=0, anchor='w')
        for i in headers:
            self.treeShowData.column(i, width=100, anchor='w')

        for i in headers: 
            self.treeShowData.heading(i,text=i, anchor = 'center')

        for row in data[1:20]:
            self.treeShowData.insert("","end", values = row)

            # Create a Frame for the inputs
        self.input_frame = ttk.Frame(self.paned, padding=(200, 100))
        self.input_frame.grid(
            row=1, column=1, padx=(20, 10), pady=(20, 10), sticky="nsew"
        )
        # Entry Object
        self.entryTableName = ttk.Entry(self.input_frame)
        self.entryTableName.insert(0, "Nom de la Table créée (si plusieurs table incrémentation avec 0)")
        self.entryTableName.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="nsew")


        # Button
        self.buttonLOGON = ttk.Button(self.input_frame, text="sélectionner le fichier LOGON",
        command = self.select_file)
        self.buttonLOGON.grid(row=3, column=0, padx=5, pady=10, sticky="nsew")
            
        
        #Button Next Pan
        self.nextButtonPan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.nextButtonPan, weight=3)
        self.nextButtonPan.grid(row=3, column=1)
        self.buttonNext = ttk.Button(self.nextButtonPan, text="Valider", command=next_page)
        self.buttonNext.grid(row=1,column=1)


        def previous_page():
            self.controller.show_frame(selectAttachedFile)

        #Button Previous Pan
        self.previousButtonPan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.previousButtonPan, weight=3)
        self.previousButtonPan.grid(row=3, column=0)
        self.buttonPrevious = ttk.Button(self.previousButtonPan, text="Précedent", command=previous_page)
        self.buttonPrevious.grid(row=0,column=0)

class validation(ttk.Frame):
    def __init__(self, parent,controller):
        ttk.Frame.__init__(self)
        self.controller = controller
        
        #make the app responsive
        for index in [0, 1, 2]:
            self.columnconfigure(index=index, weight=1)
            self.rowconfigure(index=index, weight=1)

        self.setup_widgets()
    

    def setup_widgets(self):
        def create_table():
            print(os.path.abspath('fich_modif_todate.txt'))
            cmd_fastload =  f"fastload < {os.path.abspath('fich_modif_todate.txt')} >> log/log_fastload.txt "
            print("############### DEBUT FASTLOAD ##########################")
            os.system(cmd_fastload)
            print("############### FIN FASTLOAD ##########################")
            with open('log/log_fastload.txt', 'r') as f:
                last_line = f.readlines()[-2]
            self.labelResults.config(text = last_line)

        # Panedwindow
        self.paned = ttk.PanedWindow(self)
        self.paned.grid(row=0, column=0, pady=(25, 5), sticky="nsew")

        self.labelResultPan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.labelResultPan, weight=3)
        self.labelResultPan.grid(row=3, column=0)
        self.labelResults = ttk.Label(self.labelResultPan, text=" ")
        self.labelResults.grid(row=0,column=1)

        self.nextButtonPan = ttk.Frame(self.paned, padding=5)
        self.paned.add(self.nextButtonPan, weight=3)
        self.nextButtonPan.grid(row=3, column=0)
        self.buttonNext = ttk.Button(self.nextButtonPan, text="Je valide", command=create_table)
        self.buttonNext.grid(row=0,column=1)


class App(tk.Tk):
     
    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
         
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
         
        # creating a container
        self.container = container = tk.Frame(self) 
        container.grid(row = 0, column = 0, sticky ="nsew")
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
  
        # initializing frames to an empty array
        self.frames = {} 
        
        self.AppData = {
            "messages": win32com.client.CDispatch,
            "log_on_file":"Test",
            "attachement_name_selected": str
        }
    
        self.show_frame(retrieveMail)

    # to display the current frame passed as
    #  
    def show_frame(self, controller):
        if controller not in self.frames:
            print(controller)
            self.frames[controller] = frame = controller(self.container, self)
            print(frame)
            print(self.frames[controller])
            frame.grid(row=0, column=0, sticky="nsew")
        else:
            frame = self.frames[controller]
            if frame != retrieveMail:
                self.frames[controller] = frame = controller(self.container, self)
                frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()


