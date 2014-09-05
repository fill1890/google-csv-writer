### Written in Python 2.7. The gdata module requires installation.

### Modules:

# Attempt to import GData, and display an error if it fails.
try:
    import gdata.spreadsheet.service
    from gdata.service import BadAuthentication
except:
    print('Unable to access the gdata module.')

# Attempt to import other dependencies. These should be included with Python.
try:
    import csv
    from Tkinter import *
    import ttk
    import tkFileDialog
    from os import system
    from os.path import expanduser
except:
    print('Python was unable to import required modules, that should have been installed by Python.')

### Functions & Classes: googleData class, contains all google-related code,
### sheetManager class, manages the spreadsheet, GUI class, manages GUI

class googleData:
    # Login function. Authenticates 
    def login(self, username, password, error, callback, **kwargs):
        # Initialize spreadsheet service
        self.client = gdata.spreadsheet.service.SpreadsheetsService()
        # Get user details
        self.client.email = username
        self.client.password = password
        # Log in
        try:
            self.client.ProgrammaticLogin()
            # Run the given callback function, with arguments, if successful
            if callback:
                callback(**kwargs)
        # Run the given error function if authentication fails
        except BadAuthentication:
            error()
        # Catch other errors
        except:
            print('Unable to connect to Google. This could be caused by an internet connect problem.')
        
        
    # sheetFeed function. Gets spreadsheet feed, runs when logged in
    def sheetFeed(self):
        # Clear list of spreadsheets to avoid repeated items
        cur_items = GUI.elements.spreadsheetTree.get_children()
        for item in cur_items:
            GUI.elements.spreadsheetTree.delete(item)
        # Get spreadsheet feed
        self.sheet_feed = sheet_feed = self.client.GetSpreadsheetsFeed()
        # Add available spreadsheets to the GUI list using GUI.addSpreadsheet
        for i in range(0, len(sheet_feed.entry)):
            GUI.addSpreadsheet(sheet_feed.entry[i], i)
        return None

    # worksheetFeed function. Gets worksheet feed, runs when spreadsheet is selected
    def worksheetFeed(self):
        # Clear list of worksheets to avoid repeated items
        cur_items = GUI.elements.worksheetTree.get_children()
        for item in cur_items:
            GUI.elements.worksheetTree.delete(item)
        # Get selected sheet
        selected_sheet = GUI.elements.spreadsheetTree.selection()[0]
        # Extract id using corresponding data in the sheet feed
        self.sheet_id = self.sheet_feed.entry[int(selected_sheet)].id.text.split('/')[-1]
        # Get worksheet feed according to sheet id
        self.wk_feed = wk_feed = self.client.GetWorksheetsFeed(self.sheet_id)
        # Add available spreadsheets to the GUI list using GUI.addWorksheet
        for i in range(0, len(wk_feed.entry)):
            GUI.addWorksheet(wk_feed.entry[i], i)
        return None

    # cellsFeed function. Gets cell feed, runs when worksheet is selected
    def cellFeed(self):
        # Get selected worksheet
        selected_wksheet = GUI.elements.worksheetTree.selection()[0]
        # Extract id using corresponding data in the worksheet feed
        self.wk_id = self.wk_feed.entry[int(selected_wksheet) - 1].id.text.split('/')[-1]
        # Get cell feed according to sheet and worksheet id
        cell_feed = self.client.GetCellsFeed(self.sheet_id, self.wk_id)
        # Set class 'out' property to the cell feed
        self.out = cell_feed
        # Enable 'Write CSV' button in GUI
        GUI.elements.writeCSVButton.state(['!disabled'])
        return None

class sheetManager:
    # constructSheet function. Constructs a sheet from cell feed, called by writeCSV.
    def constructSheet(self):
        # Get cell feed from googleData 'out' property
        cells_feed = googleData.out
        # Initialise sheet list
        self.sheet = sheet = []
        # Get row and column to start writing sheet at
        min_column = GUI.variables.startCol.get() - 1
        min_row = GUI.variables.startRow.get() - 1
        # Loop through cells in cell feed
        for cell in cells_feed.entry:
            # Set column and row to current column and row
            column = int(cell.cell.col)
            row = int(cell.cell.row)
            # Ensure cell is within writing area
            if column > min_column and row > min_row:
                # Create new list inside sheet if new column, otherwise
                # append to column
                cur_func = lambda: sheet.append([cell.cell.text]) if len(sheet) < column else sheet[column - 1].append(cell.cell.text) 
                cur_func()

    # WriteCSV function. Runs when 'Write CSV' button is clicked
    def writeCSV(self):
        # Use constructSheet to create the sheet to be written
        self.constructSheet()
        # Get home folder
        home = expanduser('~')
        # Ask for file name and path to write
        newFile = tkFileDialog.asksaveasfilename(defaultextension='csv', initialdir=home, title='Save As')
        sheet = self.sheet
        # Write file if path is valid
        if newFile != '':
            # Create file if necessary
            system("touch " + newFile)
            # Open file and start CSV writer
            csvFile = open(newFile, 'w')
            writer = csv.writer(csvFile)
            # Loop through cells and add associated row to CSV file
            # This will be extended later to use the longest column and add empty cells where necessary
            for cell in range(0, len(sheet[0])):
                output = []
                for x in range(0, len(sheet) - 1):
                    try:
                        output.append(sheet[x][cell])
                    except:
                        pass
                writer.writerow(output)
            # Save and close file (Necessary in Python 2)
            csvFile.close()

# GUI class. Contains GUI functions and information.
class GUI:
    # Create empty element and variable classes
    class elements: pass
    class variables: pass
    # Login class. Contains information for login window
    class login:
        # Create empty element and variable classes
        class elements: pass
        class variables: pass
        # Construct function. Creates window
        def construct(self):
            self.top = Toplevel()
            self.create()
        # Create function. Creates elements and binds events
        def create(self):
            self.root = ttk.Frame(self.top)
            username = self.variables.username = StringVar()
            password = self.variables.password = StringVar()
            self.elements.userLabel = ttk.Label(self.root, text='Username:')
            self.elements.userEntry = ttk.Entry(self.root, textvariable=self.variables.username)
            self.elements.passLabel = ttk.Label(self.root, text='Password:')
            self.elements.passEntry = ttk.Entry(self.root, textvariable=self.variables.password, show="*")
            self.elements.loginButton = ttk.Button(self.root, text='Login', default='active', command=lambda:googleData.login(username.get(), password.get(), self.loginError, self.finishLogin, caller=self))
            self.elements.cancelButton = ttk.Button(self.root, text='Cancel', command=lambda:self.top.destroy())
            self.elements.error = ttk.Label(self.root, text='Incorrect username/password', foreground='red')
            self.root.bind_all('<Return>', lambda _:self.elements.loginButton.invoke())
            self.build()
        # Build function. Adds elements to window
        def build(self):
            self.elements.userLabel.grid(column=0, row=0)
            self.elements.userEntry.grid(column=1, row=0, columnspan=2)
            self.elements.userEntry.focus()
            self.elements.passLabel.grid(column=0, row=1)
            self.elements.passEntry.grid(column=1, row=1, columnspan=2)
            self.elements.loginButton.grid(column=1, row=2)
            self.elements.cancelButton.grid(column=2, row=2)
            self.root.grid()
        # FinishLogin function. used as login callback.
        def finishLogin(_, caller):
            # Set status to 'Logged in'
            GUI.variables.login_state_label.set('Logged in')
            # Destroy top window
            caller.top.destroy()
            # Remove 'Log in' button
            GUI.elements.loginButton.grid_remove()
            # Run sheetFeed function
            googleData.sheetFeed()
        # LoginError function. Used as error function
        def loginError(self):
            # Ad the label indicating an error to the login window
            self.elements.error.grid(column=1, columnspan=2, row=3)
    # Construct function. Creates window
    def construct(self, top):
        self.root = ttk.Frame(top)
        self.create()
    # Create function. Creates elements, binds events, and sets initial states where necessary
    def create(self):
        self.variables.login_state_label = login_state_label = StringVar()
        self.variables.startRow = startRow = IntVar()
        startRow.set(1)
        self.variables.startCol = startCol = IntVar()
        startCol.set(1)
        login_state_label.set('Not logged in to Google')
        self.elements.loginLabel = ttk.Label(self.root, textvariable=login_state_label)
        self.elements.loginButton = ttk.Button(self.root, text='Log in', command=lambda:self.login().construct())
        self.elements.spreadsheetLabel = ttk.Label(self.root, text='Spreadsheets')
        self.elements.worksheetLabel = ttk.Label(self.root, text='Worksheets')
        self.elements.spreadsheetTree = ttk.Treeview(self.root, height='15', show='tree', selectmode='browse')
        self.elements.worksheetTree = ttk.Treeview(self.root, height='15', show='tree', selectmode='browse')
        self.elements.spreadsheetTree.bind('<<TreeviewSelect>>', lambda _:googleData.worksheetFeed())
        self.elements.worksheetTree.bind('<<TreeviewSelect>>', lambda _:googleData.cellFeed())
        self.elements.startRowLabel = ttk.Label(self.root, text='Start at row:')
        self.elements.startColLabel = ttk.Label(self.root, text='Start at column:')
        vcmd = (self.root.register(self.onValidate), '%S')
        self.elements.startRowEntry = ttk.Entry(self.root, validate='key', validatecommand=vcmd, textvariable=startRow, width=4)
        self.elements.startColEntry = ttk.Entry(self.root, validate='key', validatecommand=vcmd, textvariable=startCol, width=4)
        self.elements.writeCSVButton = ttk.Button(self.root, text='Write CSV', command=sheetManager.writeCSV)
        self.elements.writeCSVButton.state(['disabled'])
        self.build()
    # Build function. Adds elements to window
    def build(self):
        self.elements.loginLabel.grid(column=0, row=0, sticky=W)
        self.elements.loginButton.grid(column=1, row=0)
        self.elements.spreadsheetLabel.grid(column=0, row=1, columnspan=2)
        self.elements.spreadsheetTree.grid(column=0, row=2, padx=10, pady=10, columnspan=2)
        self.elements.worksheetLabel.grid(column=2, row=1)
        self.elements.worksheetTree.grid(column=2, row=2, padx=10, pady=10)
        self.elements.startRowLabel.grid(column=0, row=3)
        self.elements.startRowEntry.grid(column=1, row=3)
        self.elements.startColLabel.grid(column=0, row=4)
        self.elements.startColEntry.grid(column=1, row=4)
        self.elements.writeCSVButton.grid(column=2, row=3, rowspan=2)
        self.root.grid()
    # addSpreadsheet function. Adds spreadsheet item to GUI. Uses lambda to be compact.
    addSpreadsheet = lambda self, sheet, iid: self.elements.spreadsheetTree.insert('', 'end', iid, text=sheet.title.text)
    # addWorksheet function. Adds worksheet item to GUI. Uses lambda to be compact.
    addWorksheet = lambda self, wk_sheet, iid: self.elements.worksheetTree.insert('', 'end', iid + 1, text=wk_sheet.title.text)
    # onValidate function. Validates text entry for start column and row.
    # Only allows numbers to be entered
    def onValidate(_, S):
        try:
            int(S)
            return True
        except ValueError:
            return False

# If running as primary program, run
if __name__ == '__main__':
    # Contruct window
    top = Tk()
    top.title('CSV Writer')
    # Initialize classes
    GUI = GUI()
    googleData = googleData()
    sheetManager = sheetManager()
    # Create GUI
    GUI.construct(top)
    # Start event loop
    top.mainloop()
