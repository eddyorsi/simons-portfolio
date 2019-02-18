import pandas as pd
from gspread_pandas import Spread
from tkinter import *
from tkinter.filedialog import askopenfilename
import os

# Initialize UI
class SampleApp(Tk):
    def __init__(self):
        Tk.__init__(self)
        self._frame = None
        self.switch_frame(StartPage)
        
# Destroy existing frame and create new one to switch between frames
    def switch_frame(self, frame_class, optional=False):
        if optional:
            new_frame = frame_class(self, 2)
        else:
            new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.grid()

# Initialize Start Page
class StartPage(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.cde = Button(self, text='Click Here', command=lambda: master.switch_frame(Download))
        self.cde.grid(row=3, column=1)
        self.init_window()

    def init_window(self):
        self.master.title("Update HR Log")
        self.grid()
        
# Ask user to select file to cross reference with existing google sheets
class Download(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        download_label = Label(self, text="Select SBF Report", justify=LEFT)
        download_label.grid(row=1, column=0, sticky=W)
        new_directory = os.path.expanduser("~") + "/Downloads/"
        filenames=[]
        # Show file name
        text = Text(self, state='disabled', width=50, height=1)
        text.configure(state='normal')
        text.insert('end', new_directory)
        text.configure(state='disabled')
        text.grid(row=1, column=4)
		# Create necessary buttons
        browse_button = Button(self, text="Browse", command=lambda: file_browser("Select SBF Report", text))
        browse_button.grid(row=1, column=2, padx=(10, 15), pady=(15, 10))
        
        go_label = Label(self, text="Start Program")
        go_label.grid(row=11, column=0)
        
        cont = Button(self, text="GO!", command=lambda: updatesheet(opensheet(), text.get(1.0, END)))
        cont.grid(row=11, column=2)
        
		# Create function to browse and select files    
        def file_browser(msg, text):
            new_directory = askopenfilename(title=msg)
            text.configure(state='normal')
            text.delete(1.0, 'end')
            text.insert(1.0, new_directory)
            text.configure(state='disabled')
		
		# Open google sheets and select 'HR TVL LOG'
        def opensheet():

            creds=Spread('Update Log',"HR TVL Log")
            return creds
            
		# Cross reference with selected file, update paid and unpaid requisitions
        def updatesheet(HR,sbffile):

            HR.open_sheet('Unpaid')
            sbffile = sbffile.strip('\n')
            df=HR.sheet_to_df()
			# Read in excel file
            sbf=pd.read_excel(sbffile,skiprows=2)
            df=df.reset_index().rename(columns={'index':'Last Name','Item Description ':'Item Description'})
            
			# Get TVL number from Item Description column
            df["TVL"]=df["Item Description"].str.split(" ",expand=True)[1]
            
			# Find intersection of requisitions
            inter=df[df['TVL'].isin(sbf['Requisition/CPV'])]
            inter['Amount']=pd.to_numeric(inter['Amount'])
            HR.open_sheet('Paid')
            paid=HR.sheet_to_df()
            paid=paid.reset_index().rename(columns={'index':'Last Name'})

            both = pd.DataFrame()
            one = pd.DataFrame()
            two = pd.DataFrame()
            three = pd.DataFrame()
			
			# Make sure both the TVL number and amount match
            one["TVL"] = inter['TVL']
            one['Amount'] = inter['Amount']
            one.reset_index().drop(columns='index')
            two["TVL"] = sbf['Requisition/CPV']
            two["Amount"] = sbf['Invoice Amount']
            three = pd.merge(right=one, left=two, how='inner')
            both = inter[inter["TVL"].isin(three['TVL'])]
			
			# Update paid list with newly paid items
            order=list(df)
            newpaid = pd.concat([paid,both],axis=0,ignore_index=True).reindex(columns=order)
            newunpaid = df[~df['TVL'].isin(inter['TVL'])]
			
			# Find items where amounts don't not match
            matchmis=inter[~inter['TVL'].isin(both["TVL"])]
            matchmis = matchmis.reset_index().drop(columns='index')
            HR.open_sheet('Amount mismatch')
            toconcat = HR.sheet_to_df()
            toconcat = toconcat.reset_index()
            rang = []
            for i in range(len(toconcat), len(toconcat) + len(matchmis), 1):
                rang.append(i)
            matchmis = matchmis.set_index(pd.Index(rang))
            mismatch = pd.concat([toconcat, matchmis], axis=0, ignore_index=True).reindex(columns=order)
            
            # Update Google Sheets Paid and Unpaid lists
            Spread.df_to_sheet(HR,sheet="Paid",df=newpaid,replace=True,index=False)
            Spread.df_to_sheet(HR,sheet="Unpaid",df=newunpaid,replace=True,index=False)
			# Send newly paid items to Justpaid sheet
            Spread.df_to_sheet(HR,sheet='Justpaid',df=both,replace=True,index=False)
			# Send mismatch items to Amount mismatch sheet
            Spread.df_to_sheet(HR,sheet='Amount mismatch',df=mismatch,replace=True,index=False)
# Run UI
if __name__ == "__main__":
    app = SampleApp()
    app.geometry("700x350")
    app.mainloop()
