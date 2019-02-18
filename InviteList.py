import textract
import re
import pandas as pd
import os
from tkinter import *
from tkinter.filedialog import askopenfilename
from pathlib import Path

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
        self.cde = Button(self, text='Begin Here', command=lambda: master.switch_frame(Download))
        self.cde.grid(row=3, column=1)
        self.init_window()

    def init_window(self):
        self.master.title("Retrieve Danfords Invite List")
        self.grid()

# Ask user to select file to PDF file to read
class Download(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        download_label = Label(self, text="Select Danford's PDF", justify=LEFT)
        download_label.grid(row=1, column=0, sticky=W)
        new_directory = os.path.expanduser("~") + "/Downloads/"
        
        # Show filename and add necessary buttons
        filenames = []
        text = Text(self, state='disabled', width=50, height=1)
        text.configure(state='normal')
        text.insert('end', new_directory)
        text.configure(state='disabled')
        text.grid(row=1, column=4)
        browse_button = Button(self, text="Browse", command=lambda: file_browser("Point to Download directory!", text))
        browse_button.grid(row=1, column=2, padx=(10, 15), pady=(15, 10))

        go_label = Label(self, text="Start Program")
        go_label.grid(row=11, column=0)

        cont = Button(self, text="Continue", command=lambda: readpdf(text.get(1.0, END)))
        cont.grid(row=11, column=2)
		
		# Function to browse and select files
        def file_browser(msg, text):
            new_directory = askopenfilename(title=msg)
            text.configure(state='normal')
            text.delete(1.0, 'end')
            text.insert(1.0, new_directory)
            text.configure(state='disabled')

		# Read PDF and pull relevant information
        def readpdf(path):
            path=path.strip('\n')
			# Read file
            soup = textract.process(path, method='tesseract', language='eng')
            string=str(soup)
			# Search for names in between date and some other words
            maybe=re.findall('[0-9][0-9]-[0-9][0-9]-[0-9][0-9] [0-9]* Room Charge.*?Of Room',string)
            dates=[]
            names=[]
            for i in range(0,len(maybe),1):
                datesearch=re.findall("[0-9][0-9]-[0-9][0-9]-[0-9][0-9] [0-9]*",maybe[i])
                namesearch=re.findall('(?<=From )(.*)(?= Of)',maybe[i])
                
                for k in range(0,len(datesearch),1):
                    ind=datesearch[k].split(' ')
                dates.append(ind[0])
                names.append(namesearch)
			# Make dataframe with all names (not unique)
            fin=pd.DataFrame(columns=['Names','Start Date','End Date'])
            fin['Names']=names

			# Find first and last instance of each name, retrieve start and end dates
			#with those indices
            def findI(lst,string):
                indices = [i for i, x in enumerate(lst) if x == string]
                start=min(indices)
                end=max(indices)
                return start, end

            for i in range(0,len(fin),1):
                index=findI(names,names[i])
                fin.at[i,'Start Date']=dates[min(index)]
                fin.at[i,'End Date']=dates[max(index)]
			
			# Store dates and unique names in dataframe, save to desktop
            result=fin.loc[fin.astype(str).drop_duplicates().index]
            result=result.reset_index().drop(columns='index')
            result['Names']=result['Names'].astype(str).str.strip("['']")

            out_path = str(Path.home())+"/desktop/InviteList.xlsx"
            writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
            result.to_excel(writer)
            writer.save()

# Run UI
if __name__ == "__main__":
    app = SampleApp()
    app.geometry("700x350")
    app.mainloop()