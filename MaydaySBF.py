from tkinter import *
from tkinter.filedialog import askopenfilename
import os
import pandas as pd
import numpy as num
import selenium
from selenium import webdriver
import re
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
        self.master.title("SBF vs Mayday")
        self.grid()

# Ask user to select two files to cross reference invoices.
class Download(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        
        # Show filename and add necessary buttons
        download_label = Label(self, text="Select Mayday Report", justify=LEFT)
        download_label.grid(row=1, column=0, sticky=W)
        new_directory = os.path.expanduser("~") + "/Downloads/"

        filenames=[]
        text = Text(self, state='disabled', width=50, height=1)
        text.configure(state='normal')
        text.insert('end', new_directory)
		text.configure(state='disabled')
		text.grid(row=1,column=4)
		
        browse_button = Button(self, text="Browse", command=lambda: file_browser("Point to Download directory!", text))
        browse_button.grid(row=1, column=2, padx=(10, 15), pady=(15, 10))
        
        browse_button2 = Button(self, text="Browse", command=lambda: file_browser("Point to Download directory!",text2))
        browse_button2.grid(row=3,column=2)
        
        download_label2=Label(self, text="Select SBF Statement Detail")
		download_label2.grid(row=3,column=0)

        text2 = Text(self, state='disabled', width=50, height=1)
        text2.configure(state='normal')
        text2.insert('end', new_directory)
        text2.configure(state='disabled')
        text2.grid(row=3, column=4)
        
        go_label=Label(self,text="Start Program")
        go_label.grid(row=11, column=0)

        go_button = Button(self, text="GO!",command=lambda: seleniumpart(text.get(1.0,END),text2.get(1.0,END)))
        go_button.grid(row=11, column=2, sticky=E + W + S + N, padx=(10, 15))

        var = IntVar()
        
        go_on = Button(self, text="Continue", command=lambda: var.set(1))
        go_on.grid(row=15,column=2)
        go_on_label=Label(self,text="Sign in and click Continue")
        go_on_label.grid(row=15,column=0)

		# Function to browse and select files
        def file_browser(msg, text):
            new_directory = askopenfilename(title=msg)
            text.configure(state='normal')
            text.delete(1.0, 'end')
            text.insert(1.0, new_directory)
            text.configure(state='disabled')

		# Function to cross reference files and input into database
        def seleniumpart(file,file2):
            # Get filenames and open files
            file=file.strip('\n')
            file2=file2.strip('\n')
            np = pd.read_excel(file, encoding='cp1252', skiprows=8, header=0)
            ap = pd.read_excel(file2, encoding='cp1252', skiprows=5, header=0)
            
            # Drop exra columns from both dataframes
            ap = ap.drop(columns=['Unnamed: 0', 'Unnamed: 4', 'Unnamed: 11', 'Unnamed: 14'], axis=1)
            np = np.drop(columns=['Unnamed: 0'], axis=1)

            # Extract Invoices in preparation for intersection
            notpaidlist = list(np['Invoice #'])
            paidlist = list(ap['Invoice'])

            # Find intersection of invoices
            def intersection(lst1, lst2):
                lst3 = [str(value) for value in lst1 if str(value) in lst2]
                return lst3

            inter = intersection(notpaidlist, paidlist)

            # Retrieve dataframe with invoices only from intersection of not-paid list and paid list
            dates = ap[ap['Invoice'].isin(inter)]

            # Select relevant columns
            checkdates = dates.iloc[:, [4, 7, 9]]

            # Remove rows with no check dates for Mayday input
            withnull = checkdates.replace(r'\s+', num.nan, regex=True)
            nonull = withnull[withnull['Check Dt'].notnull()]

            # Retrieve invoices with missing check dates
            missing = dates[dates['Check Dt'].isnull()].drop_duplicates('Invoice')

            # Keep only unique invoices
            fin = nonull.drop_duplicates('Invoice').reset_index().drop(columns='index')

            # Make check dates a string instead of datettime for inputing into mayday
            fin['Check Dt'] = fin['Check Dt'].dt.strftime('%Y-%m-%d')
            
			# Open Chrome Driver
            drivepath=str(Path.home()) + "/desktop/EddysApps/chromedriver"
            driver = webdriver.Chrome(drivepath)

            # Login into Mayday and open Accounting
            def LoginAccounting():
                # Open Mayday
                driver.get('http://mayday.ad.scgp.stonybrook.edu/login.php')

                # Waiting for robot checker to be complete
                go_on.wait_variable(var)
                
                # Open accounting
                payable = driver.find_elements_by_xpath('//*[@class="ThemeOffice2003MainFolderText"]')
                payable[4].click()

            LoginAccounting()

            itr = len(fin)
            errorcount = 0
            entercount = 0

            # Creates empty error and entered list
            errors = pd.DataFrame(columns=['Amount', 'Invoice', 'Check Dt'])
            entered = pd.DataFrame(columns=['Amount', 'Invoice', 'Check Dt'])

            # Inputing data into Mayday
            for i in range(0, itr, 1):
                # Search by not paid, invoice number and amount, submit search
                driver.find_element_by_xpath("//select[@id='code_invoice_status']/option[text()='Not Paid (N)']").click()
                driver.find_element_by_xpath('//*[@id="invoice_number"]').send_keys(str(fin.loc[i][1]))
                driver.find_element_by_xpath('//*[@id="amount"]').send_keys(str(fin.loc[i][0]))
                driver.find_elements_by_xpath('//*[@type="submit"]')[0].click()

                # Find all links with specified information and select first link
                try:
                    possibilities = driver.find_elements_by_xpath("//a[contains(@href,'invoice.php?id')]")
                    possibilities[0].click()

                # If there is no link, go back to search menu and try next item
                except:
                    # Return to search, clear search.
                    driver.find_element_by_xpath('//*[@type="button"]').click()
                    driver.find_element_by_xpath('//*[@type="button"]').click()

                    # Add failure to list of faulty invoices
                    errors.loc[errorcount] = fin.loc[i]
                    errorcount = errorcount + 1
                    continue

                # Input date paid, and invoice status to paid, submit.
                driver.find_element_by_xpath('//*[@id="date_paid"]').send_keys(str(fin.loc[i][2]))
                driver.find_element_by_xpath('//*[@id="date_paid"]').click()
                driver.find_element_by_xpath("//select[@id='code_invoice_status']/option[text()='Paid (P)']").click()
                driver.find_element_by_xpath('//*[@type="submit"]').click()

                # Return to search, clear search.
                driver.find_element_by_xpath('//*[@type="button"]').click()
                driver.find_element_by_xpath('//*[@type="button"]').click()

                # Add success to list of entered invoices
                entered.loc[entercount] = fin.loc[i]
                entercount = entercount + 1

            entinv = list(entered['Invoice'])
            newint = intersection(entinv, notpaidlist)
            notentered = np[~np['Invoice'].isin(newint)]

            # This list contains invoices that match a requisition number from Mayday
            npnull = np[~np['Requisition'].isnull()]
            apnull = ap[~ap["Invoice"].isnull()]
            npreq = list(npnull['Requisition'])
            apinv = list(apnull['Invoice'])
            reqinter = [P for P in npreq if P in apinv]
            reqinv = ap[ap['Invoice'].isin(reqinter)]

            # Select rows with non-empty descriptions
            descNP = np[~np['Description'].isnull()]
            descAP = ap[~ap['Description'].isnull()]

            # Select rows with AMEX in description
            amexNP = descNP[descNP['Description'].str.contains('AMEX')].reset_index().drop(columns='index')
            amexAP = descAP[descAP['Description'].str.contains('AMEX')].reset_index().drop(columns='index')

            # Match dollar amounts from both files only with AMEX in description
            npAmount = list(amexNP['Amount'])
            apAmount = list(amexAP['Amount'])
            intAmount = [j for j in npAmount if j in apAmount]
            amex_amountmatch = amexAP[amexAP['Amount'].isin(intAmount)]


            # Reopen Accounting
            try:
                payable = driver.find_elements_by_xpath('//*[@class="ThemeOffice2003MainFolderText"]')
                payable[4].click()
            except:
                try:
                    LoginAccounting()
                except:
                    driver = webdriver.Chrome(driver)
                    LoginAccounting()
            ItemAmount = list()
            for L in range(0, len(amexNP)):
                # Reset search
                driver.find_element_by_xpath('//*[@type="button"]').click()

                # Search by amex invoice and submit
                driver.find_element_by_xpath("//select[@id='code_invoice_status']/option[text()='Not Paid (N)']").click()
                driver.find_element_by_xpath('//*[@id="invoice_number"]').send_keys(str(amexNP.loc[L][0]))
                driver.find_elements_by_xpath('//*[@type="submit"]')[0].click()

                # Retrieve all text
                allamount = driver.find_elements_by_xpath("//tr")
                soup = list()
                for j in allamount:
                    soup.append(j.text)
                longstring = ' '.join(soup)

                # Find all amounts in lines with letter L (find item amounts)
                findLs = re.findall('L .*?[0-9]+\.[0-9]+', longstring)
                longstring2 = ' '.join(findLs)

                numb = re.findall('[0-9]+?\.[0-9]+', longstring2)
                for K in range(0, len(numb), 1):
                    ItemAmount.append(numb[K])

                # Reset variables
                longstring = None
                findLs = None
                longstring2 = None
                
            # This list contains paid amex items  who's amount matches an amex item from the database
            MatchItem = amexAP[amexAP['Amount'].isin(ItemAmount)]
			
			# Save file with all dataframes onto desktop
            out_path = str(Path.home()) + "/desktop/Results.xlsx"
            writer = pd.ExcelWriter(out_path, engine='xlsxwriter')

            entered.to_excel(writer, sheet_name='Entered')
            errors.to_excel(writer, sheet_name='Errors')
            missing.to_excel(writer, sheet_name='Missing Check Dt')
            notentered.to_excel(writer, sheet_name='Not Entered')
            reqinv.to_excel(writer, sheet_name='Req. and Inv. Match')
            amex_amountmatch.to_excel(writer, sheet_name='AMEX_Amountmatch')
            MatchItem.to_excel(writer, sheet_name='Amex Item Match')
            writer.save()

# Run UI
if __name__ == "__main__":
    app = SampleApp()
    app.geometry("700x350")
    app.mainloop()

