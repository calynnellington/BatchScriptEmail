#-------------------------------------------------------------------------------
# author: Calynn Ellington
# Documentation for win32com: http://docs.activestate.com/activepython/2.4/pywin32/com.html
#-------------------------------------------------------------------------------

# Necessary Modules.------------------------------------------------------------
import win32com.client as win32
import re
from tkinter import *
from tkinter.filedialog import askopenfilename
import tkinter.messagebox

# Class for selecting the file.-------------------------------------------------
class FilenameClass():
    def __init__(self):
        self.location = 'User Import.txt'

    def getFile(self, identity):
        self.file_opt = options = {}
        options['defaultextension'] = '.txt'
        options['filetypes'] = [('Text Document (.txt)', '.txt'),
                                ('all files', '.*')]
        self.filename = askopenfilename(**self.file_opt)
        if self.filename:
            if 'Assurx User Import' in identity:
                self.location = self.filename
                app.get_txt_File['bg'] = '#0d0'
                user_file = open(self.filename, 'r')
                user_total = user_file.read()
                remove_lines = user_total.splitlines()
                for user in remove_lines:
                    regex_tab = re.compile('\\t')
                    user_info = regex_tab.split(user)
                    app.users.append(user_info)
            else:
                app.loadButton['bg'] = '#e10'


# Main Class.-------------------------------------------------------------------
class Application(Frame, Tk):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.users = []
        self.fileOBJtxt = FilenameClass()
        self.createWidgets()

    def createWidgets(self):
        '''Builds the interface of the program.'''

        # Define the default values for the options for the buttons
        # Grid layout options
        self.rowconfigure(0, minsize=5)
        self.width = 54
        self.grid(padx=5)
        self.columnconfigure
        self.rowconfigure
        self.OfficialDisclaimer = """*REDACTED FOR PORTFOLIO SHOWCASE.*"""
        self.loadButton_gopt = {'row':1,'column':0,'padx': 2, 'pady': 5}
        self.loadButton_wopt = {'width': round(self.width),'bg':'#e10'}
        self.send_gopt = {'row':40,'column':0,'padx': 2, 'pady': 5}
        self.send_wopt = {'width': round(self.width/4)}
        self.loadButton()
        self.trainingCheckBox()
        self.italianCheckBox()
        self.signatureInput()
        self.send()

    def loadButton(self):
        '''Button that calls the filename class which allows the user to select
        the text file they wish to use.'''

        self.get_txt_File = Button(self, text="Load User List", \
        command=lambda: self.fileOBJtxt.getFile('User Import'))
        for key, value in self.loadButton_wopt.items():
            self.get_txt_File[key] = value
        self.get_txt_File.grid(**self.loadButton_gopt)

    def trainingCheckBox(self):

        self.training_var = IntVar()
        self.training = Checkbutton(self, text="Include training video?", \
        variable=self.training_var)
        self.training.grid(row=2, sticky=W, padx=1, pady=1)

    def italianCheckBox(self):
        self.italian_var = IntVar()
        self.italian = Checkbutton(self, text="Italian Users?", \
        variable=self.italian_var)
        self.italian.grid(row=3, sticky=W, padx=1, pady=1)

    def signatureInput(self):
        Label(self, text="Signature Name:").grid(row=4, sticky=W, pady=1, padx=1)
        self.entry = Entry(self, bg='#fff', width=30)
        self.entry.grid(row=4, column=0)

    def send(self):
        self.send_emails = Button(self, text="Send", command=lambda: \
        self.emailScript(self.entry.get(), self.training_var.get(), self.italian_var.get()))
        for key, value in self.send_wopt.items():
            self.send_emails[key] = value
        self.send_emails.grid(**self.send_gopt)

    def emailScript(self, signature, training, italian):
        if signature:
            check1 = tkinter.messagebox.askyesno('Verify information.', 'Are all options correct?')
            if check1 == True:
                # Interfaces with outlook, outlook does not have to be open in order for this to work.
                outlook = win32.Dispatch('outlook.application')
                if italian == 1:
                    print('italian')
                    if training == 1:
                        print('training 1')
                        # For each user send the following email.
                        for user in self.users:
                            mail = outlook.CreateItem(0)
                            mail.To = user[2]
                            mail.Subject = '*REDACTED FOR PORTFOLIO SHOWCASE.*'
                            mail.Body = '''*REDACTED FOR PORTFOLIO SHOWCASE.*'''.format(user[1].split(' ', 1)[0], user[0], signature)
                            mail.Send()
                    elif training == 0:
                        print('training 0')
                        # For each user send the following email.
                        for user in self.users:
                            mail = outlook.CreateItem(0)
                            mail.To = user[2]
                            mail.Subject = '*REDACTED FOR PORTFOLIO SHOWCASE.*'
                            mail.Body = '''*REDACTED FOR PORTFOLIO SHOWCASE.*'''.format(user[1].split(' ', 1)[0], user[0], signature)
                            mail.Send()
                elif italian == 0:
                    print('not italian')
                    if training == 1:
                        print('training 1')
                        # For each user send the following email.
                        for user in self.users:
                            mail = outlook.CreateItem(0)
                            mail.To = user[2]
                            mail.Subject = '*REDACTED FOR PORTFOLIO SHOWCASE.*'
                            mail.Body = '''*REDACTED FOR PORTFOLIO SHOWCASE.*'''.format(user[1].split(' ', 1)[0], user[0], signature)
                            mail.Send()
                    elif training == 0:
                        print('training 0')
                        # For each user send the following email.
                        for user in self.users:
                            mail = outlook.CreateItem(0)
                            mail.To = user[2]
                            mail.Subject = '*REDACTED FOR PORTFOLIO SHOWCASE.*'
                            mail.Body = '''*REDACTED FOR PORTFOLIO SHOWCASE.*'''.format(user[1].split(' ', 1)[0], user[0], signature)
                            mail.Send()
            else:
                pass
            tkinter.messagebox.showinfo('Script Finished.', 'To verify that the emails were succesfully sent, please check your sent items in Outlook')
        else:
            tkinter.messagebox.showinfo('Missing Parameter', 'Please enter a signature name.')
            pass

# Initialization parameters.----------------------------------------------------
if __name__ == '__main__':
    app = Application()
    app.master.title('New User Notification Tool')
    app.master.geometry('405x150+100+100')
    app.master.resizable(width=False, height=False)
    app.mainloop()
