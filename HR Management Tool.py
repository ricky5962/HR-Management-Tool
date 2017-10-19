from tkinter import filedialog
import tkinter
import os
import pandas as pd

# THe HR Management tool
# The tool has 2 layers of options
# The user can view and make changes to the applicant list, which is stored in a excel spread sheet
# The user can also move target applicants to the interview list and add a date for the interview


# Read the applicants and interviewees into data frames
try:
    applicants = pd.read_excel('/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Applicants.xlsx')
    interviewList = pd.read_excel('/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Interview List.xlsx')
except:
    applicantsColumn = ['ID','Last','First','Skill']
    interviewListColumn = ['ID','Last','First','Skill','Interview Date']
    applicants = pd.DataFrame(columns=applicantsColumn)
    interviewList = pd.DataFrame(columns=interviewListColumn)

# The HR management tool components


class App:
    # Constructor for HR management tool
    def __init__(self):
        self.main_menu = ['Main menu','View Applicants', 'Manage Applicants', 'View Interview List']
        self.view_applicants = ['View applicants', 'View all','Search by name' ,'Search by skill', 'Back to Main Menu']
        self.manage_applicants = ['Manage applicants','Add a new applicant', 'Delete an applicant',
                                  'Add to interview list', 'Back to Main Menu']
        self.current_menu = self.main_menu

    # toString method print the applicant list
    def toString(self):
        menu = '' + self.current_menu[0] + '\n'
        for i in range(1,len(self.current_menu)):
            menu = menu + str(i)+ '. ' + self.current_menu[i] +'\n'
        return menu

    # returns the current menu of the app
    def current(self):
        return self.current_menu

    # view applicants list
    def viewApplicants(self):
        return applicants.to_string(index=False)

    # view interview list
    def viewInterviewList(self):
        return interviewList.to_string(index=False)

    # search the applicant list by name
    def searchByName(self, last, first):
        return applicants.loc[(applicants['Last'] == last) & (applicants['First'] == first)]

    # search the applicant list by skill
    def searchBySkill(self, searchTerm):
        return applicants.loc[(searchTerm.lower() in applicants['Skill'].lower()  )]

    # add a new applicant to the list
    def addNewApplicant(self,id, last, first, skill):
        applicants.set_value(len(applicants),'ID',id)
        applicants.set_value(len(applicants)-1, 'Last', last)
        applicants.set_value(len(applicants)-1, 'First', first)
        applicants.set_value(len(applicants)-1, 'Skill', skill)
        return applicants

    # add the target to the interview list
    def addToInterviewList(self, new, date):
        interviewList.set_value(len(interviewList), 'ID', new['ID'].to_string(index=False))
        interviewList.set_value(len(interviewList) - 1, 'Last', new['Last'].to_string(index=False))
        interviewList.set_value(len(interviewList) - 1, 'First', new['First'].to_string(index=False))
        interviewList.set_value(len(interviewList) - 1, 'Skill', new['Skill'].to_string(index=False))
        interviewList.set_value(len(interviewList) - 1, 'Interview Date', date)
        return interviewList

    # delete target aplpicant from the list
    def deleteApplicant(self, last, first):
        return applicants[~((applicants['Last'] == last) & (applicants['First'] == first))]

    # switch menu
    def switch(self, menu_name):
        menu_name = menu_name.lower()
        if menu_name == self.main_menu[0].lower():
            self.current_menu = self.main_menu
        elif menu_name == self.view_applicants[0].lower():
            self.current_menu = self.view_applicants
        elif menu_name == self.manage_applicants[0].lower():
            self.current_menu = self.manage_applicants

    # update the changes in applicant data frame to the excel file
    def update(self, applicants):
        self.applicants = applicants
        applicants.to_excel('/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Applicants_temp.xlsx')
        os.remove(r'/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Applicants.xlsx')
        os.rename(r'/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Applicants_temp.xlsx',
                  r'/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Applicants.xlsx')
        return

    # update the changes in interview data frame to the excel file
    def update_interview(self, interview):
        self.interviewList = interview
        interviewList.to_excel('/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Interview List_temp.xlsx')
        os.remove(r'/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Interview List.xlsx')
        os.rename(r'/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Interview List_temp.xlsx',
                  r'/Users/xiao/Dropbox/Python Project/Python Class/PythonClassesExp/Interview List.xlsx')

    # a method that reads applicant's info from user selected file
    # method is not implemented
    def readFromFile(self):
        info = {'ID':'', 'Last':'', 'First':'','Skill':''}
        try:
            root = tkinter.Tk()
            file = filedialog.askopenfilename(parent=root, title='Choose a file')
            if file != None:
                info = open(file,'r')
                last = info.readline()
                first = info.readline()
                skill = info.readline()
        except:
            print('Please select a supported file type.')
        return info

    # print a format separator(+)
    def separater(self):
        print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')

# the sequence of the tool
print('Welcome to the HR management tool prototype!\nEnter \'Quit\' to Exit the app at anytime.')
app = App()
isRunning = 1
while isRunning == 1:
    print(app.toString())
    userInput = input('Please enter number or option to proceed:')
    app.separater()
    userInput = userInput.lower()
    current = app.current_menu[0].lower()

    if current == 'main menu':
        if userInput == '1' or userInput == 'View Applicants'.lower():
            app.switch('view applicants')
        elif userInput == '2' or userInput == 'Manage Applicants'.lower():
            app.switch('manage applicants')
        elif userInput == '3' or userInput == 'View Interview List'.lower():
            print(app.viewInterviewList())
        elif userInput == 'Quit'.lower():
            isRunning = 0
        else:
            print('Please enter a valid option.')
    elif current == 'view applicants':
        if userInput == '1' or userInput == 'view all'.lower():
            print(app.viewApplicants())
            app.separater()
        elif userInput == '2' or userInput == 'search by name'.lower():
            name = input('Please enter the first and last name separated by space:')
            app.separater()
            last = name.split(' ')[1]
            first = name.split(' ')[0]
            app.separater()
            print(app.searchByName(last, first))
            app.separater()
        elif userInput == '3' or userInput == 'search by skill'.lower():
            app.separater()
            skill = input('Please enter the skill you are looking for:')
            app.separater()
            print(app.searchBySkill(skill))
            app.separater()
        elif userInput == '4' or userInput == 'back to main menu'.lower():
            app.switch('main menu')
        elif userInput == 'Quit'.lower():
            isRunning = 0
        else:
            print('Please enter a valid option.')
    elif current == 'manage applicants':
        if userInput == '1' or userInput == 'add a new applicant'.lower():
            print('Please enter the information of the new applicant.(please enter skills separated by comma)')
            id = input('ID:')
            last = input('Last name:')
            first = input('First name:')
            skill = input('Skills:')
            app.update(app.addNewApplicant(id,last,first,skill))
            app.separater()
            print(app.viewApplicants())
        elif userInput == '2' or userInput == 'delete an applicant'.lower():
            print('Please enter the name of the applicant you want to remove.')
            last = input('Last name:')
            first = input('First name:')
            app.separater()
            app.update(app.deleteApplicant(last, first))
            print(app.applicants)
        elif userInput == '3' or userInput == 'add to interview list'.lower():
            print('Please enter the name of the applicant you want to move to the interview list.')
            last = input('Last name:')
            first = input('First name:')
            interviewTime = input('Please set a interview time for this applicant:')
            app.separater()
            app.addToInterviewList(app.searchByName(last, first),interviewTime)
            app.deleteApplicant(last, first)
            print(app.viewInterviewList())
            app.update(applicants)
            app.update_interview(interviewList)
        elif userInput == '4' or userInput == 'back to main menu'.lower():
            app.switch('main menu')
        elif userInput == 'Quit'.lower():
            isRunning = 0
        else:
            print('Please enter a valid option.')
else:
    app.separater()
    print('Thank you for using HR management tool.')
    app.separater()
