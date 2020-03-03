import os
import string
from shutil import copyfile
import pandas as pd
import xlrd
import numpy as np
import datetime
import openpyxl as pyxl
import tqdm
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from time import sleep
import matplotlib.pyplot as plt

class Student(object):
    
    def __init__(self,filename):
        self.filename = filename
        self.firstname = None
        self.lastname = None
        self.fullname = None
        self.email = None
        self.email_handle = None
        self.submission = None
        self.graded_submission = None
        self.score = None
        return

class TA(object): 
    
    def __init__(self,base,item,formulas=None):
        self.base = base
        self.item = item
        self.solutions = None
        self.solutions_q_start = None
        self.formulas = formulas
        self.grades = pd.DataFrame(columns=["name","score"])
        self.submissions_directory = os.path.join(base,item,"submissions")
        self.submissions_to_grade_directory = os.path.join(base,item,"submissions_to_grade")
        if not os.path.isdir(self.submissions_to_grade_directory):
            os.makedirs(self.submissions_to_grade_directory)
        self.grade_report_directory = os.path.join(base,item,"grade_reports")
        if not os.path.isdir(self.grade_report_directory):
            os.makedirs(self.grade_report_directory)
        self.log = open("/".join([base,item,"log.txt"]), "w+")

    def get_solution_details(self,solutions):  
        self.solutions_q_start = 20
        self.solutions_points = solutions.iloc[15,2]
        return
        
    def open_solutions(self,file):
        fname = "/".join([self.base,self.item,file])
        if self.formulas==True:
            self.solutions = pyxl.load_workbook(filename=fname)
            self.solutions = pd.DataFrame(self.solutions["solutions"].values)
        else:
            self.solutions = pd.read_excel(fname, sheet_name="solutions", header=None)
            self.solutions.columns = list(string.ascii_uppercase)[:self.solutions.shape[1]]
            self.solutions = self.solutions.fillna(np.NaN)
        self.get_solution_details(self.solutions)
        return
    
    def clean_submissions(self):
        for f in os.listdir(self.submissions_directory):
            if f.lower().endswith(".xlsx"):
                emailhandle = f.split("_")[1]
                fnew = "_".join([emailhandle, ".".join([self.item,"xlsx"])])
                fnew_location = os.path.join(self.submissions_to_grade_directory, fnew)
                copyfile(os.path.join(self.submissions_directory,f), fnew_location)
        return  
    
    def open_submission(self,student):
        fname = os.path.join(self.submissions_to_grade_directory,student.filename)
        try:
            if self.formulas is True:
                student.submission = pyxl.load_workbook(filename=fname)
                student.submission = pd.DataFrame(file["submission"].values)
            else:
                student.submission = pd.read_excel(fname, sheet_name="submission",header=None) 
        except xlrd.XLRDError as error:
            self.log.write("  Error when opening submission.\n")
            self.log.write("  Error: {} \n".format(str(error))) 
            pass              
        else:
            student.submission.columns = list(string.ascii_uppercase)[:student.submission.shape[1]]
            student.submission = student.submission.fillna(np.NaN)
            return
    
    def check_submission(self,student):
        if student.submission.shape != self.solutions.shape:
            student.submission_correct_format = False
        else:
            student.submission_correct_format = True
        return
    
    def complete_student_profile(self,student):
        try:
            student.firstname = student.submission.iloc[2,2]
            student.lastname = student.submission.iloc[3,2]
            student.email = student.submission.iloc[4,2]
            student.fullname = " ".join([student.firstname.upper(), student.lastname.upper()])
            student.email_handle = student.email.rsplit("@")[0].lower()
            student.ID = [student.fullname,student.email_handle,student.email]
        except AttributeError as error:
            self.log.write("  Error when completing student profile.\n")
            self.log.write("  Error: {} \n".format(str(error)))
            student.submission_correct_format = False
        return    

    def grade_submission(self,student):  
        sol = self.solutions.iloc[self.solutions_q_start:,]
        student.graded_submission = sol[sol!=student.submission.iloc[self.solutions_q_start:,]]
        student.score = self.solutions_points - student.graded_submission.count().sum()
        return
    
    def write_student_report(self,student):   
        filename = "_".join([student.email_handle, self.item, "gradereport.txt"])
        filename = os.path.join(self.grade_report_directory,filename)
        f = open(filename,"w+")
        f.write("---------------------------------------------------------\n")
        f.write("STATS-250\n")
        f.write("\nProblem set: {}".format(self.item.upper()))
        f.write("\nGrade report for {}".format(student.fullname))
        f.write("\nPlease send any questions to ldegeest@suffolk.edu\n")
        f.write("---------------------------------------------------------\n")
        f.write("\nTotal possible points: {}".format(self.solutions_points))
        f.write("\nYour score: {}".format(student.score))
        if student.score < self.solutions_points:
            f.write("\n{} cells marked incorrect:".format(self.solutions_points - student.score))
            for i in student.graded_submission.stack().index.values:
                cell = i[1]+str(i[0]+1)
                f.write("\n - {}".format(cell))
            f.close()
        else: 
            f.write("\nPerfect score. Great work!!")
            f.close()
        return
    
    def write_grade_report(self):
        filename = "_".join([self.item,"gradereport.txt"])
        filename = os.path.join(self.base,self.item,filename)
        f = open(filename,"w+")
        f.write("STATS-250")
        f.write("\nGrade report for {}\n\n".format(self.item.upper()))
        f.write("Total submissions: {}\n\n".format(len(os.listdir(self.submissions_to_grade_directory))))
        f.write("Summary of graded submissions:\n")
        grades_summary = self.grades.describe().round(decimals=2)
        f.write(" {}\n".format(grades_summary))
        f.write("\nGrading completed at: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        f.close()
        return
    
    def grade_all_submissions(self):
        self.log.write("Grading started at: {}\n".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        self.log.write("---------------------------------------------------------\n\n")
        n = len(os.listdir(self.submissions_to_grade_directory))
        with tqdm.tqdm_notebook(total=n) as progressbar:
            for submitted_file in tqdm.tqdm(os.listdir(self.submissions_to_grade_directory)):
                student = Student(submitted_file)
                self.log.write("Student: {}\n".format(student.filename))
                self.open_submission(student)
                if student.submission is None:
                    self.log.write("  Submission not graded.\n\n")
                    pass
                else:
                    self.check_submission(student)
                    self.complete_student_profile(student)
                    self.log.write("  Correct format: {}\n".format(student.submission_correct_format))
                    self.log.write("  Dimensions: {}\n".format(student.submission.shape))
                    if student.submission_correct_format is True:
                        self.log.write("  Grading {} for {}\n".format(self.item,student.fullname))
                        self.grade_submission(student)
                        self.log.write("  Score: {}\n\n".format(student.score))
                        self.write_student_report(student)
                        self.grades.loc[len(self.grades)] = [student.email_handle,student.score]
                    else:
                        self.log.write("  Information missing or format incorrect. Submission not graded.\n\n")
                    pass
                progressbar.update()
        self.grades["score"] = self.grades["score"].astype(np.float)
        self.write_grade_report()
        grades_filename = os.path.join(self.base,self.item,"grades.csv")
        self.grades.to_csv(grades_filename)
        self.log.write("\n---------------------------------------------------------\n")
        self.log.write("Grading completed at: {}".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        self.log.close()
        return
    
    def email_setup(self,email,password, student_email_extension):
        self.email_address = email
        self.password = password
        self.student_email_extension = student_email_extension
        return

    def send_grade_report(self,to_address,attachment):
        msg = MIMEMultipart()
        msg['From'] = self.email_address
        msg['To'] = to_address
        msg['Subject'] = " ".join(["Grade report for Excel", self.item.upper()]) 
        text1 = "Hi,"
        text2 = """Please find attached your grade report.\n\nLet me know if you have any questions!\n\nSend questions to ldegeest@suffolk.edu.\n\nLawrence"""    
        body = "\n\n".join([text1,text2])
        msg.attach(MIMEText(body, 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(attachment.name))
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.email_address, self.password)
        text = msg.as_string()
        server.sendmail(self.email_address, to_address, text)
        server.quit()
        return                    
        
    def send_all_grade_reports(self):
        n = len(os.listdir(self.grade_report_directory))
        with tqdm.tqdm_notebook(total=n) as progressbar:
            for grade_report in os.listdir(self.grade_report_directory):
                if grade_report.lower().endswith("txt"):
                    with open(os.path.join(self.grade_report_directory, grade_report), "r") as attachment:
                        to_address = "".join([grade_report.split("_")[0], self.student_email_extension])                   
                        self.send_grade_report(to_address,attachment)
                        sleep(1)
                else:
                    pass
                progressbar.update()
        return