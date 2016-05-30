#! /usr/bin/python
# -*- coding: UTF-8 -*-
#
# Copyright 2015 Balasankar C <balasankarc@autistici.org>
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# .
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# .
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import argparse
import json
import statistics
import sys
import textwrap
from argparse import RawTextHelpFormatter as rt
import os
import time

import requests
from lxml import etree
from pyPdf import PdfFileReader
from PySide.QtCore import *
from PySide.QtGui import *
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer

import pdftableextract as pdf


class GUI(QWidget):

    def __init__(self):
        super(GUI, self).__init__()
        self.grabber = None
        self.initSplash()
        self.initUI()

    def initSplash(self):
        splash_pix = QPixmap('myImage.jpg')
        splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
        splash.setMask(splash_pix.mask())
        splash.show()
        splash.showMessage("Loading", color=Qt.white)
        self.grabber = ResultGrabber()
        splash.finish(self)

    def initUI(self):

        self.setGeometry(0, 0, 300, 300)
        self.setWindowTitle('ResultGrabber')

        self.label1 = QLabel("ResultGrabber")
        font = QFont("FreeMono", 20)
        font.setBold(True)
        self.label1.setFont(font)
        self.label1.setMaximumHeight(50)
        self.label1.setAlignment(Qt.AlignCenter)

        self.cbox = QComboBox()
        self.cbox.addItem("Select Examination")
        for exam in self.grabber.exams:
            self.cbox.addItem(exam)

        self.cbox.currentIndexChanged.connect(self.comboClicked)

        self.csvButton = QPushButton("Generate CSV")
        self.csvButton.clicked.connect(self.csvClicked)
        self.pdfButton = QPushButton("Generate Summary PDF")
        self.quitButton = QPushButton("Quit")
        self.quitButton.clicked.connect(QCoreApplication.instance().quit)
        self.downloadButton = QPushButton("Download and Process")
        self.downloadButton.clicked.connect(self.downloadClicked)

        self.startText = QLineEdit()
        self.startText.setPlaceholderText("Starting Registration Number")
        self.endText = QLineEdit()
        self.endText.setPlaceholderText("Ending Registration Number")
        self.startText.textChanged.connect(self.textChanged)
        self.endText.textChanged.connect(self.textChanged)
        grid = QGridLayout()
        grid.addWidget(self.label1, 0, 0, 1, 3)
        grid.addWidget(self.cbox, 1, 0, 1, 3)
        grid.addWidget(self.startText, 2, 0, 1, 3)
        grid.addWidget(self.endText, 3, 0, 1, 3)
        grid.addWidget(self.downloadButton, 4, 0, 1, 3)
        grid.addWidget(self.csvButton, 5, 0)
        grid.addWidget(self.pdfButton, 5, 1)
        grid.addWidget(self.quitButton, 5, 2)
        self.setLayout(grid)
        self.label1.setFocus(Qt.OtherFocusReason)
        self.csvButton.setEnabled(False)
        self.pdfButton.setEnabled(False)
        self.downloadButton.setEnabled(False)
        self.startText.setEnabled(False)
        self.endText.setEnabled(False)
        self.show()

    def comboClicked(self):
        if self.cbox.currentText() == "Select Examination":
            self.csvButton.setEnabled(False)
            self.pdfButton.setEnabled(False)
            self.downloadButton.setEnabled(False)
            self.startText.setEnabled(False)
            self.endText.setEnabled(False)
        else:
            self.csvButton.setEnabled(False)
            self.pdfButton.setEnabled(False)
            self.downloadButton.setEnabled(False)
            self.startText.setEnabled(True)
            self.endText.setEnabled(True)

    def downloadClicked(self):
        start = self.startText.text()
        end = self.endText.text()
        code = self.grabber.exams[self.cbox.currentText()]
        print "Downloading"
        self.grabber.download(code, start, end)
        print "Completed Downloading"
        self.csvButton.setEnabled(True)
        self.pdfButton.setEnabled(True)

    def textChanged(self):
        if self.startText.text() != "" and self.endText.text() != "":
            if self.startText.text().isdigit() and \
                    self.endText.text().isdigit():
                        if len(self.startText.text()) == 8 and \
                                len(self.endText.text()) == 8:
                            flag = True
                        else:
                            flag = False
            else:
                flag = False
        else:
            flag = False
        if flag:
            self.downloadButton.setEnabled(True)
            self.csvButton.setEnabled(False)
            self.pdfButton.setEnabled(False)
        else:
            self.csvButton.setEnabled(False)
            self.pdfButton.setEnabled(False)
            self.downloadButton.setEnabled(False)

    def csvClicked(self):
        print "Button Clicked"

    def pdfClicked(self):
        print "Button Clicked"


class ResultGrabber(object):

    def __init__(self):
        self.result_subject = {}
        self.result_register = {}
        self.url = 'http://projects.mgu.ac.in/bTech/btechresult/index.php?\
                module=public&attrib=result&page=result'
        self.exams = {}
        self.getexamlist()

    def getexamlist(self):
        '''
        This method prints a list of exams whose results are available
        '''
        try:
            page = requests.get(self.url)
            pagecontenthtml = page.text
            tree = etree.HTML(pagecontenthtml)
            code = tree.xpath('//option/@value')
            text = tree.xpath('//option')
            for i in range(1, len(code)):
                self.exams[text[i].text.strip()] = str(code[i])
        except Exception as e:
            print e
            print "There are some issues with the connectivity.",
            print "May be due to heavy load. Please try again later"
            sys.exit(0)

    def download(self, examcode, start, end):
        '''
        Downloads the results of register numbers from 'start' to 'end'.
        '''
        start = int(start)
        end = int(end)
        verbosity = 1
        try:
            for count in range(start, end + 1):
                if verbosity == 1:
                    print "Roll Number #", count
                else:
                    sys.stdout.write(
                        "\r%.2f%%" % (float((count - start) * 100) /
                                      float(end - start)))
                    sys.stdout.flush()
                payload = dict(exam=examcode, prn=count, Submit2='Submit')
                r = requests.post(self.url, payload)
                if r.status_code == 200:
                    with open('result' + str(count) + '.pdf', 'wb') as \
                            resultfile:
                        for chunk in r.iter_content():
                            resultfile.write(chunk)
            print ""
        except Exception as e:
            print e
            print "There are some issues with the connectivity.",
            print "May be due to heavy load. Please try again later"
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

    def process(start, end):
        '''This method processes the specified results and populate necessary data
        structures.'''
        global result, exam
        subjects = []
        registers = []
        badresult = []
        for count in range(start, end + 1):
            try:
                if verbosity == 1:
                    print "Roll Number #", count
                else:
                    sys.stdout.write(
                        "\r%.2f%%" % (float(count - start) * 100 / (end - start)))
                    sys.stdout.flush()
                pages = ["1"]
                f = open("result" + str(count) + ".pdf", "rb")
                PdfFileReader(f)          # Checking if valid pdf file
                f.close()
                cells = [pdf.process_page("result" + str(count) + ".pdf", p)
                         for p in pages]
                cells = [item for sublist in cells for item in sublist]
                li = pdf.table_to_list(cells, pages)[1]
                for i in li:
                    if 'Branch' in i[0]:
                        collegepos = i[0].index('College : ')
                        branchpos = i[0].index('Branch : ')
                        namepos = i[0].index('Name : ')
                        registerpos = i[0].index('Register No : ')
                        exampos = i[0].index('Exam Name : ')
                        college = i[0][collegepos:branchpos][9:].strip().title()
                        branch = i[0][branchpos:namepos][9:].strip().title()
                        exam = i[0][exampos:][11:].strip().title()
                        register = i[0][registerpos:exampos][13:].strip()
                        name = i[0][namepos:registerpos][7:].strip()
                        if college not in result:
                            result[college] = {}
                        if branch not in result[college]:
                            result[college][branch] = {}
                    elif 'Mahatma' in i[0]:
                        pass
                    elif 'Sl. No' in i[0]:
                        pass
                    elif 'Semester Result' in i[1]:
                        pass
                    else:
                        subject = [i][0][1]
                        internal = i[2]
                        external = i[3]
                        if internal == '-':
                            internal = 0
                        else:
                            internal = int(internal)
                        if external == '-':
                            external = 0
                        else:
                            external = int(external)
                        res = i[5]
                        if register not in registers:
                            registers.append(register)
                        if subject not in subjects:
                            subjects.append(subject)
                        if register not in result[college][branch]:
                            result[college][branch][register] = {}
                        result[college][branch][register]["name"] = name
                        if subject not in result[college][branch][register]:
                            result[college][branch][register][subject] = {}
                        result[college][branch][register][subject] = \
                            [internal, external, internal + external, res]
            except Exception as e:
                print e
                badresult.append(count)
                continue
        if(len(badresult) > 0):
            print "\nUnavailable Results Skipped"
            for invalid in badresult:
                print "Roll Number #", invalid
        jsonout = json.dumps(result, indent=4)
        outfile = open('output.json', 'w')
        outfile.write(jsonout)
        outfile.close()
        print ""
        return subjects, registers


    def generatepdf():
        '''This method generates summary pdf from the results of result processor.
        '''
        global exam

        doc = SimpleDocTemplate("report.pdf", pagesize=A4,
                                rightMargin=72, leftMargin=72,
                                topMargin=50, bottomMargin=30)
        Story = []
        doc.title = "Exam Result Summary"
        # Defining different text styles to be used in PDF
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='Center1', alignment=1, fontSize=18))
        styles.add(ParagraphStyle(name='Center2', alignment=1, fontSize=13))
        styles.add(ParagraphStyle(name='Normal2', bulletIndent=20))
        styles.add(ParagraphStyle(name='Normal3', fontSize=12))
        for college in result1:
            for branch in result1[college]:
                # PDF Generation begins
                Story.append(Paragraph(college, styles["Center1"]))
                Story.append(Spacer(1, 0.25 * inch))
                Story.append(Paragraph(exam, styles["Center2"]))
                Story.append(Spacer(1, 12))
                numberofstudents = len(
                    result1[college][branch].itervalues().next())
                Story.append(Paragraph(branch, styles["Center2"]))
                Story.append(Spacer(1, 0.25 * inch))
                Story.append(Paragraph("Total Number of Students : %d" %
                                       numberofstudents, styles["Normal2"]))
                Story.append(Spacer(1, 12))
                for subject in result1[college][branch]:
                    marklist = [int(result1[college][branch][subject][x][0])
                                for x in result1[college][branch][subject]]
                    average = statistics.mean(marklist)  # Calculating mean
                    # Calculating standard deviation
                    stdev = statistics.pstdev(marklist)
                    passlist = {x for x in result1[college][branch][
                        subject] if 'P' in result1[college][branch][subject][x]}
                    faillist = {x for x in result1[college][branch][
                        subject] if 'F' in result1[college][branch][subject][x]}
                    absentlist = {x for x in result1[college][branch][
                        subject] if 'AB' in result1[college][branch][subject][x]}
                    passcount = len(passlist)
                    failcount = len(faillist)
                    absentcount = len(absentlist)
                    percentage = float(passcount) / numberofstudents
                    subjectname = "<b>%s</b>" % subject
                    passed = "<bullet>&bull;</bullet>Students Passed : %d" \
                        % passcount
                    failed = " <bullet>&bull;</bullet>Students Failed : %d" \
                        % failcount
                    absent = " <bullet>&bull;</bullet>Students Absent : %d" \
                        % absentcount
                    percentage = " <bullet>&bull;</bullet>Pass Percentage : %.2f"\
                        % percentage
                    average = " <bullet>&bull;</bullet>Average Marks : %.2f" \
                        % average
                    stdev = "<bullet>&bull;</bullet>Standard Deviation : %.2f" \
                        % stdev
                    Story.append(Paragraph(subjectname, styles["Normal"]))
                    Story.append(Spacer(1, 12))
                    Story.append(Paragraph(passed, styles["Normal2"]))
                    Story.append(Spacer(1, 12))
                    Story.append(Paragraph(failed, styles["Normal2"]))
                    Story.append(Spacer(1, 12))
                    Story.append(Paragraph(absent, styles["Normal2"]))
                    Story.append(Spacer(1, 12))
                    Story.append(Paragraph(percentage, styles["Normal2"]))
                    Story.append(Spacer(1, 12))
                    Story.append(Paragraph(average, styles["Normal2"]))
                    Story.append(Spacer(1, 12))
                    Story.append(Paragraph(stdev, styles["Normal2"]))
                    Story.append(Spacer(1, 12))
                Story.append(PageBreak())  # Each department on new page
        doc.build(Story)


    def getsummary():
        ''' This method generates overall average and standard deviation'''
        infile = open('output.json', 'r')
        filecontent = infile.read()
        infile.close()
        jsondata = json.loads(filecontent)
        result = {}
        final = {}
        for college in jsondata:
            for department in jsondata[college]:
                if department not in result:
                    result[department] = {}
                for subject in jsondata[college][department]:
                    if subject not in result[department]:
                        result[department][subject] = []
                    for student in jsondata[college][department][subject]:
                        result[department][subject].append(
                            jsondata[college][department][subject][student][0])
        for department in result:
            if department not in final:
                final[department] = {}
            for subject in result[department]:
                if subject not in final[department]:
                    final[department][subject] = []
                if len(result[department][subject]) > 1:
                    subjectaverage = statistics.mean(result[department][subject])
                    subjectdev = statistics.stdev(result[department][subject])
                    final[department][subject].append(subjectaverage)
                    final[department][subject].append(subjectdev)
                else:
                    final[department][subject].append(0)
                    final[department][subject].append(0)
        for department in final:
            print department
            for subject in final[department]:
                print "\t", subject
                print "\t\t Average : %.2f" % final[department][subject][0]
                print "\t\t Standard Deviation : %.2f " \
                    % final[department][subject][1]


    def tocsv(subjects, registers):
        f = open('output.json')
        content = f.read()
        result = json.loads(content)["Adi Shankara Institute Of Engineering And Technology"][
            "Computer Science And Engineering"]
        outputstring = ""
        for register in registers:
            details = result[register]
            outputstring += register
            name = details["name"]
            details.pop("name")
            outputstring += "," + name
            for subject in subjects:
                for mark in details[subject]:
                    outputstring += "," + str(mark)
            outputstring += "\n"
        csvout = open('output.csv', 'w')
        csvout.write(outputstring)
        csvout.close()

    if __name__ == '__main__':
        # Defining commandline options
        parser = argparse.ArgumentParser(
            description='Download and Generate Result Summaries',
            formatter_class=rt, add_help=False)
        parser.add_argument(
            "-h", "--help", help="\t\tShow this help and exit",
            action="store_true")
        parser.add_argument(
            "-l", "--list", help="\t\tList exam codes", action="store_true")
        parser.add_argument(
            "-d", "--download", help="\t\tDownload results", nargs=3,
            metavar=('START', 'END', 'EXAM'))
        parser.add_argument(
            "-p", "--process", help="\t\tDownload results", nargs=2,
            metavar=('START', 'END'))
        parser.add_argument(
            "-v", "--verbose", help="Enable Verbosity", action="store_true")
        args = parser.parse_args()

        verbosity = 0
        if args.help:
            print ""
            print "usage: getresult.py [-h] [-l LIST] [--download START END EXAM] \
                    [--process START END]"
            print "Download and Generate Result Summaries"
            print "\noptional arguments:"
            print "   -h/--help\t\t\t\tshow this help message and exit"
            print "   -v/--verbose\t\t\t\tEnable verbosity"
            print "   -l/--list\t\t\t\tList exam codes"
            print "   -d/--download START END EXAM\t\tDownload results"
            print "   -p/--process START END\t\tProcess results"
            sys.exit(0)
        if args.verbose:
            verbosity = 1
        if args.list:
            getexamlist(url)
        if args.download:
            start = int(args.download[0])
            end = int(args.download[1])
            exam = int(args.download[2])
            print "Downloading Results"
            download(url, exam, start, end)
        if args.process:
            start = int(args.process[0])
            end = int(args.process[1])
            result = {}
            subjects = []
            register = []
            print "Processing Results"
            subjects, registers = process(start, end)
            tocsv(subjects, registers)
            # generatepdf()
            # getsummary()


def main():
    app = QApplication(sys.argv)
    ex = GUI()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
