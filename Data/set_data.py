## Built-in imports go here
import csv, logging, os

import tkinter as tk
import tkinter.messagebox as tkmb

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Image, PageBreak, Paragraph

## imports from files go here
## End of Imports

class Delegate:

    def make_dir(self, action=False):
        desktop = os.path.expanduser("~/Desktop")

        dir = "NYMUN FORMS"
        directory = os.path.join(desktop, dir)
        dirs = ["Individual", "3", "4", "5", "6"]

        logging.debug("Action => {}".format(action))
        if (action == False):
            if os.path.isdir(directory) != True:
                os.chdir(desktop)
                logging.info("Changed working directory to => '{}' and creating a new directory => '{}'".format(os.getcwd(), dir))
                os.mkdir(dir)
                os.chdir(dir)
                for dir in dirs:
                    os.mkdir(dir)
                return True

            else:
                logging.info("Changed working directory to => '{}' and directory => '{}' already exists!".format(os.getcwd(), dir))
                return False

        elif (action == True):
            os.chdir(directory)
            logging.info("Current working directory => '{}'".format(os.getcwd()))

            # creating csv file
            # with open("NYMUN Forms.csv", a) as



            files_to_create, files_created, errors = len(self.info), 0, 0
            logging.info("No of files to create => {}".format(files_to_create))

            for serial_no in self.info:
                try:
                    detail = self.info[serial_no]
                    os.chdir(detail[2])
                    logging.info("Current working directory => '{}'".format(os.getcwd()))

                    #Beginning of pdf
                    ## Creating a serial_no for pdf file
                    file  = "{}.pdf".format(str(serial_no))
                    logging.info("Creating a new file => '{}'".format(file))

                    c = canvas.Canvas(file)
                    c.setAuthor(("Team (NYMUN {} 2018)".format(chr(0xa9))))
                    c.setTitle("Delegate Information (NYMUN {} 2018)".format((chr(0xa9))))
                    c.setSubject("Complete Delegate Information")
                    c.setFont("Helvetica", 18)
                    c.drawString(70, 800,"REGISTRATION INFO OF TS :")
                    # Image("logo.png")
                    c.drawInlineImage((os.path.join(self.original_path, "Data\\logo.png")), 5, 795, 40, 40)
                    c.setFont("Helvetica", 16)
                    c.drawString(550, 800, "1 / 2")
                    c.setFont("Helvetica", 18)

                    x, y = 15, 766
                    name_list, experience_list = [], []
                    for index in range(len(detail)):
                        rem_country, rem_name, rem_experience = False, False, False
                        if (index == 0):
                            c.drawString(330, 800, detail[index][0: -5])
                            c.line(0, 792, 600, 792)
                            c.setFont("Helvetica-Bold", 12)

                        elif (index == 2):
                            continue

                        else:
                            opt = str(self.info_pattern[detail[2]][index]) + " :"

                            # a line space for photo url
                            if opt.startswith("Photo"):
                                c.drawString(x, y, opt)
                                y -= 22
                                c.drawString(x, y, "")

                            # Gender will on the same line as name
                            elif opt.startswith("Gender"):
                                y += 22
                                c.drawString(380, y, opt)
                                pass

                            # Country preference will be on the same line as Committee
                            elif opt.startswith("Country"):
                                rem_country = True
                                y += 22
                                c.drawString(322, y, opt)
                                pass

                            # PREVIOUS EXPERIENCE will be on the next page
                            elif opt.startswith("Previous"):
                                rem_experience = True
                                pass

                            elif opt.startswith("Committee"):
                                c.drawString(x, y, opt)

                            elif "Name" in opt and "Private" not in opt:
                                rem_name = True
                                c.drawString(x, y, opt)

                            else:
                                c.drawString(x, y, opt)
                            c.setFont("Helvetica", 12)

                            Gender = ["Male", "Female", "Other"]
                            if detail[index].startswith("https"):
                                c.drawString(x, y, detail[index])

                            elif detail[index] in Gender:
                                c.drawString(450, y, detail[index])

                            elif rem_country == True:
                                c.drawString(450, y, detail[index])
                                c.line(x, y-8.5, 575, y-8.5)
                                y += 10

                            elif rem_name == True:
                                c.drawString(240, y, detail[index])
                                name_list.append(detail[index])

                            elif rem_experience == True:
                                experience_list.append(detail[index])

                            else:
                                c.drawString(240, y, detail[index])
                            c.setLineWidth(0.2)
                            c.setFont("Helvetica-Bold", 12)
                            y -= 22

                    #Ending of first page
                    c.setFont("Courier-BoldOblique", 10)
                    c.line(0, 15, 600, 15)
                    c.drawString(x, 3, "NYMUN FORMS {} 2018".format(chr(0xa9)))
                    c.drawString(350, 3, "Credits: Faizan Ahmad(Asst.Director IT)")
                    c.showPage()

                    # start of second page
                    p = PageBreak()
                    p.drawOn(c, 0, 1000)
                    c.drawInlineImage((os.path.join(self.original_path, "Data\\logo.png")), 5, 795, 40, 40)
                    c.setFont("Helvetica", 18)
                    c.drawString(80, 800,"DELEGATE/DELEGATION PREVIOUS EXPERIENCE")
                    c.setFont("Helvetica", 16)
                    c.drawString(550, 800, "2 / 2")
                    c.line(0, 792, 600, 792)
                    x, y = 15, 766
                    style = ParagraphStyle(name='Normal', fontName='Helvetica-Oblique', fontSize=12)
                    for index in range(len(name_list)):
                        c.setFont("Helvetica-BoldOblique", 16)
                        c.drawString(x, y, "{}'s Previous Experience :".format(name_list[index]))
                        y -= 55
                        c.setFont("Helvetica-Oblique", 12)
                        if experience_list[index] != "":
                            para = Paragraph(experience_list[index], style)
                            para.wrapOn(c, 450, 108)
                            para.drawOn(c, x+20, y)

                        else:
                            c.drawString(x+20, y, "NO INFO. PROVIDED")
                        y -= 73
                        c.line(x, y+16.5, 575, y+16.5)

                    #Ending of second page
                    c.setFont("Courier-BoldOblique", 10)
                    c.line(0, 15, 600, 15)
                    c.drawString(x, 3, "NYMUN FORMS {} 2018".format(chr(0xa9)))
                    c.drawString(350, 3, "Credits: Faizan Ahmad(Asst.Director IT)")
                    c.showPage()
                    c.save()
                    logging.info("File => '{}' was created successfully with no errors reported.".format(file))
                    os.chdir("..")
                    files_created += 1

                except Exception as exp:
                    logging.warn("Exception raised => {0} , while creating file => '{1}.pdf'".format(exp, serial_no))
                    errors += 1
                    os.chdir("..")
                    pass

            logging.info("No of files created => {}".format(files_created))
            self.files_created = files_created
            self.dir = dir
            self.errors = errors
            return True

        else:
            return False

    def get_data(self):
        info, info_pattern = {}, {}
        serial_count =  0
        with open(self.file, 'r') as infile:
            logging.info("Opening file => '{}' in read-mode".format(self.file))
            data = csv.reader(infile)
            logging.debug("File => '{}' opened successfully, {}".format(self.file, data))
            for line in data:
                if serial_count < 1:
                    info_pattern["Individual"] = tuple(line[0:3] + line[3:11])
                    info_pattern["3"] = tuple(line[0:3] + line[12:34])
                    info_pattern["4"] = tuple(line[0:3] + line[35:64])
                    info_pattern["5"] = tuple(line[0:3] + line[65:101])
                    info_pattern["6"] = tuple(line[0:3] + line[102:145])

                elif serial_count >= 1:
                    len_serial = len(str(serial_count))
                    serial_no = "S#"+"0"*(4 - len_serial)+str(serial_count)
                    if line[2] == "Individual":
                        info[serial_no] = tuple(line[0:3] + line[3:11])
                    elif line[2] == "3":
                        info[serial_no] = tuple(line[0:3] + line[12:34])
                    elif line[2] == "4":
                        info[serial_no] = tuple(line[0:3] + line[35:64])
                    elif line[2] == "5":
                        info[serial_no] = tuple(line[0:3] + line[65:101])
                    elif line[2] == "6":
                        info[serial_no] = tuple(line[0:3] + line[102:145])

                serial_count += 1

        self.info_pattern = info_pattern
        self.info = info
        return True

    def __init__(self, file):
        original_path = os.getcwd()
        self.original_path = original_path
        self.file = file
        self.get_data()
        self.make_dir()
        self.make_dir(action=True)
        #     logging.info("Process was completed successfully with errors reported => '{}'.".format(self.errors))
        #     # tkmb.showinfo("Successfull", "{0} Files written in directory -> {1}, Go To Your Desktop And Look For '{1}' Directory".format(self.files_created, self.dir))
        # else:
        #     logging.warn("Process could not complete, error were encountered!")
        #     # tkmb.showerror("Error", "Contact Your Developer!")
