import json
import openai
import os
import random
import re
import calendar
import time
from time import sleep
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, time, date
import csv
import argparse
from json.decoder import JSONDecodeError
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from datetime import datetime
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, PageBreak, Image, Spacer)
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.pagesizes import LETTER, inch
from reportlab.graphics.shapes import Line, LineShape, Drawing
from reportlab.lib.colors import Color


openai.api_key = "YourToken"

originFolderPath = r''

currentfolder = os.getcwd()
projectStart = datetime(2022, 1, 17, 10, 13, 1)
topics = ["Catering", "Supply", "Advisor", "Finances",
          "Attorney", "German Newspaper", "Moving", ""]

topicsIT = [
    "Remote work and virtual collaboration tools",
    "Cybersecurity best practices for employees",
    "IT infrastructure management and optimization",
    "Defining Project Scope and Objectives",
    "Creating a Project Plan and Timeline",
    "Resource Allocation and Task Assignment",
    "Managing Project Risks and Dependencies",
    "Vendor Selection and Management",
    "Managing Stakeholder Expectations",
    "Ensuring Effective Communication and Collaboration",
    "Monitoring and Controlling Project Progress",
    "Quality Assurance and Testing",
    "IT service management and ITIL framework",
    "IT governance and compliance",
    "Software testing and quality assurance",
    "DevOps and continuous integration/continuous deployment (CI/CD)",
    "IT career development and skills for the future",
    "IT asset management and software licensing"
]
topicsGeneral = [
    "Project planning",
    "Risk management",
    "Stakeholder management",
    "Resource allocation",
    "Project scheduling",
    "Cost estimation",
    "Scope management",
    "Quality assurance",
    "Change management",
    "Communication planning",
    "Team coordination",
    "Project documentation",
    "Project budgeting",
    "Issue tracking",
    "Decision-making",
    "Project monitoring and control",
    "Vendor management",
    "Project governance",
    "Agile project management",
    "Procurement management"
]
topicsManagement = [
    "Implementing Agile Methodologies (e.g., Scrum, Kanban)",
    "Optimizing Resource Allocation and Planning",
    "Identifying and Mitigating Project Risks",
    "Tracking Key Performance Indicators (KPIs) for Department Performance",
    "Effective Stakeholder Engagement and Management",
    "Establishing Processes for Project Documentation and Reporting",
    "Promoting Team Collaboration and Communication",
    "Implementing Quality Assurance and Control Measures",
    "Managing Change within Projects",
    "Fostering Continuous Improvement and Innovation"
]
topicsHR = [
    "Talent Acquisition and Management",
    "Employee Engagement and Development",
    # "HR Technology and Data Analytics",
    # "Sustainability and Resilience",
    # "Employee onboarding and orientation",
    # "Performance management and goal setting",
    # "Workplace diversity and inclusion",
    # "Employee engagement and motivation",
    # "Training and development opportunities",
    "Culture: Fostering a Culture of Innovation and Entrepreneurship",
    "Startup Talent Acquisition Strategies for ",
    "Equity and Incentive Compensation: Designing Reward Systems for Success",
    "Onboarding Pioneers: Creating Impactful Experiences for Project Newcomers",
    "Building High-Performance Teams in the Project Environment",
    "Motivating and Engaging Startup Employees within Project",
    "Diversity and Inclusion in the Project Ecosystem",
    "Learning and Development for Project Employees: Nurturing Startup Skills",
    "Succession Planning for Sustainable Project Growth",
    "Navigating Employment Law Compliance for Project Startups",
    "HR Tech Solutions for Efficient Project Operations",
    "Employee Relations and Conflict Management in the Project Context",
    "Well-being and Work-Life Integration in the Project Journey",
    "Attracting Top Talent to Project: Creating a Magnetic Culture",
    # "Workplace policies and procedures",
    # "Workplace safety and occupational health",
    # "Employee benefits and compensation",
    # "Leave management and time off policies",
    # "Conflict resolution and mediation",
    # "Career development and growth opportunities",
    # "Workplace communication and feedback",
    # "Workplace wellness programs",
    # "Recognition and rewards",
    # "Workplace flexibility and work-life balance",
    # "Employee relations and grievance handling",
    # "Ethics and compliance in the workplace",
    "Talent acquisition and recruitment",
    "Succession planning and talent management",
    "Employee offboarding and exit processes"
]
topicsPM = [
    "Workload and task management",
    "Setting and prioritizing goals",
    "Delegation and teamwork",
    "Time management and productivity",
    "Decision-making and problem-solving",
    "Effective communication skills",
    "Leadership and management styles",
    "Performance evaluation and feedback",
    "Managing conflicts and difficult situations",
    "Employee development and coaching",
    "Workplace collaboration and coordination",
    "Change management and adaptation",
    "Project planning and execution",
    "Resource allocation and optimization",
    "Budgeting and financial management",
    "Risk management and mitigation",
    "Quality control and process improvement",
    "Strategic planning and goal alignment",
    "Innovation and creativity in the workplace",
    "Stakeholder management and relationship building"
]
topicsPR = [
    "Media relations and press releases",
    "Social media management and engagement",
    "Brand management and messaging",
    "Crisis communication and reputation management",
    "Content creation and storytelling",
    "Event planning and coordination",
    "Influencer and blogger outreach",
    "Internal communications and employee engagement",
    "Public speaking and presentation skills",
    "Stakeholder and community relations",
    "Market research and audience analysis",
    "Measurement and analytics in PR",
    "Cross-functional collaboration with marketing and sales teams",
    "Strategic partnerships and collaborations",
    "Corporate social responsibility (CSR) initiatives"
]
problemsGeneral = [
    "Time management and productivity",
    # "Stress and burnout",
    # "Work-life balance",
    "Financial difficulties",
    "Health issues and wellness",
    "Relationship problems",
    # "Procrastination",
    # "Lack of motivation",
    # "Low self-esteem and self-confidence",
    "Decision-making challenges",
    # "Fear of failure",
    "Communication problems",
    "Conflict resolution",
    # "Loneliness and social isolation",
    # "Grief and loss",
    "Parenting challenges",
    # "Coping with change and uncertainty",
    # "Boundary setting and assertiveness",
    "Career dissatisfaction",
    "Lack of direction or purpose"
]
ITDepartment = [
    "Sebastian Weber",
    "Leon Becker",
    "Isabella Roth",
    "Benjamin Klein",
    "Emilia  Maier-Lorenz"
]
PRHRDepartment = [
    "Elena Martinez",
    "Johan Patel",
    "Mia Nakamura",
    "Liam Gonzalez",
    "Alessia Kim"]
PMDepartment = [
    "Luca Schmidt",
    "Rohan Mehta",
    "Benjamin Klein",
]
ProjectDepartment = [
    "Luca Schmidt",
    "Aisha Abadi",
    "Nico van der Berg",
    "Isha Kapoor",
    "Lucas Costa",
    "Eva Bjørnstad",
    "Rohan Mehta"]
correspondence = [
    "Emre O'Connor",
    "Lara Singh",
    "Antonio Silva",
    "Nikolai Petrov",
    "Sara Choudhury",
    "Niklas Petrov",
    "Amelie Richter",
    "David Krüger",
    "Amara Nguyen"]


class Folder:
    def __init__(self, path, name, creationTime, modTime, description, topic, department):
        self.path = path
        self.name = name
        self.creationTime = creationTime
        self.modTime = modTime
        self.description = description
        self.topic = topic
        self.department = department


class File:
    def __init__(self, path, title, text):
        self.path = path
        self.title = title
        self.text = text


def createFolder(folderName, desciption, date: datetime, topic: str, department: str):
    # folder creation------------------------------------------------------
    folderPath = originFolderPath + f"{folderName}"
    os.mkdir(folderPath)
    creationTime = datetime(date.year, random.randint(date.month, 12), random.randint(1, 28),
                            random.randint(8, 20), random.randint(0, 59), random.randint(0, 59))
    randomMonth= random.randint(1, 6)
    lastPossibleUpdate = datetime(2023, randomMonth,
                        random.randint(1, calendar.monthrange(2023, randomMonth)[1]), 10, 13, 1)
    modTime = random.randint(creationTime.timestamp(), lastPossibleUpdate.timestamp())
    modTime = datetime.fromtimestamp(modTime)
    modTime = datetime(modTime.year,modTime.month, modTime.day,random.randint(8, 20), random.randint(0, 59), random.randint(0, 59))
    currentFolder = Folder(folderPath, folderName,
                           creationTime, modTime, desciption, topic, department)
    # file creation---------------------------------------------------------
    createFile(random.randint(4, 11), currentFolder)  # 5-11, amount of files in folder
    subFolderNum = random.randint(1, 3)  # amount of subfolder in folder
    if subFolderNum > 0:
        createSubfolder(subFolderNum, currentFolder)
    os.utime(currentFolder.path, (creationTime.timestamp(), modTime.timestamp()))
    return modTime.timestamp()


def createSubfolder(numberSubFol, folder: Folder):
    randomCreation = folder.modTime.timestamp()-folder.creationTime.timestamp()
    message = 'Create ' + str(numberSubFol) + 'foldernames for the '+folder.department+' as one \
        json Object {"foldernames":[{"foldername":, "description":}] }, that go into specifics of desciption" ' + folder.description + \
        '", and two sentence desciption, give the response back in json format'
    res = chat(message)
    print(res)
    for i in res["foldernames"]:
        creationTime = folder.creationTime.timestamp() + \
            random.randrange(int(randomCreation))
        randomMod = folder.modTime.timestamp()-creationTime
        modTime = creationTime + random.randrange(int(randomMod))
        folderPath = folder.path + "//" + i["foldername"]
        os.mkdir(folderPath)
        subfolder = Folder(
            folderPath, i["foldername"], datetime.fromtimestamp(creationTime), datetime.fromtimestamp(modTime), i["description"], folder.topic, folder.department)
        createFile(random.randint(3, 5), subfolder)  # amount of files in folder
        os.utime(folderPath, (creationTime, modTime))


def createFile(numberFiles, folder: Folder):
    for i in range(numberFiles):
        items = ["PDF", "Email", "CSV", "XML", "TXT", "Readme"]
        probabilities = [0.3, 0.25, 0.1, 0.05, 0.2, 0.1]
        chosen_item = random.choices(items, probabilities)[0]
        sleep(30)
        switch_case(chosen_item, folder)


def switch_case(chosen_item, folder):
    if chosen_item == "PDF":
        process_pdf(folder)
    elif chosen_item == "Email":
        process_email(folder)
    elif chosen_item == "TXT":
        process_txt(folder)
    elif chosen_item == "XML":
        process_xml(folder)
    elif chosen_item == "Readme":
        process_readme(folder)
    elif chosen_item == "CSV":
        process_csv(folder)
    else:
        print("Invalid item.")


def process_csv(folder):
    items = ["people", "meetings", "tickets"]
    probabilities = [0.35, 0.35, 0.3]
    chosen_item = random.choices(items, probabilities)[0]
    if chosen_item == "people":
        peopleCSV(folder)
    elif chosen_item == "meetings":
        meetingCSV(folder)
    elif chosen_item == "tickets":
        ticketCSV(folder)
    else:
        print("Invalid item.")


def process_pdf(folder):
    pdfVariantOne(folder)


def process_email(folder):
    emailPDF(folder)


def process_txt(folder: Folder):
    randomString = random.choice(folder.topic)
    message = '''Create a python: {"txt":[{"Filename":, "Text":}]}.\
        Create a realistic filename with Underscores related to "''' + \
        folder.description + '''". Start with a text that is related to'''+randomString + \
        '''in '''+folder.department+'''  brought up by'''+random.choice(PRHRDepartment)+''' and addresses'''+random.choice(problemsGeneral)+'''. Talk about fake but realistic meetings and set goals. Create only one \
            text about descibed topic with more than 50 words with minor spelling mistakes. The text should be in common not refined speach. \
                Short to midlong sentences. Format the text into praragraphs.'''
    res = chat(message)
    try:
        filePath = folder.path + "//" + res["txt"][0]["Filename"].split(".")[0] + ".txt"
        with open(filePath, "w") as file:
            file.write(res["txt"][0]["Text"])
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        print(res)
        return

    setCreationDate(folder, filePath)


def process_xml(folder: Folder):
    message = '''"Create a json: {"xml":[{"Filename": , "rows": \
        [{"serviceName": "example", "MetaInfo": "example metadata"}]}]}. \
        Filename can be creative but related to context of the file.
        serviceName is a service in the field of Catering, Provisioning, Support, Entertainment, \
        Eventmanagement, Lawyer etc.. Create 10 services with creative names. MetaInfo contains one sentence describing \
        the service and a realistic German Person to contact for service in a realistic fake company.'''
    res = chat(message)
    print(res)
    try:
        root = ET.Element("Data")
        root = ET.Element("root")
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        print(res)
        return
 
    try:
        for item in res["xml"]:
            xml_element = ET.SubElement(root, "xml")
            for row in item["rows"]:
                row_element = ET.SubElement(xml_element, "row")
                ET.SubElement(row_element, "serviceName").text = row["serviceName"]
                ET.SubElement(row_element, "MetaInfo").text = row["MetaInfo"]

        tree = ET.ElementTree(root)
        tree.write(folder.path + "//" + res["xml"][0]["Filename"] + ".xml")
        setCreationDate(folder, folder.path + "//" +
                        res["xml"][0]["Filename"] + ".xml")
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        print(res)
        return

def generate_random_datetime():
    start_date = datetime(2022, 2, 1)
    random_days = random.randint(1, 365)
    random_date = start_date + timedelta(days=random_days)
    random_time = random.randint(0, 23), random.randint(
        0, 59), random.randint(0, 59)
    random_datetime = datetime(
        random_date.year, random_date.month, random_date.day, *random_time)
    return random_datetime.isoformat()


def process_readme(folder: Folder):
    message = '''Create one json Object: {"txt":[{"Filename":, "Text":}]}\
          create a creative Filename similar to the topic of:"''' + \
        folder.description + '''" . Create only one text about specific '''+ folder.department +''' '''+ random.choice(folder.topic)+''' \
            mention'''+random.choice(ProjectDepartment) + \
        '''and''' + random.choice(ProjectDepartment)+''' and how the team\
         wants to solve '''+random.choice(
        problemsGeneral) +'''in 200 words. Add some guidelines or realistic numbers \
        to support statements with minor spelling mistakes.'''
    res = chat(message)
    try:
        filePath = folder.path + "//" + \
            res["txt"][0]["Filename"].split(".")[0]+"README" + ".md"
        with open(filePath, "w") as file:
            file.write(res["txt"][0]["Text"])
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        return
    setCreationDate(folder, filePath)


def peopleCSV(folder: Folder):

    message = '''"Create one json data: \
        {"pdf":[{"Filename":, "people": [{"firstName": "Lukas","lastName": "Schmidt","phone": \
            "+49 1512 3456789","address": ", 12345 Berlin"}]}]} \
        . The Filename should have similar topic to ''' + folder.name + \
        '''. The json should contain 5 realistic German names, first and last name, realistic \
            address where they live and phonenumbers. Only use double quotes.'''
    res = chat(message)
    print("tryingCSV")
    try:
        peopledata = res["pdf"][0]["people"]
        csvPath = folder.path + "//" + res["pdf"][0]["Filename"] + ".csv"
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        return
    with open(csvPath, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["First Name", "Last Name",
                        "Email", "Phone", "Address"])
        for person in peopledata:
            try:
                writer.writerow(
                    [person["firstName"], person["lastName"], createEmail(person["firstName"] + person["lastName"]), person["phone"], person["address"]])
            except Exception as e:
                print(
                    "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
                return
    setCreationDate(folder, csvPath)


def meetingCSV(folder: Folder):
    message = ''' Create one json Object:\
        {"pdf":[{"Filename":, "people": [{"firstName": "Lukas","lastName": "Schmidt","phone": "+49 1512 3456789"}]}]} \
        The json should contain up to 5 realistic German names, realistic first and last name  \
        and phonenumbers. Filename should be contextual related to''' + \
        folder.name + '''. '''
    res = chat(message)
    print("tryingCSV")
    try:
        peopledata = res["pdf"][0]["people"]
        csvPath = folder.path + "//" + res["pdf"][0]["Filename"] + ".csv"
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        return
    # Write the data to the CSV file
    with open(csvPath, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["First Name", "Last Name",
                        "Phone", "Time", "MeetingRoom"])
        currentTime = datetime.now()
        for person in peopledata:
            randomDateTime = datetime.fromtimestamp(random.randint(
                int(projectStart.timestamp()), int(currentTime.timestamp())))
            randomRoomString = "Loc_" + random.choice("ABFGHINO")+"G_" + random.choice(
                "ABCDEFGHIJKLMNOPQRSTUV") + str(random.randint(0, 25))+"_" + str(random.randint(26, 798))
            try:
                writer.writerow(
                    [person["firstName"], person["lastName"], person["phone"], randomDateTime.date(), randomTimestamp(), randomRoomString])
            except Exception as e:
                print(
                    "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
                return
    setCreationDate(folder, csvPath)


def randomTimestamp():
    startTime = time(8, 0)
    endTime = time(17, 0)
    timeDiff = datetime.combine(date.today(
    ), endTime) - datetime.combine(date.today(), startTime)
    randomMinutes = random.choice(range(0, timeDiff.seconds // 60 + 1, 15))
    randomTimestamp = (datetime.combine(date.today(
    ), startTime) + timedelta(minutes=randomMinutes, seconds=0)).time()
    return randomTimestamp


def ticketCSV(folder: Folder):
    message = '''Create one json Object: {"pdf":[{"Filename":, "people": [{"firstName": "","lastName": "","phone": "+49"}]}]}. \
        The json should contain 10 to 15 realistic German names, first and last name, \
            and phonenumbers.The Filename should be related to "''' + folder.description + \
        '''.'''
    res = chat(message)
    print("tryingCSV")
    try:
        peopledata = res["pdf"][0]["people"]
        csvPath = folder.path + "//" + res["pdf"][0]["Filename"] + ".csv"
    except Exception as e:
        print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
        return
    # Write the data to the CSV file
    with open(csvPath, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["First Name", "Last Name",
                        "Phone", "Ticketnumber"])
        for person in peopledata:
            randomRoomString = "Loc_" + random.choice("ABFGHINO")+"G_" + random.choice(
                "ABCDEFGHIMNRTU") + str(random.randint(0, 8))+"_"+str(random.randint(26, 78))
            randomTicketnumber = randomTicket()
            try:
                writer.writerow(
                    [person["firstName"], person["lastName"], person["phone"], randomTicketnumber, randomRoomString])
            except Exception as e:
                print(
                    "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
                return
    setCreationDate(folder, csvPath)


def setCreationDate(folder, filenamePath):
    if folder.creationTime.timestamp() < folder.modTime.timestamp():
        creationTime = creatCreationTime(folder)
    else:
        creationTime = folder.creationTime.timestamp()
    if random.randint(0, 1) == 0 and folder.creationTime.timestamp() < folder.modTime.timestamp():
        modTime = createmodTime(folder, creationTime)
    else:
        modTime = creationTime
    os.utime(filenamePath, (creationTime, modTime))


def pdfVariantOne(folder: Folder):
    strings = ['Problems talking about budget and teams', 'Goals regarding the project and named stakeholders and organisations',
               'Guidelines for visiting sites, in company or talking to customers etc.',
               'Documentation about Project naming project teams and staff by name', 'Documentation about Hiring staff']

    random_string = random.choice(strings)

    stringnumbere = [" ", " numbered "]
    numbrand = random.choice(stringnumbere)
    message = '''Create a python parseable jsonFormat: \
        {"pdf":[{"Author":, "Filename":, "Title":,"Description":, "Content":["heading":,"text":]}]} \
            give the response back in json format.\
          Author should be a realistic German name. Realistic creative filename with Underscores. \
            Title similar to the Filename. Desciption is a short title for text in less than 6 words.Create the content regarding very specific ''' + \
        random_string + '''and '''+random.choice(problemsGeneral)+''' , with''' + numbrand + \
        '''headings and text about descibed topic in English with 1000. Be creative maybe add \
            realistic fake numbers and statistics to the context and the project '''+project+'''. Mention \
                some of the responsible from the Project Management Department people: ''' + \
        ','.join(PMDepartment)

    res = chat(message)
    if folder.creationTime.timestamp() < folder.modTime.timestamp():
        creationTime = creatCreationTime(folder)
    else:
        creationTime = folder.creationTime.timestamp()
    if random.randint(0, 1) == 0 and folder.creationTime.timestamp() < folder.modTime.timestamp():
        modTime = createmodTime(folder, creationTime)
    else:
        modTime = creationTime
    PDFPSReporte('psreport.pdf', res, folder,
                 int(creationTime), int(modTime), "PDF", "none", "none", "none")


def emailPDF(folder: Folder):
    strings = ['Problems talking about budget and teams', 'Goals regarding the project and named stakeholders and organisations',
               'Guidelines for visiting sites, in company or talking to customers etc.',
               'Documentation about Project naming project teams and staff by name', 'Documentation about Hiring staff']

    random_string = random.choice(strings)

    message = '''Create a python parseable jsonFormat: {"pdf":[{"Author":, "Filename":, \
        "Title":,"Description":, "Text":"Dear team..."}]} give the response \
        back in json format. Author should be a realistic German name. Realistic \
            creative filename with Underscores, related to the topic of"''' + \
        folder.description + '''" Title similar to the Filename. Create the Text regarding very specific ''' + \
        random_string + '''and '''+random.choice(problemsGeneral)+'''. Be very creative add some \
            realistic fake numbers regarding the topic. It should sound like a personal email. \
            Desciption is a short title for text in less than 6 words. Mention some of the responsible people from \
                the Project Management Department reagarding '''+project+'''. Additionally two people and how they can help: ''' + \
        ','.join(PMDepartment)
    res = chat(message)
    if folder.creationTime.timestamp() < folder.modTime.timestamp():
        creationTime = creatCreationTime(folder)
    else:
        creationTime = folder.creationTime.timestamp()
    if random.randint(0, 1) == 0 and folder.creationTime.timestamp() < folder.modTime.timestamp():
        modTime = createmodTime(folder, creationTime)
    else:
        modTime = creationTime
    toPerson = random.choice(correspondence)
    departmentEndings = ["IT", "PR", "Management", "HR"]
    departmentccPerson = random.choice(departmentEndings)
    fromPerson = random.choice(PRHRDepartment)
    ccPerson = random.choice(ITDepartment)+"-"+departmentccPerson
    PDFPSReporte('psreport.pdf', res, folder,
                          int(creationTime), int(modTime), "Email", toPerson, fromPerson, ccPerson)


def creatCreationTime(folder):
    randomCreation = folder.modTime.timestamp()-folder.creationTime.timestamp()
    creationTime = folder.creationTime.timestamp(
    ) + random.randrange(int(randomCreation))
    return creationTime


def createmodTime(folder, creationTime):
    randomMod = folder.modTime.timestamp()-creationTime
    modTime = creationTime + random.randrange(int(randomMod))
    return modTime


def randomTicket():
    randomTicketnumber = "T" + \
        str(random.randint(0, 7)) + "_" + str(random.randint(1, 3)) + \
        "_"+str(random.randint(34, 1500))
    return randomTicketnumber


def createEmail(person: str):
    person = person.replace(" ", "")

    platform = ["@web.de", "@hotmail.com",
                "@outlook.com", "@yahoo.com", "@gmx.com"]
    email = person + random.choice(platform)
    return email


class FooterCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []
        self.width, self.height = LETTER
        self.ticket = randomTicket()

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            if (self._pageNumber > 1):
                self.draw_canvas(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_canvas(self, page_count):
        page = "Page %s of %s" % (self._pageNumber, page_count)
        x = 128
        self.saveState()
        self.setStrokeColorRGB(0, 0, 0)
        self.setLineWidth(0.5)

        text = "Ticketnr."+self.ticket
        self.setFont('Helvetica-Bold', 8)

        self.drawString(30, 750, text)
        self.drawImage(currentfolder + '//logoClear.png', self.width - inch * 2, self.height-50,
                       width=100, height=30, preserveAspectRatio=True, mask='auto')
        self.line(30, 740, LETTER[0] - 50, 740)
        self.line(66, 78, LETTER[0] - 66, 78)
        self.setFont('Times-Roman', 10)
        self.drawString(LETTER[0]-x, 65, page)
        self.restoreState()


class PDFPSReporte():
    def __init__(self, path, res, folder: Folder, creationTime, modTime, type, toPerson, fromPerson, ccPerson):
        self.path = path
        self.res = res
        self.folder = folder
        self.creationTime = creationTime
        self.styleSheet = getSampleStyleSheet()
        self.elements = []
        self.modTime = modTime

        self.colorOhkaGreen0 = Color((45.0/255), (166.0/255), (153.0/255), 1)
        self.colorOhkaGreen1 = Color((182.0/255), (227.0/255), (166.0/255), 1)
        self.colorOhkaGreen2 = Color((140.0/255), (222.0/255), (192.0/255), 1)
        self.colorOhkaBlue0 = Color((54.0/255), (122.0/255), (179.0/255), 1)
        self.colorOhkaBlue1 = Color((122.0/255), (180.0/255), (225.0/255), 1)
        self.colorOhkaGreenLineas = Color(
            (50.0/255), (140.0/255), (140.0/255), 1)
        if type == "Email":
            self.emailPage(toPerson, fromPerson, ccPerson)
            try:
                filenamePath = folder.path + "//" + \
                    res["pdf"][0]["Filename"].split(".")[0] + ".pdf"
                self.doc = SimpleDocTemplate(filenamePath, pagesize=LETTER)
            except Exception as e:
                print(
                    "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
                print(res)
                return

            self.doc.multiBuild(self.elements, canvasmaker=FooterCanvas)
            os.utime(filenamePath, (self.creationTime, self.modTime))
        else:
            self.firstPage()
            self.nextPagesHeader(True)
            self.remoteSessionTableMaker()
            try:
                filenamePath = folder.path + "//" + \
                    res["pdf"][0]["Filename"].split(".")[0] + ".pdf"
                self.doc = SimpleDocTemplate(filenamePath, pagesize=LETTER)
            except Exception as e:
                print(
                    "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
                return

            self.doc.multiBuild(self.elements, canvasmaker=FooterCanvas)
            os.utime(filenamePath, (self.creationTime, self.modTime))

    def emailPage(self, toPerson, fromPerson, ccPerson):
        d = Drawing(500, 1)
        line = Line(-15, 0, 483, 0)
        customColor = Color(red=(0/255), green=(0/255), blue=(0/255))
        line.strokeColor = customColor
        line.strokeWidth = 2
        d.add(line)
        self.elements.append(d)

        # Pick a random string from the list
        psDetalle = ParagraphStyle(
            'Resumen', fontSize=10, leading=14, justifyBreaks=1, alignment=TA_LEFT, justifyLastLine=1)
        print(self.res)
        date = datetime.fromtimestamp(self.creationTime)

        psHeaderText = ParagraphStyle(
            'Hed0', fontSize=14, alignment=TA_LEFT, borderWidth=3, textColor=Color((0/255), (0/255), (0/255), 1))
        try:
            text = self.res["pdf"][0]["Title"]
            paragraphReportHeader = Paragraph(text, psHeaderText)
            self.elements.append(paragraphReportHeader)
            spacer = Spacer(10, 10)
            self.elements.append(spacer)

            text = str(date.strftime('%A')) + " "+str(date.date()) +\
                "<br/>to: " + toPerson+", " + createEmail(toPerson) +\
                "<br/>Cc: " + \
                ccPerson.split("-")[0] + ", " + \
                ccPerson.replace(" ", "")+"@growsimpler.com"

        except Exception as e:
            print(
                "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
            return

        paragraphReportSummary = Paragraph(text, psDetalle)
        self.elements.append(paragraphReportSummary)
        spacer = Spacer(10, 20)
        self.elements.append(spacer)
        styles = getSampleStyleSheet()
        text_style = styles["Normal"]
        introduction = ["Hello", "Dear", "Good day", "Greetings"]
        intro = random.choice(introduction)
        try:

            text = self.res["pdf"][0]["Text"].strip()
            sentences = re.split(r'(?<=[,.])\s*', text)
            res = ' '.join(sentences[1:-2])
            res = intro + " " +\
                re.split(r'\s+|\.', toPerson)[0] + ",<br/>" + res
            self.elements.append(Paragraph(res, text_style))
            self.elements.append(Spacer(1, 0.2 * inch))
        except Exception as e:
            print(
                "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
            print(res)
            return

        emailClosings = ['Kind regards', 'Best regards', 'Sincerely', 'Warm regards', 'Regards',
                         'Yours faithfully', 'Yours sincerely', 'Thank you', 'Thanks', 'Best wishes']
        closing = random.choice(emailClosings) + \
            ", <br/>" + fromPerson
        addressStyle = ParagraphStyle(
            'Resumen', fontSize=8, leading=14, justifyBreaks=1, alignment=TA_LEFT, justifyLastLine=1)
        address = "<br/>" + \
            fromPerson.replace(" ", "")+"-"+self.folder.department.split("-")[0]+  "@growsimpler.com" + \
            " <br/>"+self.folder.department+"<br/><br/>Hauptzentrale - GrowSimpler<br/>Mainzer \
                Landstraße 1893 Bahnhofsviertel<br/>60329 Frankfurt am Main"
        paragraphclosing = Paragraph(closing, psDetalle)
        paragraphAddress = Paragraph(address, addressStyle)

        self.elements.append(paragraphclosing)
        self.elements.append(paragraphAddress)

    def firstPage(self):
        img = Image(currentfolder + '//logoClear.png', kind='proportional')
        img.drawHeight = 0.6*inch
        img.drawWidth = 1.7*inch
        img.hAlign = 'LEFT'
        self.elements.append(img)
        d = Drawing(500, 1)
        line = Line(-15, 0, 483, 0)
        customColor = Color(red=(0/255), green=(2020.0/255), blue=(59/255))
        line.strokeColor = customColor
        line.strokeWidth = 2
        d.add(line)
        self.elements.append(d)

        spacer = Spacer(30, 100)
        self.elements.append(spacer)

        img = Image(currentfolder + '//logoClearCentere.png',
                    kind='proportional')
        img.drawHeight = 3.5*inch
        img.drawWidth = 3.5*inch

        self.elements.append(img)

        spacer = Spacer(10, 140)
        self.elements.append(spacer)
        string_list = ["Berlin", "Frankfurt", "Bremen",
                       "München", "Duisburg", "Leipzig"]

        # Pick a random string from the list
        randomCity = random.choice(string_list)
        psDetalle = ParagraphStyle(
            'Resumen', fontSize=10, leading=14, justifyBreaks=1, alignment=TA_LEFT, justifyLastLine=1)
        print(self.res)
        date = datetime.fromtimestamp(self.creationTime)
        psHeaderText = ParagraphStyle(
            'Hed0', fontSize=14, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaGreen0)
        try:
            text = self.res["pdf"][0]["Title"]

            paragraphReportHeader = Paragraph(text, psHeaderText)
            self.elements.append(paragraphReportHeader)
            spacer = Spacer(10, 3)
            self.elements.append(spacer)
            text = self.res["pdf"][0]["Author"] + "<br/>" + \
                randomCity + " - " + str(date.date())
        except Exception as e:
            print(
                "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
            return
        paragraphReportSummary = Paragraph(text, psDetalle)
        self.elements.append(paragraphReportSummary)
        self.elements.append(PageBreak())

    def nextPagesHeader(self, isSecondPage):
        if isSecondPage:
            psHeaderText = ParagraphStyle(
                'Hed0', fontSize=16, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaGreen0)
            try:
                text = self.res["pdf"][0]["Title"]
            except Exception as e:
                print("Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
                return
            paragraphReportHeader = Paragraph(text, psHeaderText)
            self.elements.append(paragraphReportHeader)
            spacer = Spacer(10, 10)
            self.elements.append(spacer)
            d = Drawing(500, 1)
            line = Line(-15, 0, 483, 0)
            line.strokeColor = self.colorOhkaGreenLineas
            line.strokeWidth = 2
            d.add(line)
            self.elements.append(d)

            spacer = Spacer(10, 1)
            self.elements.append(spacer)

            d = Drawing(500, 1)
            line = Line(-15, 0, 483, 0)
            line.strokeColor = self.colorOhkaGreenLineas
            line.strokeWidth = 0.5
            d.add(line)
            self.elements.append(d)

            spacer = Spacer(10, 22)
            self.elements.append(spacer)

    def remoteSessionTableMaker(self):
        psHeaderText = ParagraphStyle(
            'Hed0', fontSize=12, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaBlue0)
        try:
            text = self.res["pdf"][0]["Description"]
        except Exception as e:
            print(
                "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
            return
        paragraphReportHeader = Paragraph(text, psHeaderText)
        self.elements.append(paragraphReportHeader)

        spacer = Spacer(10, 22)
        self.elements.append(spacer)

        styles = getSampleStyleSheet()
        heading_style = styles["Heading1"]
        text_style = styles["Normal"]
        try:
            for section in self.res["pdf"][0]["Content"]:
                self.elements.append(
                    Paragraph(section["heading"], heading_style))
                self.elements.append(Paragraph(section["text"], text_style))
                self.elements.append(Spacer(1, 0.2 * inch))
        except Exception as e:
            print(
                "Chatgpt ditched the template, doing a retry.\n Exception: " + str(e))
            return


def checkForChatgptcreativity(res):
    start_index = res.find("// Add")
    end_index = res.find("\n", start_index)
    start_indexP = res.find("...")
    end_indexP = res.find("\n", start_index)

    if start_index != -1 and end_index != -1:
        modified_json = res[:start_index] + res[end_index:]
        return modified_json

    if start_indexP != -1 and end_indexP != -1:
        modified_json = res[:start_indexP] + res[end_indexP:]
        return modified_json
    return res


def chat(message):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            temperature=0.4,
            frequency_penalty=0.2,
            max_tokens=2000,
            messages=[
                {"role": "user", "content": f"{message}"},
            ]
        )
    except Exception as e:
        print("Exception:" + str(e))
        print("Timout occured! Waiting for rate limit")
        sleep(20)
        chat(message)
    try:
        checkedString = checkForChatgptcreativity(
            response['choices'][0]['message']['content'])
        response = json.loads(checkedString)
        return response
    except (JSONDecodeError, UnboundLocalError) as e:
        print("Waiting for rate limit")
        sleep(20)

        print("Rerunning request")
        chat(message)
        return


if __name__ == "__main__":
    topics = [topicsHR, topicsIT, topicsPM, topicsPR, topicsManagement, topicsGeneral]
    department = ["IT-Department", "HR-Department", "Management-Department", "PR-Department", "Projekt-Department", "GrowSimpler"]
    project = ["CleverCommute", "EcoPulse-Projects", "TerraTransit", "ProNetiva", "ConnectXcel"]
    parser = argparse.ArgumentParser(description='returns credentials')
    parser.add_argument('--mainfolder', type=str, help='path to the credentials file')
    parser.add_argument('--department',  choices=department, help='Select a department ["IT-Department", "HR-Department", "Management-Department", "PR-Department", "Projekt-Department"]')
    parser.add_argument('--project', '-p', type=int, choices=range(0, len(project)), help='["CleverCommute", "EcoPulse-Projects", "TerraTransit", "ProNetiva", "ConnectXcel"]')
    parser.add_argument('--topic', '-t', type=int, choices=range(0, len(topics)), help='Pick a number, starts with 0, [topicsHR, topicsIT, topicsPM, topicsPR, topicsManagement]')
    parser.add_argument('--number', '-n', type=int, help='number of topfolders')

    args = parser.parse_args()

    department = args.department
    topics = topics[args.topic]
    originFolderPath = args.mainfolder + "//"
    project = project[args.project]
    numberTopFolder = args.number
    print("Your choice:\n")
    print("Department: " + department)
    print("TopicList: " + str(topics))
    print("Folderpath: " + originFolderPath)
    print("Project: " + project)
    print("NumberFolders: " + str(numberTopFolder))

    promt = 'Create ' + str(numberTopFolder) + ' creative realistic "foldername" and four sentence desciption \
        {"foldernames":[{"foldername":, "description":}]},\
        the Project that the folder is related to is a project management company' + project + '. talk about \
        general topics regarding GrowSimpler. The description should contain what the project is about. Give the response back in json format'
    creationTime = datetime(2022, random.randint(2, 12), random.randint(1, 27),
                            random.randint(8, 20), random.randint(0, 59), random.randint(0, 59))
    res = chat(promt)
    modtimes = []
    for i in res["foldernames"]:
        print(i)
        modtime = createFolder(i["foldername"], i["description"], creationTime, topics, department)
        modtimes.append(modtime)
    modTime = max(modtimes)
    os.utime(originFolderPath, (creationTime.timestamp(), modTime))
