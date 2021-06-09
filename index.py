import yaml, os, docx
from docx.document import Document as DOCument
from docx.shared import Inches, Pt
from docx.table import _Column, _Row, _Cell
from copy import deepcopy
from docx.enum.style import WD_STYLE_TYPE

with open("config.yml", "r") as f:
    config_yml = yaml.load(f, Loader=yaml.FullLoader)

print(config_yml)


class YouTil:
    @staticmethod
    def cleanPathName(text):
        arr = ["\\", "/", "<", ">", "|", '"', "?", "*", ":"]
        for pattern in arr:
            text = text.replace(pattern, "")
        return text

    @staticmethod
    def makedir(path: str):
        if not os.path.exists(path):
            os.mkdir(path)


class Week:
    def __init__(self) -> None:
        self.setWeekDates()
        self.completed_activities = self.getActivities("Activities Completed")
        self.in_progress_activities = self.getActivities("In Progress")
        self.planned_activities = self.getActivities("Plan for Next Week")
        self.ultimatum = self.getUltimatum()
        self.dumpUltimatum()

    def setWeekDates(self):
        print("-" * 15 + " Dates " + "-" * 15)
        self.week_no = int(input("Week Number        : "))
        self.starting_date = input("Week Starting Date : ")
        self.ending_date = input("Week Ending Date   : ")

    def getActivities(self, activity_type):
        print("-" * 15 + f" {activity_type} " + "-" * 15)
        return [
            self.getTaskDetails(c + 1) for c in range(int(input("Number of tasks: ")))
        ]

    def getTaskDetails(self, count):
        task_description = input(f"{count}. Task Description: ")
        task_date = input(f"{count}. Date            : ")
        return {"task_description": task_description, "task_date": task_date}

    def getUltimatum(self):
        return {
            "week_no": self.week_no,
            "week_dates": {"start": self.starting_date, "end": self.ending_date},
            "completed_activities": self.completed_activities,
            "in_progress_activities": self.in_progress_activities,
            "planned_activities": self.planned_activities,
        }

    def dumpUltimatum(self):
        cwd = os.path.dirname(os.path.realpath(__file__))
        path_ = f"{cwd}\\weekConfigs"
        YouTil.makedir(path_)
        with open(f"{path_}\\{self.week_no}.yml", "w") as f:
            f.write(yaml.dump(self.ultimatum, sort_keys=False))


class Converter:
    def __init__(self) -> None:
        self.setUltimatum()
        self.createDocx()

    def createDocx(self):
        doc: DOCument = docx.Document()
        self.createTable(doc)
        doc.save("sample.docx")

    def setUltimatum(self):
        week_no = int(input("Week No: "))
        cwd = os.path.dirname(os.path.realpath(__file__))
        path_ = f"{cwd}\\weekConfigs"
        with open(f"{path_}\\{week_no}.yml", "r") as f:
            self.ultimatum = yaml.load(f, Loader=yaml.FullLoader)

    def createTable(self, doc: DOCument):
        no_com_act = len(self.ultimatum["completed_activities"])
        no_ip_act = len(self.ultimatum["in_progress_activities"])
        no_pla_act = len(self.ultimatum["planned_activities"])
        no_rows = 10 + no_com_act + no_ip_act + no_pla_act
        table = doc.add_table(rows=no_rows, cols=4, style=doc.styles["Table Grid"])
        #  First row --------------
        second_col: _Cell = table.rows[0].cells[1]
        for col in table.rows[0].cells[2:]:
            second_col.merge(col)
        #  Second row --------------
        second_col: _Cell = table.rows[1].cells[1]
        for col in table.rows[1].cells[2:]:
            second_col.merge(col)
        #  Activities Completed row --------------
        first_col: _Cell = table.rows[5].cells[0]
        for col in table.rows[5].cells[1:]:
            first_col.merge(col)
        #  In Progress row --------------
        in_progress_row = 6 + no_com_act
        first_col: _Cell = table.rows[in_progress_row].cells[0]
        for col in table.rows[in_progress_row].cells[1:]:
            first_col.merge(col)
        #  Plan for Next Week row --------------
        planned_row = in_progress_row + 1 + no_ip_act
        first_col: _Cell = table.rows[planned_row].cells[0]
        for col in table.rows[planned_row].cells[1:]:
            first_col.merge(col)
            table.autofit = False
        table.rows[-1].cells[0].width = Inches(3)
        table.rows[0].cells[0].width = Inches(0.5)


# Week()
# Converter()
"""
1
11/01/2021
17/01/2021
1
understanding fundamental concepts of python
12/01/2021
1
change in problem statement
14/01/2021
1
understanding of concepts
18/02/2021

"""

doc: DOCument = docx.Document("./template.docx")
for i in doc.tables:
    print(dir(i))
table = doc.tables[0]
normal_style = doc.styles.add_style("C_ProjectID", WD_STYLE_TYPE.PARAGRAPH)
normal_style.font.size = Pt(15)
normal_style.font.bold = True
para: _Cell = table.cell(0, 3).paragraphs[0]
para.style = doc.styles["C_ProjectID"]
para.text = config_yml["project"]["id"]
table.add_row()
doc.save("simple.docx")