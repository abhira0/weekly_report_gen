import os
from copy import deepcopy

import docx
import yaml
from docx.document import Document as DOCument
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt
from docx.table import _Cell, _Column, _Row

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
        doc: DOCument = docx.Document("./template.docx")
        self.createTable(doc)
        doc.save("sample.docx")

    def setUltimatum(self):
        week_no = int(input("Week No: "))
        cwd = os.path.dirname(os.path.realpath(__file__))
        path_ = f"{cwd}\\weekConfigs"
        with open(f"{path_}\\{week_no}.yml", "r") as f:
            self.ultimatum = yaml.load(f, Loader=yaml.FullLoader)

    def createTable(self, doc: DOCument):
        table = doc.tables[0]
        C_ProjectID = doc.styles.add_style("C_ProjectID", WD_STYLE_TYPE.PARAGRAPH)
        C_ProjectID.font.size = Pt(15)
        C_ProjectID.font.bold = True
        # Project ID
        para: _Cell = table.cell(0, 3).paragraphs[0]
        para.style = doc.styles["C_ProjectID"]
        para.text = config_yml["project"]["id"]
        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        # Project Title
        para: _Cell = table.cell(1, 3).paragraphs[0]
        para.style = doc.styles["C_ProjectID"]
        para.text = config_yml["project"]["title"]
        table.add_row()


# Week()
Converter()
