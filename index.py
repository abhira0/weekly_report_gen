import yaml, os

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
        self.getUltimatum()

    def getUltimatum(self):
        week_no = int(input("Week No: "))
        cwd = os.path.dirname(os.path.realpath(__file__))
        path_ = f"{cwd}\\weekConfigs"
        with open(f"{path_}\\{week_no}.yml", "r") as f:
            self.ultimatum = yaml.load(f, Loader=yaml.FullLoader)
        print(self.ultimatum)


Week()
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