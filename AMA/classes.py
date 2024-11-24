import os
import json
from datetime import date, timedelta
from docx import Document
from docx.shared import Inches
from pathlib import Path
from docx.shared import Pt

class Internship():
    def __init__(self, week_no:str, required_hrs:str, completed_hrs:str, start_date:str, current_week_start:str):
        self.week_no = int(week_no)
        self.required_hrs = int(required_hrs)
        self.completed_hrs = int(completed_hrs)
        self.start_date = start_date
        self.current_week_start = current_week_start

    def date_to_str(date:date) -> str:
        return date.strftime("%B %d, %Y")
    
    def get_week_dates(self, week_start:str)->list[date]:
        """Returns a list of 5 consecutive days starting from a given date YYYY/MM/DD"""
        temp_list = []
        week_start_date = Internship.str_to_date(week_start)
        for i in range(5):
            temp_list.append(week_start_date)
            week_start_date += timedelta(days=1)
        
        week_dates = []
        for day in temp_list:
            week_dates.append(Internship.date_to_str(day))
        return week_dates
    
    def get_remaining_hrs(self) -> int:
        return self.required_hrs - self.completed_hrs
    
    def account_absent(self, hours:int = 8) -> None:
        self.completed_hrs - hours
        print('Deducted ' + str(hours) + 
              ' hours from completed hours.')
    
    def account_diff_hours(self, hours:int) -> None:
        self.completed_hrs - hours
        if hours > 0:
            print("Added " + str(hours) + 
                  " hours to completed hours")
        else:
            print('Deducted ' + str(hours) + 
              ' hours from completed hours.')
    
    def increment_week_no(self, num:int = 1) -> None:
        self.week_no += num
        print('Week no. incremented by ' + str(num))
    
    def str_to_date(datestring:str) -> date:
        temp = str(datestring).split('-')
        return date(int(temp[0]), int(temp[1]), int(temp[2]))

class Report(Internship):
    def __init__(self, week_no:str, required_hrs:str, completed_hrs:str, start_date:str, current_week_start, template_file, out_file, save_file):
        super().__init__(week_no, required_hrs, completed_hrs, start_date, current_week_start)
        self.template_file = Path(template_file)
        self.out_file = Path(out_file)
        self.save_file = Path(save_file)

    
    def save(self) -> None:
        """Saves the current report object as a json in save location."""
        attributes = [a for a in dir(self) if not a.startswith('__')]

        save_data = {}

        for attribute in attributes:
            save_data[attribute] = getattr(self, attribute)
        print('Save data generated.')
        
        with open(self.save_file) as file:
            json.dump(save_data, file, indent=4)
        print('Report data saved to ' + self.save_file + ".")

    def turn_bold(paragraph):
        for run in paragraph.runs:
            run.bold = True
    
    def edit_document(self):
        doc = Document(self.template_file)
        data = self.generate_data(str(self.current_week_start))

        style = doc.styles['Normal']
        font = style.font
        font.name = "Calibri Light"

        lg_style = doc.styles.add_style('Larger', 1)
        lg_font = lg_style.font
        lg_font.name = "Calibri Light"
        lg_font.size = Pt(14)

        for row in doc.tables:
            for cell in row._cells:
                for key, value in data.items():
                        if key in cell.text:
                            cell.text = cell.text.replace(key, value)
                            print('Replaced "' + key + '" with "' + value + '" in table.')
                            for paragraph in cell.paragraphs:
                                paragraph.style = doc.styles['Normal']
                                if key in ['{week_no}', '{week_end}','{remaining_hrs}']:
                                    paragraph.style = doc.styles['Larger']
                                Report.turn_bold(paragraph)
        
        for paragraph in doc.paragraphs:
            for key, value in data.items():
                if key in paragraph.text:
                        for run in paragraph.runs:
                            run.text = run.text.replace(key, value)
                            print('Replaced "' + key + '" with "' + value + '" in paragraph.')
                        paragraph.style = doc.styles['Normal']
                        Report.turn_bold(paragraph)
                    
            # for key, value in data.items():
            #     if key in paragraph.text:
            #             for run in paragraph.runs:
            #                 run.text = run.text.replace(key, value)
            #             print('Replaced "' + key + '" with "' + value + '".')
        
        new_filename = str(self.out_file.name).replace("###", f"{self.week_no:03d}")
        with open(new_filename, 'w') as out_file:
            doc.save(self.out_file.with_name(new_filename))
        
        print('Saved output at ' + str(new_filename) + '.')
    
    def generate_data(self, week_start:str) -> dict:
        """Returns a dict of the data to be used in making the new document"""
        
        start_date = Internship.str_to_date(self.start_date)
        week_start = Internship.str_to_date(week_start)
        
        weekdays = self.get_week_dates(week_start)
        week_end = weekdays[-1]

        week_no = 0
        counter_week = start_date
        while counter_week < week_start:
            counter_week += timedelta(weeks=1)
            week_no += 1
        
        week_hrs = 6*5
        week_mins = week_hrs*60
        
        data_dict = { # TODO: Edit this to conform to the template document
        '{week_no}': str(week_no),
        '{week_start}': Internship.date_to_str(week_start),
        '{week_end}': str(week_end),
        '{day_1}': weekdays[0],
        '{day_2}': weekdays[1],
        '{day_3}': weekdays[2],
        '{day_4}': weekdays[3],
        '{day_5}': weekdays[4],
        '{week_hrs}': str(week_hrs),
        '{week_mins}': f'{week_mins:,}',
        '{remaining_hrs}': str(self.get_remaining_hrs())
        }

        print('Output: \n')
        for key, value in data_dict.items():
            print(f"{key}: {value}")

        return data_dict