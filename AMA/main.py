import classes
import os
from pathlib import Path

if __name__ == "__main__":
    cwd = Path(os.getcwd())
    # print('Current working directory: ', cwd)
    # start_date = "2024-9-23"
    # last_week_start = "2024-11-18"
    out_path = cwd / 'outputs' / 'output###.docx'
    # print('Outpath: ', out_path)
    template_path = cwd / 'template' / 'AMA_template.docx'
    # print('Template path: ', template_path)
    save_path = cwd / 'save-data.json'
    # print('Save path: ', save_path)

    # ama = classes.Report("1", "600", "0", start_date, last_week_start, template_path, out_path, save_path)
    # ama.generate_data(last_week_start)
    # ama.save_report()
    # ama = classes.Report.load_report("save-data.json")
    # ama.generate_data(last_week_start)

    # ama = classes.Report(1, 600, 0, "2024-09-23", "2024-09-23", template_path, out_path, save_path)
    ama = classes.Report.load_report(save_path)
    # ama.add_hour_difference(1)
    ama.update_internship()
    ama.edit_document()
    # if absent, uncomment:
    # ama.add_hour_difference(-6)
    # ama.edit_document(days_in_week=4)
    ama.save_report()
