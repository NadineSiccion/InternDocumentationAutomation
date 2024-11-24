import classes
import os
from pathlib import Path

if __name__ == "__main__":
    cwd = Path(os.getcwd())
    print('Current working directory: ', cwd)
    start_date = "2024-9-23"
    last_week_start = "2024-09-23"
    out_path = cwd / 'outputs' / 'output###.docx'
    print('Outpath: ', out_path)
    template_path = cwd / 'template' / 'AMA_template.docx'
    print('Template path: ', template_path)
    save_path = cwd / 'save-data.json'
    print('Save path: ', save_path)

    ama = classes.Report("1", "600", "0", start_date, last_week_start, template_path, out_path, save_path)
    ama.edit_document()
