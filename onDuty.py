"""
Author: Mike Fuson
Purpose: Created to help my brother quickly determine 
    which crew members are working in Nashville EMS.

Notes: Simplifies scheduling and reduces time spent
    figuring out shift assignments.

PS: Love You Bro
"""

import pandas as pd
import re
import math
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import os
import sys

ON_DUTY_CODES = [
    "EDIEMSEV15EDI",
    "EMSSHFDIF",
    "EMSSHFDIFES09",
    "EMSSHFDIFES33",
    "STWEP",
    "STWEA",
    "*STWEP",
    "*STWEA",
    "+OTEMS15",
    "+OOCEDTCPM",
    "+OOCEDTCAM",
    "EDI15-SE",
    "+OOCEDTCAMES21",
    "+OOCEDTCPMES21",
    "EDIEMSEV15EDI",    
    "+EDIEMSEV15EDI",
    "+OTEMS15-SS",
    "OTS15",
    "EMSSHFDIFEMSS01",
    "+OOCAC"
]
GENERIC_ON_DUTY_CODES = [
    "."
]
NOT_WORK_CODES = [
    "*ESTNWPM",
    "*ESTNWAM",
    ".ESTNWAM",
    ".ESTNWPM",
    ".HPM",
    ".HAM",
    ".PFLMPMMaternity",
    ".PFLMAMMaternity",
    ".FMLA-AM",
    ".FMLA-PM",
    ".PLPM",
    ".PLAM",
    ".EMS MLAM3",
    ".EMS MLPM3",
    ".PLPM",
    ".PLAM",
    ".VAM",
    ".VPM"
]

# Mapping of old field names to new ones a result of the excel import
FIELD_RENAME_MAP = {
    'Unnamed: 2': 'ID',
    'Unnamed: 3': 'Name',
    'Unnamed: 5': 'Code',
    'Unnamed: 6': 'From',
    'Unnamed: 7': 'Through',
    'Unnamed: 8': 'Hours'
}

# Apply renaming to each record before cleaning
def rename_fields(records):
    renamed_records = []
    for entry in records:
        renamed_entry = {}
        for k, v in entry.items():
            new_key = FIELD_RENAME_MAP.get(k, k)  # Replace if in map, else keep original
            renamed_entry[new_key] = v
        renamed_records.append(renamed_entry)
    return renamed_records

def get_report_dates(rosterFileName):
    wb = load_workbook(rosterFileName, data_only=True)
    ws = wb.active  # or wb["Sheet1"]
    centerHeader = ws.oddHeader.center.text if ws.oddHeader else None
    centerFooter = ws.oddFooter.center.text if ws.oddFooter else None

    m = re.search(r"\[(\d{2}/\d{2}/\d{4})\]", centerHeader)
    if m:
        report_date = m.group(1)
        report_date = report_date.replace("/", "-")
    else:
        print("No Report Date Found")
        report_date = ''

    m = re.search(r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})", centerFooter)

    if m:
        generated_date = m.group(1)
        generated_date = generated_date.replace("/", "-")
    else:
        print("No Generated Date/Time Found")
        generated_date = ''

    return report_date,generated_date

def load_roster_report(rosterFileName, base_dir=None):
    # If a base_dir is given and the filename is not already absolute
    if base_dir and not os.path.isabs(rosterFileName):
        file_path = os.path.join(base_dir, rosterFileName)
    else:
        file_path = rosterFileName

    # Normalize to absolute path
    file_path = os.path.abspath(file_path)

    df = pd.read_excel(file_path)
    data = df.to_dict(orient='records')
    cleanData = clean_fields(data)
    readableFieldsData=rename_fields(cleanData)
    
    return readableFieldsData

def is_empty_or_nan(value):
    return (
        value is None or
        (isinstance(value, float) and math.isnan(value)) or
        (isinstance(value, str) and value.strip() == '')
    )

def clean_fields(records):
    cleaned = []
    for entry in records:
        new_entry = {}

        for k, v in entry.items():
            # Skip keys with empty values
            if is_empty_or_nan(v):
                continue

            # Remove 'Unnamed: X' keys with empty values
            if re.match(r'^Unnamed: \d+$', k):
                if not is_empty_or_nan(v):
                    new_entry[k] = v
            else:
                new_entry[k] = v

        # Skip entries that are completely empty OR contain 'Rank' as any value
        if new_entry and 'Rank' not in [str(v).strip() for v in new_entry.values()]:
            cleaned.append(new_entry)

    return cleaned

def propagate_group(data, label_key="Group"):
    result = []
    current_label = None

    for entry in data:
        if isinstance(entry, dict):
            keys = list(entry.keys())

            # If it's a single-key dict (label)
            if len(keys) == 1:
                current_label = entry[keys[0]]
                continue  # Skip adding this entry to result

            # Apply the current label to the entry
            if current_label is not None:
                entry[label_key] = current_label

            result.append(entry)

    return result

def filter_on_duty(data, allowed_codes, cancel_codes, generic_codes):
    """
    Filters a list of dicts (with "Name" and "Code") to find who is on duty.

    On duty means:
      - A person has at least one allowed code (direct match or contains a generic code)
      - And they do not have any cancel codes.

    Args:
        data (list[dict]): Each dict has "Name" and "Code".
        allowed_codes (set[str]): Exact allowed code strings.
        cancel_codes (set[str]): Exact cancel code strings.
        generic_codes (set[str]): Generic substrings that may appear inside a Code.

    Returns:
        list[dict]: Filtered rows for on-duty names, only with allowed/generic codes.
    """

    def is_allowed(code):
        # Exact allowed OR contains any generic substring
        if code in allowed_codes:
            return True
        return any(g in code for g in generic_codes)

    # Build Name -> set of codes
    codes_by_name = {}
    for entry in data:
        name = entry.get("Name")
        code = entry.get("Code")
        if not name or not code:
            continue
        codes_by_name.setdefault(name, set()).add(code)

    # Determine who is on duty
    on_duty_names = set()
    for name, codes in codes_by_name.items():
        has_allowed = any(is_allowed(c) for c in codes)
        has_cancel = any(c in cancel_codes for c in codes)
        if has_allowed and not has_cancel:
            on_duty_names.add(name)

    # Return filtered rows
    result = []
    for entry in data:
        name = entry.get("Name")
        code = entry.get("Code")
        if name in on_duty_names and is_allowed(code):
            result.append(entry)

    return result

def assign_shift(data):
    for entry in data:
        from_time = entry.get("From")
        through_time = entry.get("Through")

        if from_time == "06:00": # and through_time == "18:00":
            entry["Shift"] = "AM"
        elif from_time == "18:00": # and through_time == "06:00":
            entry["Shift"] = "PM"
        else:
            entry["Shift"] = "Special"
    return data

def find_unique_lables(data, label_key="group"):
    unique_groups = set()
    labels=[]
    for entry in data:
        if isinstance(entry, dict) and label_key in entry:
            unique_groups.add(entry[label_key])

    for group in sorted(unique_groups):
        labels.append(group)
    return labels

def truncate_group_values(data, separator="/"):
    """
    Truncate the 'Group' field of each dictionary in a list to everything after the separator.

    Args:
        data (list of dict): Input array of dictionaries.
        separator (str): The delimiter to split on (default is "/").

    Returns:
        list of dict: Updated list with truncated 'Group' values.
    """
    for item in data:
        if "Group" in item and isinstance(item["Group"], str):
            # Split and strip spaces around the separator
            parts = item["Group"].split(separator)
            if len(parts) > 1:
                item["Group"] = parts[-1].strip()
    return data

def sort_by_first_string(data):
    """
    Sorts a list of lists by the first string in each sublist.

    Args:
        data (list): A list of lists where the first element of each sublist is a string.

    Returns:
        list: A new list sorted by the first element of each sublist.
    """
    return sorted(data, key=lambda x: x[0])



def create_ouput_spreadsheet(data,outFileName,reportDate,generatedDateTime):
    # for entry in range(0,20):
    #     print(data[entry])

    #Create workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "ON Duty"

    #Collect all unique column names (keys)
    medicHeader = ['Shift','Unit','Paramedic','EMT','3rd Employee','Shift','Unit','Paramedic','EMT','3rd Employee']
    cheifHeader = ['Shift','Unit','Chief','Shift','Unit','Chief']
    #Truncated the Group String
    data = truncate_group_values(data, separator="/")

    #Write each row of data
    groupLabels = find_unique_lables(data,'Group')
    groupLabelsSorted = sorted(groupLabels, key=lambda x: x.split("/", 1)[-1])


    asstRow = []
    medicRows =[]
    distRows = []
    specialRows =[]
    otherRows = []
    travelAMRows = []    
    travelPMRows = []
    AsstFound = False

    for group in groupLabelsSorted:
        entries = [entry for entry in data if group in entry.get('Group', '')]
        if 'Medic' in group:
            AMentries = [AMentry for AMentry in entries if 'AM' in AMentry.get('Shift', '')]
            PMentries = [PMentry for PMentry in entries if 'PM' in PMentry.get('Shift', '')]           

            if len(AMentries) > 3:
                print("More than 3 Employees on {} AM".format(group))
                AMrow = [AMentries[0]['Shift'],AMentries[0]['Group'],AMentries[0]['Name'],AMentries[1]['Name'],AMentries[2]['Name']]  
            elif len(AMentries) == 3:
                AMrow = [AMentries[0]['Shift'],AMentries[0]['Group'],AMentries[0]['Name'],AMentries[1]['Name'],AMentries[2]['Name']]
            elif len(AMentries) == 2:
                AMrow = [AMentries[0]['Shift'],AMentries[0]['Group'],AMentries[0]['Name'],AMentries[1]['Name'],'']
            elif len(AMentries) == 1:
                print("Less than 2 Employees on {} AM".format(group))
                AMrow = [AMentries[0]['Shift'],AMentries[0]['Group'],AMentries[0]['Name'],'','']
            elif len(AMentries) == 0:                
                print("No Employees on {} AM".format(group))
                AMrow = ['','','','','']

            if len(PMentries) > 3:
                print("More than 3 Employees on {} PM".format(group))
                PMrow = [PMentries[0]['Shift'],PMentries[0]['Group'],PMentries[0]['Name'],PMentries[1]['Name'],PMentries[2]['Name']]  
            elif len(PMentries) == 3:
                PMrow = [PMentries[0]['Shift'],PMentries[0]['Group'],PMentries[0]['Name'],PMentries[1]['Name'],PMentries[2]['Name']]
            elif len(PMentries) == 2:
                PMrow = [PMentries[0]['Shift'],PMentries[0]['Group'],PMentries[0]['Name'],PMentries[1]['Name'],'']
            elif len(PMentries) == 1:
                print("Less than 2 Employees on {} PM".format(group))
                PMrow = [PMentries[0]['Shift'],PMentries[0]['Group'],PMentries[0]['Name'],'','']
            elif len(PMentries) == 0:                              
                print("No Employees on {} PM".format(group))
                PMrow = ['','','','','']

            if not(all(item == '' for item in (AMrow+PMrow))):
                medicRows.append(AMrow+PMrow)

        elif 'On-Duty' in group:
            AsstFound = True
            AMentries = [AMentry for AMentry in entries if 'AM' in AMentry.get('Shift', '')]
            PMentries = [PMentry for PMentry in entries if 'PM' in PMentry.get('Shift', '')]

            if len(AMentries) == 1:
                AMASSrow = [AMentries[0]['Shift'],'ASST',AMentries[0]['Name']]
            elif len(AMentries) == 0:                              
                print("No ASS on Duty {} PM".format(group))
                AMASSrow = ['AM','ASST','']

            if len(PMentries) == 1:
                PMASSrow = [PMentries[0]['Shift'],'ASST',PMentries[0]['Name']]
            elif len(PMentries) == 0:                              
                print("No ASS on Duty {} PM".format(group))
                PMASSrow = ['PM','ASST','']

            if not(all(item == '' for item in (AMASSrow+PMASSrow))):
                asstRow = (AMASSrow+PMASSrow)
                
        elif ('District' in group) and not('EMS' in group):   
            AMentries = [AMentry for AMentry in entries if 'AM' in AMentry.get('Shift', '')]
            PMentries = [PMentry for PMentry in entries if 'PM' in PMentry.get('Shift', '')]

            if len(AMentries) == 1:
                AMdistrow = [AMentries[0]['Shift'],AMentries[0]['Group'],AMentries[0]['Name']]
            elif len(AMentries) == 0:                              
                print("No Chief on {} PM".format(group))
                AMdistrow = ['','','']

            if len(PMentries) == 1:
                PMdistrow = [PMentries[0]['Shift'],PMentries[0]['Group'],PMentries[0]['Name']]
            elif len(PMentries) == 0:                              
                print("No Chief on {} PM".format(group))
                PMdistrow = ['','','']

            if not(all(item == '' for item in (AMdistrow+PMdistrow))):
                distRows.append(AMdistrow+PMdistrow)
        elif 'Travelers AM' in group:
            if len(entries) != 0:
                for travNum in range(len(entries)):
                    travelAMRow = [entries[travNum]['From']+'-'+entries[travNum]['Through'],entries[travNum]['Group'],entries[travNum]['Name']]
                    if not(all(item == '' for item in (travelAMRow))):
                        travelAMRows.append(travelAMRow)
        elif 'Travelers PM' in group:
            if len(entries) != 0:
                for travNum in range(len(entries)):
                    travelPMRow = [entries[travNum]['From']+'-'+entries[travNum]['Through'],entries[travNum]['Group'],entries[travNum]['Name']]
                    if not(all(item == '' for item in (travelPMRow))):
                        travelPMRows.append(travelPMRow)            
        else:            
            if len(entries) != 0:
                for specNum in range(len(entries)):
                    specialRow = [entries[specNum]['From']+'-'+entries[specNum]['Through'],entries[specNum]['Group'],entries[specNum]['Name']]
                    if not(all(item == '' for item in (specialRow))):
                        specialRows.append(specialRow)

    if not(AsstFound):
        asstRow = ['AM','ASST','','PM','ASST','']
        print("No ASSes on Duty")
    ReportDateRow        = ['','','On-Duty Roster Report\n'+reportDate]
    GeneratedDateTimeRow = ['','','Roster Report Generated\n'+generatedDateTime]

    specialRows = sort_by_first_string(specialRows)

    ws.append(ReportDateRow)
    ws.append(GeneratedDateTimeRow)
    ws.append([])
    ws.append(cheifHeader)
    ws.append(asstRow)
    for row in distRows:
        ws.append(row)
    ws.append([]); ws.append([])
    ws.append(medicHeader)
    for row in medicRows:
        ws.append(row)
    ws.append([]);  
    for row in travelAMRows:
        ws.append(row)
    ws.append([]);  
    for row in travelPMRows:
        ws.append(row)
    ws.append([]);  
    for row in otherRows:
        ws.append(row)   
    ws.append([]);
    for row in specialRows:
        ws.append(row)

    #Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)  # e.g. "A", "B", "C"
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2  # add some padding
        ws.column_dimensions[col_letter].width = adjusted_width

    # --- Page landscape ---
    ws.page_setup.orientation = "landscape"

    # --- Scaling mode: Fit to width ---
    # This means it will shrink contents to fit the width of 1 page,
    # and let the height span across as many pages as needed
    ws.page_setup.scale = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 0 means "as many as needed"
    #Save to file
    wb.save(outFileName)

def browse_for_excel():
    # Create a hidden root window (so only the dialog shows)
    root = tk.Tk()
    root.withdraw()

    # Open file dialog
    file_paths = filedialog.askopenfilenames(
        title="Select Roster File(s)",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    root.destroy()
    if not file_paths:  # user canceled
        print("No input file selected. Exiting...")
        sys.exit(0)

    return list(file_paths)

def get_output_filename(reportDate,initialdir="."):
    import os, sys
    import tkinter as tk
    from tkinter import filedialog
    from datetime import datetime

    root = tk.Tk()
    root.withdraw()

    now_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
    default_name = f"On_Duty_Roster_{reportDate}.xlsx"

    # One popup: choose folder AND filename together
    path = filedialog.asksaveasfilename(
        title="Save As",
        initialdir=initialdir,
        initialfile=default_name,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    root.destroy()

    if not path:
        print("No output file selected. Exiting...")
        sys.exit(0)

    # Ensure .xlsx extension (asksaveasfilename should add it, but double-check)
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"

    return os.path.abspath(path)

def main():

    # inFileName = './rosters/Roster Report-2.xlsx'
    # outFileName = 'testOut.xlsx'

    inFileNames = browse_for_excel()
    for  inFileName in inFileNames:
        try:
            print("Input File Path: {} ".format(inFileName))  
            reportDate , generatedDateTime = get_report_dates(inFileName)
            roster = load_roster_report(inFileName)
            rosterWitGroup = propagate_group(roster)
            onDutyRoster = filter_on_duty(rosterWitGroup,ON_DUTY_CODES,NOT_WORK_CODES,GENERIC_ON_DUTY_CODES)
            onDutyShifts = assign_shift(onDutyRoster)
            outFileName = get_output_filename(reportDate)
            print("Output File Path: {} ".format(outFileName))
            create_ouput_spreadsheet(onDutyShifts,outFileName,reportDate,generatedDateTime)
        except Exception as e:
            print(f"Error processing {inFileName}: {e}")
            continue  # optional, loop naturally continues anyway

if __name__ == "__main__":
    main()