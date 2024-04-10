# excel_db.py
import openpyxl

category = {'RFP':3,'RFI':3,'GVE Support':1,'Complex Design':2}
complexity ={'High':3,'Medium':2,'Low':1}
status = {'WiP':1,'CE Pending':.1, 'Close':0}

headers = ['customer','project','category','cat_effort','complexity','com_effort','effort','status','status_effort','region','tsa','user','total_hrs']

excel_file_path = 'gve_sp_americas_report.xlsx'
gve_sheet = 'Bot'

class gve_record:
    def __init__(self, headers):
        # Initialize a dictionary with keys from the headers list and None as values
        self.headers = headers
        self.records = {header: None for header in headers}
    
    def ask_for_new_record(self):
        # Ask the user for input for each header and create a new record
        new_record = {}
        for header in self.headers:
            if 'effort' not in header:
                self.records[header] = input(f"Enter {header}: ")
            else:
                self.records[header] = None

        return self.records
    
    # no es util
    def add_new_record(self, record):
        # Adds the new record to the records dictionary
        for key, value in record.items():
            if key in self.records:
                self.records[key] = value
            else:
                print(f"Record with key '{key}' does not exist in the headers.")
    
    # no es util
    def update_record(self, key, value):
        # Update an existing record
        if key in self.records:
            self.records[key] = value
            return True
        else:
            print(f"Record with key '{key}' does not exist.")
            return False
    
    def get_records(self):
        # Return all records
        return self.records
    
    def get_headers(self):
        return self.headers
sample = {'data': 'add', 'category': 'RFP', 'complexity': 'High', 'customer': '2', 'project': '3', 'status': 'WiP'}

#new_gve_record = gve_record(headers)
def add_gve_record(record_input):
    
    new_gve_record = gve_record(headers)
    
    for key, value in record_input.items():
        if key in new_gve_record.records.keys():
            new_gve_record.records[key] = record_input[key]
    new_gve_record.records['cat_effort'] = category[new_gve_record.records['category']]
    new_gve_record.records['com_effort'] = complexity[new_gve_record.records['complexity']]
    new_gve_record.records['status_effort'] = status[new_gve_record.records['status']]
    new_gve_record.records['effort'] = new_gve_record.records['cat_effort'] * new_gve_record.records['com_effort']
    new_gve_record.records['total_hrs'] = new_gve_record.records['effort'] * new_gve_record.records['status_effort']
    
    return update_excel(new_gve_record.records)


def update_excel(new_record):
    print(f"El nuevo registro es {new_record}")
    new_data = list(new_record.values())
    
    workbook = openpyxl.load_workbook(excel_file_path)

    # Select a worksheet by name
    worksheet = workbook[gve_sheet]

    worksheet.append(new_data)
    
    workbook.save(filename=excel_file_path)
    
    for row in worksheet.iter_rows(values_only=True):
        print(row)

    return f'ok se actualizo excel {new_data}'