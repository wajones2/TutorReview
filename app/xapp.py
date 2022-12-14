import openpyxl
from pathlib import Path
from datetime import datetime


class Summary:
    def __init__(self, xlsx):
        self.xlsx = xlsx

        xlsx_file = Path('./docs/', self.xlsx)
        wb_obj = openpyxl.load_workbook(xlsx_file) 

        # Read the active sheet:
        self.sheet = wb_obj.active
        # What can we do better?

        self.student_dict = {}
        self.reusable_dict = {}
        self.template_dict = {}

    def workbook(self):
        zero_to_one = 0
        counter = 0
        row_counter = 0         # Stops program from going any farther in a row than the last cell entry
        col_stopper = 0
        self.idx_order = [1,3,4,5,6,7,8,0]
        self.reusable_list = []
        self.data_list = []
        for row in self.sheet:
            if row[0].value == None:             # Stops program from going any lower than the last cell entry
                self.organize_parent_2()
                self.organize_student_dict()
                break
            for cell in row:
                if row_counter == 0:
                    if cell.value == "What can we do better?":  
                        self.organize_parent_3_keys()
                        self.reusable_list = []
                        row_counter += 1
                        col_stopper += counter + 1
                        counter = 0    
                        break
                    counter += 1
                elif row_counter == 1:
                    self.data_list.append(cell.value)
                    if counter == 2:
                        self.reusable_list.append(cell.value)

                    if counter == col_stopper:
                        counter = 0
                        if zero_to_one == 0:
                            self.data_list[:] = [self.data_list[:-1]]
                            zero_to_one += 1
                        else:
                            self.data_list[zero_to_one:] = [self.data_list[zero_to_one:-1]]
                            zero_to_one += 1
                        break
                    else:
                        counter += 1
            # print()

    def organize_parent_3_keys(self):
        self.p3k = ["Name", "PSID", "Class", "Tutor", "Experience", "Comments", "Recommendations", "Timestamp"]
        for label in self.p3k:

            self.reusable_dict[label] = []
        self.template_dict = self.reusable_dict.copy()
        
    def organize_parent_2(self):
        self.session_count = []
        self.student_session_dict = {}
        session_check = self.reusable_list[:]
        for zero in range(len(self.reusable_list)):
            self.session_count.append(0)
        for idx in range(len(self.reusable_list)):
            self.session_count[idx] = (self.reusable_list.count(self.reusable_list[idx]) - session_check.count(session_check[0])) + 1
            session_check.pop(0)

    def organize_student_dict(self):
        session_dict = {}
        name_dict = {}
        counter = 0
        for student in self.reusable_list:
            self.student_dict[student] = {}
        for row in self.data_list:
            for idx in range(len(self.p3k)):
                self.reusable_dict[self.p3k[idx]] = row[self.idx_order[idx]]
            session_dict[self.session_count[counter]] = self.reusable_dict
            key = list(session_dict.keys())[0]
            value = list(session_dict.values())[0]
            name = row[2]
            self.student_dict[name].update(session_dict)
            self.reusable_dict = self.template_dict.copy()
            session_dict = {}
            counter += 1

    def experience_average(self):
        self.experience = 0
        self.submissions = 0
        for student in self.student_dict:
            for session in self.student_dict[student]:
                self.experience += self.student_dict[student][session]['Experience']
                self.submissions += 1
        self.experience /= self.submissions
        self.experience = float('{:.2f}'.format(self.experience))


    def organize_timestamps(self):
        # The only dates out of order are the ones where the 'tutor' is 
        # not just 'Whitney J'

        ts = []
        for student in f21.student_dict:
            for session in f21.student_dict[student]:
                ts.append(f21.student_dict[student][session]['Timestamp'])


    
f21 = Summary('F21.xlsx')
f21.workbook()
f21.experience_average()

s22 = Summary('S22.xlsx')
s22.workbook()
s22.experience_average()


combined_experience = (f21.experience + s22.experience) / 2
combined_submissions = f21.submissions + s22.submissions
print(f'Out of {combined_submissions} student reviews, '  
        f'I have a rating of {combined_experience} out of 5.')




