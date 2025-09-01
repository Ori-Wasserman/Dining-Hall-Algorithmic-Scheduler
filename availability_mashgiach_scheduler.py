import pandas as pd
# returns the availability matrix needed for the alorithm as well as the names of the people 
def createAvailMatrix(MashgiachHoursFile):
    hours = pd.read_excel(MashgiachHoursFile)
    excel_file = pd.ExcelFile(MashgiachHoursFile)
    names = excel_file.sheet_names
    days = ["Monday", "Tuesday", "Wednesday", "Thursday"]
    hours_matrix = []
    min_slots_list = []
    max_slots_list = []
    for name in names:
        hours = pd.read_excel("Mashgiach hours optimization algorithm.xlsx", name)
        person_hours = []
        for day in days:
            person_hours += list(hours[day])
            # print(person_hours)
        hours_matrix.append(person_hours)
        min_slots_list.append(int(hours["min_slots"][0]))
        max_slots_list.append(int(hours["max_slots"][0]))

    return hours_matrix, names, min_slots_list, max_slots_list

def scheduleToExcel(mon, tues, wed, thurs, names):
    sched = pd.DataFrame({"Time" : [9 + .25*i for i in range(39)], "Monday": mon, "Tuesday":tues, "Wednesday":wed, "Thursday":thurs})
    for i in range(8):
        sched = sched.replace(i, names[i])
    sched.to_excel('outputSchedule.xlsx', index=False)
# print(createAvailMatrix("Mashgiach hours optimization algorithm.xlsx"))