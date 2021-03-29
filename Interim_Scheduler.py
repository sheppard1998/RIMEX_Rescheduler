import pandas as pd
import numpy as np

def reschedule(schedule):
    """
    Aim:
        Take in the spreadsheet and sort schedule jobs according to restrictions

    Restrictions:

        If Line Status == 'Off' ignore it
        If Line Status == 'On' Schedule on that row

        If Wknd == N assign 40 for the day
        If Wknd == Y assign 20 for the day

    """

    width = len(schedule.columns) #getting width
    length = len(schedule) #getting length

    for i in range(width - 3):
        for j in range(length - 7):
            schedule.iloc[j + 6, i + 3] = '' #clearing schedule

    j = 0 #initialising row index

    for i in range(width - 3):
        rem_to_sched = float(schedule.iloc[2, i + 3]) #getting quantity

        status = schedule.iloc[j + 6, 2] #getting status
        wknd = schedule.iloc[j + 6, 1] #getting if its a weekend

        if wknd == 'N': #Deciding the days max
            max_to_sched = 40
        else:
            max_to_sched = 20

        while rem_to_sched > 0 and status == 'On' and j < (length - 7):
            so_far = 0 #how many have been scheduled for the day
            for k in range(i):
                if schedule.iloc[j + 6, k + 3] != '':
                    so_far += schedule.iloc[j + 6, k + 3] #incrementing scheduled

            max_to_sched -= so_far #decrementing max you can now schedule

            if max_to_sched >= rem_to_sched:
                #if max is higher than remaining then schedule all remaining
                if rem_to_sched != 0:
                    schedule.iloc[j + 6, i + 3] = rem_to_sched
                    rem_to_sched = 0
                else:
                    pass
            else:
                #if max is less than remaining schedule the max
                if max_to_sched != 0:
                    schedule.iloc[j + 6, i + 3] = max_to_sched
                    rem_to_sched -= max_to_sched
                    max_to_sched = 0
                else:
                    pass

                j += 1 #increment row index

                status = schedule.iloc[j + 6, 2] #get new status
                wknd = schedule.iloc[j + 6, 1] #get new weekend

                if wknd == 'N': #reset max
                    max_to_sched = 40
                else:
                    max_to_sched = 20

    return schedule

def changeRank(ranks_lol, schedule):
    """
    Change the rank of a job to the new rank assigned and sort the jobs based on
    these new ranks.

    Example test: Change In-Metals to be before Steel Service
    """

    width = len(schedule.columns) #getting width
    length = len(schedule) #getting length

    for i in range(width - 3):

        flag = True
        count = 0

        for j in range(len(ranks_lol)):
            if schedule.iloc[3, i + 3] == ranks_lol[j][0]:

                flag = False #setting rank to not needing change

            elif schedule.iloc[3, i + 3] + count == ranks_lol[j][1]:

                count += 1 #counting how many the rank must be shifted down by

        if flag == True and count != 0: #appending collateral rank changes
            ranks_lol.append([schedule.iloc[3, i + 3], schedule.iloc[3, i + 3] + count])


    for i in range(len(ranks_lol)):
        old_rank = ranks_lol[i][0]
        new_rank = ranks_lol[i][1]
        schedule.iloc[3, old_rank + 2] = new_rank #changing rank

    schedule.iloc[3, 2] = np.NaN #allow for sort exclusion

    schedule.sort_values(by=3, axis=1, inplace=True, na_position = 'first')

    schedule.iloc[3, 2] = 'RANK'

    return schedule

def main():

    file_question = "Please enter the filename and extension of the schedule"\
    " (i.e .xlsx) you would like to edit. Please ensure that the file is in "\
    "the same folder as this program: \n"

    excel_file = input(file_question) #getting file name
    print()

    new_filename_question = "What would you like to call this new schedule, "\
    "please include the file extension (i.e .xlsx): \n"
    new_filename = input(new_filename_question) # getting copy decision
    print()

    valid_input = False

    while valid_input == False:

        intro = "This program allows you to change the rank of a job and/or "\
        "reschedule. If you would like to change the rank and reschedule, type "\
        "'RANK'. If you would like to just reschedule based on manual changes to "\
        "the sheet (i.e change in line status, addition/removal of job etc), type "\
        "'RESCHEDULE': \n"

        task = input(intro) #getting desired task
        print()

        if task == 'RANK' or task == 'RESCHEDULE':
            valid_input = True
        else:
            print()
            print("You have entered an invalid input.\n")
            print()

    if task == 'RANK':
        current = "Please type the rank of the job you would like to change: \n"
        new = "Please type the rank you would like to change this job to: \n"
        ranks_lol = []

        done = False

        while done == False:

            old_rank = int(input(current)) #getting old rank
            new_rank = int(input(new)) #getting new rank

            rank_change = [old_rank, new_rank]

            ranks_lol.append(rank_change) #building rank list of lists

            valid_input = False

            while valid_input == False:
                another_question = "Would you like to change another rank? "\
                "Type 'Y' for yes and 'N' for no: \n"
                another = input(another_question) #getting another decision
                print()
                if another == 'N':
                    done = True
                    valid_input = True
                elif another == 'Y':
                    valid_input = True
                else:
                    print()
                    print("You have entered an invalid input.\n")
        schedule = pd.read_excel(excel_file) #opening excel to dataframe
        schedule = changeRank(ranks_lol, schedule) #changing ranks
        schedule = reschedule(schedule) #rescheduling

    elif task == 'RESCHEDULE':
        #reschedule
        schedule = pd.read_excel(excel_file) #opening excel to dataframe
        schedule = reschedule(schedule) #rescheduling

    schedule.to_excel(new_filename) #saving

    print("Please check the folder this program is in for the new schedule")
main()
