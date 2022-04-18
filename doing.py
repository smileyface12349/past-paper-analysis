import math
import os
import time

import openpyxl
from openpyxl import Workbook


# open spreadsheet
from utils import Q_COLUMNS, format_question_number, COLUMNS

print("Welcome to the past paper utility")
print("You are in the second stage - completing the past paper")
while True:
    print('')
    print("Please input the file name of the spreadsheet. Alternatively, input a number corresponding to one of the options listed below")
    print("The file should be stored in the current directory. Otherwise, you can provide the full path to the file")
    possibilities = []
    for file in os.listdir():
        if 'xls' in file:
            possibilities.append(file)
            print(f"({len(possibilities)-1}) {file}")
            if len(possibilities) >= 9:
                break
    filename = input("File Name: ")
    if filename == '':
        filename = '0'
    if filename.isdigit() and len(filename) == 1:
        i = int(filename)
        try:
            filename = possibilities[i]
        except:
            print(f"There is no file {i} in the list of options")
            continue
    if '.' not in filename:
        filename += '.xlsx'
    try:
        ss = openpyxl.load_workbook(filename)
        break
    except Exception as e:
        print("The following error was produced when attempting to load the file:")
        print(e)
        print("Please ensure that the file exists and the name is spelt correctly")


# choose sheet
sheet = ss['Questions']


# enter time limit
print("")
print("Are you ready to begin?")
print("Please enter the time (in minutes) allowed for the paper")
print("NOTE: This is only used to state how much time you have left. You will not be stopped if you exceed the time limit")
print("You must initially complete the paper in order, but you are encouraged to skip past questions. At the end, you will be given the opportunity to review your answers")
print("You do not have to enter answers, and could instead use the program to simply record time used. In this case, press enter without entering an answer")
print("THE CLOCK WILL BEGIN AS SOON AS YOU PRESS ENTER")
minutes = input("Number of minutes allowed for the exam: ")
while True:
    if minutes.isdigit() and 5 < int(minutes) < 1440:
        minutes = int(minutes)
        break
    else:
        minutes = input("Please enter a valid number of minutes: ")

print('')
print("The time has started. Good luck!")
print("Type 'pause' at any time to pause the timer. This should only be used for disruptions beyond your control.")
print('')

# do the paper
current_row = 2
previous_question = [None, None, None, None]  # this should never be accessed
total_elapsed = 0
while True:
    # not all values are included in the spreadsheet
    # if omitted, they must inherit from the previous row (if before any given values)
    higher_level_changed = False
    current_question = []
    blank_counter = 0
    for i in range(4):
        cell = Q_COLUMNS[i]+str(current_row)
        value = sheet[cell].value
        if value:
            current_question.append(value)
            higher_level_changed = True
        else:
            blank_counter += 1
            if higher_level_changed:
                current_question.append(None)
            else:
                current_question.append(previous_question[i])

    if blank_counter >= 4: # empty row in spreadsheet
        print('')
        print("END OF QUESTIONS")
        break

    question_number = ''.join(['('+str(q)+')' for q in current_question if q])
    marks = sheet[COLUMNS['total marks'] + str(current_row)].value

    start = time.time()
    elapsed = 0
    while True:
        print('')
        print(f"{question_number} [{marks} marks]")
        answer = input("Answer: ")
        if answer.lower() == 'pause':
            elapsed += time.time() - start
            sheet[COLUMNS['time']+str(current_row)] = elapsed
            ss.save(filename)
            total_elapsed += elapsed
            print('')
            print("The timer is now paused")
            input("Press ENTER when you would like to resume")
            print("The timer has now been resumed")
            start = time.time()
        else:
            elapsed += time.time() - start
            total_elapsed += elapsed
            sheet[COLUMNS['time'] + str(current_row)] = elapsed
            if answer:
                sheet[COLUMNS['given answer']+str(current_row)] = answer
            ss.save(filename)
            break

    print(f"That question took you {elapsed/60:.1f} minutes for {marks} marks.")
    print(f"You have used {math.floor(total_elapsed / 60)}/{minutes} minutes ({math.floor(total_elapsed / 60 / minutes * 100)}%)")

    current_row += 1


# TODO: Checking answers
print('')
print("You are welcome to go back and check your answers")
print(f"You have {minutes-total_elapsed/60:.1f} minutes remaining")

start = time.time()
print("NOTE: Time spent on each individual question is not currently tracked during review.")
input("Press ENTER once you are done checking")
print(f"You spent {(time.time()-start)/60:.3f} minutes checking your answers")
print(f"You still have {minutes-total_elapsed/60-(time.time()-start)/60:.2f} minutes of unused time")

print('')
input("Press ENTER to close the program")