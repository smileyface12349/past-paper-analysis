from openpyxl import Workbook

from utils import format_question_number, Q_COLUMNS, COLUMNS

# create spreadsheet
ss = Workbook()

# save spreadsheet
print('')
print("Welcome to the setup process for a past paper. This will create a new spreadsheet to record the results.")
print("This is the first stage, to be carried about before completing the past paper")
while True:
    print('')
    print("Please input the file name to use for the spreadsheet")
    print("It must be a valid file name, excluding the extension")
    filename = input("Spreadsheet File Name: ")
    if '.' in filename:
        print("Do not include the file extension in the name")
    filename += '.xlsx'
    try:
        ss.save(filename)
        print(f"Successfully created a spreadsheet with name {filename}")
        break
    except Exception as e:
        print("The following error was produced when attempting to save the file:")
        print(e)
        print("Please ensure that the file name is valid")


# setup sheets
sheet = ss.active
sheet.title = "Questions"
sheet.freeze_panes = 'A2'
analysis_sheet = ss.create_sheet(title="Analysis", index=1)


# create headings
HEADINGS = {
    'A': 'Q',
    'B': 'P',
    'C': 'P',
    'D': 'P',
    'E': 'Marks Attained',
    'F': 'Marks Available',
    'G': 'Percentage',
    'H': 'Time Taken (seconds)',
    'I': 'Type of Question',
    'J': 'Topic',
    'K': 'Final Given Answer',
    'L': 'Correct Answer',
    'M': 'Time (minutes)',
    'N': 'Time / mark',
    'O': 'Time Reviewing',
    'P': 'Marks Before Reviewing'
}

# TODO: Set width of columns
# TODO: Display percent as percent
# TODO: Colour scales on relevant columns
# TODO: Formulas for time in minutes and time per mark
# TODO: Formulas (in analysis) for total mark, percent, total time

for column in HEADINGS:
    sheet[column+'1'] = HEADINGS[column]


# add in questions
def print_help():
    print('')
    print('INSTRUCTIONS')
    print("Input the relevant details for each question on the paper. You can either:")
    print(" - Type + or ] to go in one part (e.g. from 1 -> 1a)")
    print(" - Type - or [ to go up one level, and progress to next part (e.g. from 1a -> 2, 2aii -> 2b)")
    print(" - Type in the question details for the shown question number [details below]")
    print("If you make a mistake, you can type 'back' or edit it manually in the spreadsheet AFTER the program completes")
    print('')
    print('QUESTION INPUT FORMAT')
    print(f"<available marks> [question type] [topic]")
    print(f"Question type should be one of: multiple-choice (0), calculation (1), written (2), other (3)")
    print(f" - Custom question types are allowed as long as they do not include any spaces")
    print(f"If [topic] is not specified, the topic entered most recently will be used")
    print(" e.g. 1 0 electricity  (1 mark multiple choice on electricity)")
    print(" e.g. 6 2              (6 mark written question on previous topic")
    print(" e.g. 3 c moments      (you can use the first letter to specify question type)")
    print(" e.g. 1 o graphs       (e.g. drawing best fit line)")
    print(" e.g. 4                (you can set the type & topic later if you wish)")
    print('')
    print("Type 'stop' once the end of the paper has been reached")
    print("Type 'help' to get this information again")
    print('')

print_help()

PREDEFINED_QUESTION_TYPES = {
    '0': 'multiple-choice',
    'm': 'multiple-choice',
    '1': 'calculation',
    'c': 'calculation',
    '2': 'written',
    'w': 'written',
    '3': 'other',
    'o': 'other'
}

# TODO: Make question type optional

depth = 0
previous_topic = 'no topic provided'
current_question = [0, 0, 0, 0]
next_vacant_row = '2'
while True:
    formatted_question_number = [None, None, None, None]  # only includes changing parts e.g. [None, None, ii]
    all_question_numbers = [None, None, None, None]  # includes all parts e.g. [1, b, ii]
    render_on_sheet = True
    for x in reversed(range(depth+1)):
        all_question_numbers[x] = format_question_number(current_question[x], x)
        if render_on_sheet:
            formatted_question_number[x] = format_question_number(current_question[x], x)
        if current_question[x] != 0:
            render_on_sheet = False  # don't display same number twice

    question_number = ''.join(['('+q+')' for q in all_question_numbers if q is not None])

    details = input(f"{question_number}: ").lower()
    if details == 'help':
        print_help()
    elif details == 'stop':
        break
    elif details == 'back':
        current_question[depth] -= 1
    elif details in ['+', ']', '>', ' ', 'in', 'tab', 'i']:
        if depth >= 3:
            print("Maximum depth reached - if the question has smaller sub-parts, they will have to be grouped together into one question")
        else:
            depth += 1
    elif details in ['-', '[', '<', 'out', 'up', 'next', 'n', '']:
        # TODO: -- to go up 2 levels etc.
        if depth <= 0:
            print("Minimum depth reached. Type 'stop' to stop")
        else:
            depth -= 1
            if current_question[depth+1] != 0:  # don't move onto next part if this part is empty (e.g. 2bi -> 2b, 2bii -> 2c)
                current_question[depth] += 1
                for x in range(depth+1, 3):
                    current_question[x] = 0
    else:
        try:
            marks, question_type, topic = details.split(' ', 3)
            previous_topic = topic
        except:
            try:
                marks, question_type = details.split(' ', 2)
                topic = previous_topic
            except:
                try:
                    marks = details
                    question_type = 'none provided'
                    topic = previous_topic
                except:
                    print("Invalid input. Type 'help' for help")
                    continue
        try:
            marks = int(marks)
        except:
            print("Number of marks must be an integer")
            continue
        if question_type in PREDEFINED_QUESTION_TYPES:
            question_type = PREDEFINED_QUESTION_TYPES[question_type]

        # add in question numbers
        for i, q in enumerate(formatted_question_number):
            if not q:
                continue
            if q.isdigit():
                q = int(q)
            sheet[Q_COLUMNS[i]+next_vacant_row] = q

        # add in other stats
        sheet[COLUMNS['total marks']+next_vacant_row] = marks
        sheet[COLUMNS['topic']+next_vacant_row] = topic.title()
        sheet[COLUMNS['type']+next_vacant_row] = question_type.title()

        # add in formulae
        sheet[COLUMNS['percent']+next_vacant_row] = f"={COLUMNS['mark attained']+next_vacant_row}/{COLUMNS['total marks']+next_vacant_row}"
        sheet[COLUMNS['time minutes'] + next_vacant_row] = f"={COLUMNS['time'] + next_vacant_row}/60"
        sheet[COLUMNS['time per mark'] + next_vacant_row] = f"={COLUMNS['time minutes'] + next_vacant_row}/{COLUMNS['total marks'] + next_vacant_row}"

        # save (in case the program stops working)
        ss.save(filename)

        print(f"Question {question_number} was added successfully")

        current_question[depth] += 1
        next_vacant_row = str(int(next_vacant_row) + 1)


print('')
print("Spreadsheet created successfully")
print(f"You can find this spreadsheet in the current directory, named {filename}")
input("Press ENTER to quit the program")
