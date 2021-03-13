from googleapiclient.discovery import build
from google.oauth2 import service_account
from collections import namedtuple
from itertools import combinations
import random

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://spreadsheets.google.com/feeds",
    'https://www.googleapis.com/auth/spreadsheets',
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

SERVICE_ACCOUNT_FILE = 'creds.json'
SAMPLE_SPREADSHEET_ID = '1LiWgohYT4O-ZgNTHM0bG9Nkv3kjKsKKs5M7awnu5O-I'

def run():
    creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    num_people = len(sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Input!A2:A1000").execute().get("values", []))
    tasks = get_tasks(sheet)
    num_tasks = len(tasks)
    people = get_people(num_people, sheet)
    history = get_history(num_people, sheet)
    block_groups = get_block_group(sheet)
    block_schedule = get_block_schedule(sheet, num_tasks, block_groups)
    
    output = []

    usage = []
    for i in range(num_people):
        person = []
        for j in range(num_tasks):
            person.append(0)
        usage.append(person)

    days = ['m', 'tu', 'w', 'th', 'f']
    days_id = 0
    for day in days:
        output_day = []
        task_id = 0
        for task in tasks:
            people_needed = int(task[0])
            cell = ''
            available_candidates = get_available_candidates(people, day, task_id)
            candidates = finalis_candidates(available_candidates, people, history, task_id)

            if people_needed < 2:
                if candidates:
                    chosen = priorities_candidates(candidates, usage, task_id).pop()
                    
                    cell += chosen[0]
                    usage[chosen[1]][task_id] += 1
                else:
                    cell += 'NSP' #No Suitable People
            else:
                groupings = get_all_groupings(candidates, people_needed)
                filtered_groupings = apply_blocks(groupings, block_groups, block_schedule, days_id, task_id)
                ordered_groupings = priorities_groupings(groupings, usage, task_id)
                chosen = ordered_groupings.pop()
                
                no_comma = len(chosen) - 1
                p_id = 0
                # print(chosen)
                for p in chosen:
                    cell += p[0]
                    usage[p[1]][task_id] += 1
                    if not p_id == no_comma:
                        cell += ', '
                    p_id += 1
    
                
            output_day.append(cell)
            task_id += 1
        output.append(output_day)
        days_id += 1
    shift_output = []
    for _ in range(num_tasks):
        row = []
        for _ in range(5):
            row.append([])
        shift_output.append(row)
    for i in range(5):
        for j in range(num_tasks):
            shift_output[j][i] = output[i][j] 
    update_sheet(sheet, shift_output)
    update_sheet_history(sheet, history, usage, num_people, num_tasks)
    # print(shift_output)
def update_sheet(sheet, data):
    range = "Output!D3"
    value_input_option = "USER_ENTERED"
    value_render_option = "DIMENSION_UNSPECIFIED"
    value_range_body = {
        'values': data,
        "majorDimension": value_render_option
    }
    sheet.values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID,
        range=range,
        valueInputOption=value_input_option,
        body=value_range_body,
    ).execute()
def update_sheet_history(sheet, history, usage, num_people, num_tasks):
    history_range = "History!B2"
    value_input_option = "USER_ENTERED"
    value_render_option = "DIMENSION_UNSPECIFIED"

    updated_history = []
    # print(num_people)
    for i in range(num_people):
        person = []
        for j in range(num_tasks):
            person.append(usage[i][j] + history[i][j])
        updated_history.append(person)
    value_range_body = {
        'values': updated_history,
        "majorDimension": value_render_option
    }
    sheet.values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID,
        range=history_range,
        valueInputOption=value_input_option,
        body=value_range_body,
    ).execute()
def apply_blocks(groupings, block_groups, block_schedule, day, task):
    # print(groupings)
    def is_group_blocked(block, group):
        block = block
        group = group
            
        for person in group:
            for b in range(1, len(block)):
                if person[0].find(block[b]) != -1:
                    # print(person[0], '==', block[b])
                    block.remove(block[b])
                    break
            if not block:
                break
        if not block:
            return True
        else:
            return False
    i = 0
    # print(groupings)
    while i < len(groupings):
        for b in range(len(block_groups)):
            if is_group_blocked(block_groups[b], groupings[i]):
                groupings.remove(groupings[i])
                print(groupings)
                i -= 1
                break
        i += 1
    
    return groupings
def get_all_groupings(candidates, people_needed):
    #generats all posable combonations that mach len of people_needed
    groupings = list(combinations(candidates, people_needed))
    return groupings             
def priorities_candidates(candidates, usage, task):
    temp = []
    
    for c in range(len(candidates)):
        temp.append([candidates[c][2] + usage[candidates[c][1]][task], c])
    def aux(e):
       return e[0]
    random.shuffle(temp)
    temp.sort(key=aux, reverse=True)
    output = []
    for t in temp:
        output.append(candidates[t[1]])
    return output
def priorities_groupings(groupings, usage, task):
    temp = []
    for c in range(len(groupings)):
        sum = 0
        for i in range(len(groupings[c])):
            sum += groupings[c][i][2]
            # print(task)
            sum += usage[groupings[c][i][1]][task]
        temp.append((sum, c))
    def aux(e):
        return e[0]
    # print(temp)
    random.shuffle(temp)
    # print(temp)
    temp.sort(key=aux, reverse=True)
    output = []
    for t in temp:
        output.append(groupings[t[1]])
    return output      
def get_available_candidates(people, day, task):
    output = []
    person_id = 0
    for person in people:
        if (person[task + 1]).find(day) != -1:
            output.append(person_id)
        person_id += 1
    return output
def finalis_candidates(candidates, people, history, task):
    output = []
    for c in candidates:
        output.append((people[c][0], c, int(history[c][task])))
    return output
def get_people(num_people, sheet):
    input_range = "Input!A2:P" + str(num_people + 1)
    people = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range=input_range).execute().get("values", [])
    return people 
def get_history(num_people, sheet):
    history_range = "History!B2:L" + str(num_people + 1) 
    history = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=history_range).execute().get(
                                     "values", [])
    history_convert = []
    for i in range(len(history)):
        person = []
        for j in range(len(history[i])):
            person.append(int(history[i][j]))
        history_convert.append(person)

    return history_convert
def get_tasks(sheet):
    tasks_range = "Output!B3:C1000"
    tasks = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=tasks_range).execute().get(
                                     "values", [])
    return tasks
def get_block_group(sheet):
    block_groups_range = "Block!A2:C1000"
    block_groups = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=block_groups_range).execute().get(
                                     "values", [])
    #[name of group, names contained by group*,..*n]
    blocks = []
    for block_group in block_groups:
        block = [block_group[0]]
        for name in block_group[1].split(', '):
            block.append(name)
        blocks.append(block)
    
    return blocks               
def get_block_schedule(sheet, num_task, block_groups):
    block_schedule_range = "Block!E2:J12"
    block_schedule = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=block_schedule_range).execute().get(
                                     "values", [])
    always_block_range = "Block!C2:C1000"
    always_block = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=block_schedule_range).execute().get(
                                     "values", [])
    schedule = []
    always_block_groups = []
    for ab in range(len(always_block)):
        for bg in range(len(block_groups)):
            if always_block[ab] == block_groups[bg][0]:
                always_block_groups.append(bg)
    for _ in range(5):
        days = []
        for _ in range(num_task):
            days.append([])
        schedule.append(days)
    for i in range(num_task):
        for j in range(5):
            
            for b in range(len(block_groups)):
                if block_schedule[i][j].find(block_groups[b][0]) != -1:
                    schedule[j][i].append(b)
                for ab in always_block_groups:
                    if schedule[j][i] != ab and b == ab:
                        schedule[j][i].append(ab)
                


    # print(schedule)
    return schedule

run()