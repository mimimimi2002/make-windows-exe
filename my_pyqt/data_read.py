import pandas as pd
import json

def data_read(file_path):

    judge_df = pd.read_excel(file_path,sheet_name=0,header=None)
    program_df = pd.read_excel(file_path,sheet_name=6)
    xls = pd.ExcelFile(file_path)

    # Extract all sheet names
    sheet_names = xls.sheet_names

    def read_option(sheet_id, row, end):
        option_df = judge_df.iloc[row+1, 1:end]
        connected = '/'.join(option_df.iloc[:end-1].astype(str).values.flatten())
        return sheet_names[sheet_id] + "/" + connected

    def get_program_data(program_name):
        program_data = {
                "program_name": program_name,
                "url": str(program_df[program_name][0]),
                "outline": " ".join(program_df[program_name][3].split()),
                "period": " ".join(program_df[program_name][4].split()),
                "schedule": " ".join(str(program_df[program_name][5]).split()),
                "cost": " ".join(str(program_df[program_name][6]).split()),
                "application_grade": " ".join(str(program_df[program_name][7]).split()),
                "contact": " ".join(str(program_df[program_name][8]).split()),
                "image_name": str(program_df[program_name][1])
            }
        return program_data

    judge_data = {}
    option_data = {}
    option_count = {}

    for i in range(judge_df.shape[0] - 1):
        programs_list = judge_df.iloc[i+1, 2:].dropna()
        query = f"H1P{i+1}"

        # Initialize a list to store each program's data
        programs = []
        option = ""

        for j in range(len(programs_list)):
            program_name = str(judge_df.iloc[i+1, 2:].dropna()[j+2])

            # Create a dictionary for each program
            program_data = get_program_data(program_name)
            option = read_option(0, i, 2)

            programs.append(program_data)

        judge_data[query] = programs

        if option:
            option_data[query] = option

    judge_df = pd.read_excel(file_path,sheet_name=1,header=None)

    option_count["P"] = i + 1

    l = 0
    p = 0
    for i in range(judge_df.shape[0] - 1):
        programs_list = judge_df.iloc[i + 1, 3:].dropna()
        p += 1
        if i % 4 == 0:
            l += 1
            p = 1

        query = f"H2L{l}P{p}"
        program_entries = []
        option = ""

        for j in range(len(programs_list)):
            program_name = str(judge_df.iloc[i + 1, 3:].dropna()[j + 3])

            program_data = get_program_data(program_name)
            option = read_option(1, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["L"] = l + 1

    judge_df = pd.read_excel(file_path,sheet_name=2,header=None)
    s = 0
    p = 0
    for i in range(judge_df.shape[0] - 1):
        programs_list = judge_df.iloc[i + 1, 3:].dropna()
        p += 1
        if i % 4 == 0:
            s += 1
            p = 1

        query = f"H3S{s}P{p}"
        program_entries = []
        option = ""

        for j in range(len(programs_list)):
            program_name = str(judge_df.iloc[i + 1, 3:].dropna()[j + 3])

            program_data = get_program_data(program_name)
            option = read_option(2, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["S"] = s + 1

    judge_df = pd.read_excel(file_path,sheet_name=3,header=None)

    pp = 1
    p = 1
    for i in range(judge_df.shape[0]-1):
        programs_list = judge_df.iloc[i+1,3:].dropna()
        p += 1
        if i % 4 == 0:
            pp += 1
            p = 1

        query = f"H4PP{pp}P{p}"
        program_entries = []
        option = ""

        for j in range(len(programs_list)):
            program_name = str(judge_df.iloc[i + 1, 3:].dropna()[j + 3])

            program_data = get_program_data(program_name)
            option = read_option(3, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["PP"] = pp + 1

    judge_df = pd.read_excel(file_path,sheet_name=4,header=None)

    style = 1
    p = 1
    for i in range(judge_df.shape[0]-1):
        programs_list = judge_df.iloc[i+1,3:].dropna()
        p += 1
        if i % 4 == 0:
            style += 1
            p = 1

        query = f"H5Style{style}P{p}"
        program_entries = []
        option = ""

        for j in range(len(programs_list)):
            program_name = str(judge_df.iloc[i + 1, 3:].dropna()[j + 3])

            program_data = get_program_data(program_name)
            option = read_option(4, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["Style"] = style + 1

    judge_df = pd.read_excel(file_path,sheet_name=5,header=None)

    for i in range(judge_df.shape[0] - 1):
        programs_list = judge_df.iloc[i+1, 2:].dropna()
        query = f"H6E{i+1}"

        # Initialize a list to store each program's data
        programs = []
        option = ""

        for j in range(len(programs_list)):
            program_name = str(judge_df.iloc[i+1, 2:].dropna()[j+2])

            # Create a dictionary for each program
            program_data = get_program_data(program_name)
            option = read_option(5, i, 2)

            programs.append(program_data)

        judge_data[query] = programs
        if option:
            option_data[query] = option

    option_count["E"] = i + 1

    return judge_data, option_data, option_count