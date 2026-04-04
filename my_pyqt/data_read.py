import pandas as pd
import json

def data_read(file_path):

    program_df = pd.read_excel(file_path,sheet_name=6)
    xls = pd.ExcelFile(file_path)

    # Extract all sheet names
    sheet_names = xls.sheet_names

    def read_option(judge_df, sheet_id, row, end):
        option_df = judge_df.iloc[row, 1:end]
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

    # 期間で決めるページのread data
    period_judge_df = pd.read_excel(file_path,sheet_name=0)

    for i in range(period_judge_df.shape[0]):
        programs_list = period_judge_df.iloc[i, 2:].dropna()
        query = f"H1P{i+1}"

        # Initialize a list to store each program's data
        programs = []
        option = ""

        for program_name in programs_list:
            # Create a dictionary for each program
            program_data = get_program_data(program_name)
            option = read_option(period_judge_df, 0, i, 2)

            programs.append(program_data)

        judge_data[query] = programs

        if option:
            option_data[query] = option

    option_count["P"] = i + 1

    period_option_count = i + 1

    #　言語で決めるページ
    language_judge_df = pd.read_excel(file_path,sheet_name=1)

    l = 0
    p = 0
    for i in range(language_judge_df.shape[0]):
        programs_list = language_judge_df.iloc[i, 3:].dropna()
        p += 1
        if i % period_option_count == 0:
            l += 1
            p = 1

        query = f"H2L{l}P{p}"
        program_entries = []
        option = ""

        for program_name in programs_list:
            program_data = get_program_data(program_name)
            option = read_option(language_judge_df, 1, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["L"] = l

    # 長期休みで決めるページ
    season_judge_df = pd.read_excel(file_path,sheet_name=2)
    s = 0
    p = 0
    for i in range(season_judge_df.shape[0]):
        programs_list = season_judge_df.iloc[i, 3:].dropna()
        p += 1
        if i % period_option_count == 0:
            s += 1
            p = 1

        query = f"H3S{s}P{p}"
        program_entries = []
        option = ""

        for program_name in programs_list:
            program_data = get_program_data(program_name)
            option = read_option(season_judge_df, 2, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["S"] = s

    # 目的で決めるページ
    purpose_judge_df = pd.read_excel(file_path,sheet_name=3)

    pp = 0
    p = 0
    for i in range(purpose_judge_df.shape[0]):
        programs_list = purpose_judge_df.iloc[i ,3:].dropna()
        p += 1
        if i % period_option_count == 0:
            pp += 1
            p = 1

        query = f"H4PP{pp}P{p}"
        program_entries = []
        option = ""

        for program_name in programs_list:

            program_data = get_program_data(program_name)
            option = read_option(purpose_judge_df, 3, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["PP"] = pp

    #　留学スタイルで決めるページ
    style_judge_df = pd.read_excel(file_path,sheet_name=4)

    style = 0
    p = 0
    for i in range(style_judge_df.shape[0]):
        programs_list = style_judge_df.iloc[i ,3:].dropna()
        p += 1
        if i % period_option_count == 0:
            style += 1
            p = 1

        query = f"H5Style{style}P{p}"
        program_entries = []
        option = ""

        for program_name in programs_list:

            program_data = get_program_data(program_name)
            option = read_option(style_judge_df, 4, i, 3)

            program_entries.append(program_data)

        judge_data[query] = program_entries
        if option:
            option_data[query] = option

    option_count["Style"] = style

    # イベントで決めるページ
    event_judge_df = pd.read_excel(file_path,sheet_name=5)

    for i in range(event_judge_df.shape[0]):
        programs_list = event_judge_df.iloc[i, 2:].dropna()
        query = f"H6E{i+1}"

        # Initialize a list to store each program's data
        programs = []
        option = ""

        for program_name in programs_list:

            # Create a dictionary for each program
            program_data = get_program_data(program_name)
            option = read_option(event_judge_df, 5, i, 2)

            programs.append(program_data)

        judge_data[query] = programs
        if option:
            option_data[query] = option

    option_count["E"] = i + 1

    return judge_data, option_data, option_count

if __name__ == "__main__":
    data_read("/Users/maseiyou/Downloads/data_test/judge_data/updated_judge.xlsx")