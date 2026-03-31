import pandas as pd
import json
import os

def check_data(file_path):

  program_df = pd.read_excel(file_path,sheet_name=6)
  program_df = program_df.set_index(program_df.columns[0])
  xls = pd.ExcelFile(file_path)

  program_names = program_df.columns

  program_names_set = set()
  features = program_df[program_df.columns[0]]

  expected_features = ["url", "image_name", "画像", "outline", "period", "schedule", "cost", "applicable_grade", "contact"]

  error_messages = []

  # プログラム名が有効か
  def check_valid_program_name(judge_df, df_name, start_index):
    for i in range(len(judge_df)):
      programs_list = judge_df.iloc[i, start_index:].dropna()
      for index, program_name in enumerate(programs_list):
        if program_name not in program_names_set:
          error_messages.append(
                f"「{df_name}」ページの {i+2}行目, {index+4}列目の「{program_name}」のプログラムはプログラムリストに入っていません。"
                "「プログラムリスト」ページのプログラム名と一致させてください。"
            )

  # 期間の選択肢の数が有効か
  def check_valid_period_option_num(judge_df, df_name, expected_period_option_count):
    i = 0
    while i < len(judge_df):
        block = judge_df.iloc[i:i+expected_period_option_count]
        if len(block) != expected_period_option_count:
          error_messages.append(
              f" 「{df_name}」ページの{i - expected_period_option_count + 2}行目の選択肢の数が間違っています。"
                "期間の選択肢の数を、「期間で決める」ページの選択肢の数と揃えてください。"
                "例：期間で決めるページの選択肢が4つあれば、同じ言語レベルの選択肢を4つにしてください"
          )
          break
        if (block.iloc[:, 1] != block.iloc[0, 1]).any():
          error_messages.append(
                f" 「{df_name}」ページの{i - expected_period_option_count + 2}行目の選択肢の数が間違っています。"
                "期間の選択肢の数を、「期間で決める」ページの選択肢の数と揃えてください。"
                "例：期間で決めるページの選択肢が4つあったら、同じ言語レベルの選択肢を4つにしてください"
            )
          break
        i += expected_period_option_count


  # 留学プログラムリストの記入のチェック
  for i, program_name in enumerate(program_names):
      col_data = program_df[program_name]
      if col_data.isna().all():
        base_dir = os.path.dirname(file_path)
        image_path = os.path.join(base_dir, "../image", f"image{i+1}.png")
        print(image_path)

        if not os.path.exists(image_path):
            error_messages.append(
                f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
                f"image{i+1}がimageフォルダに入っていません。"
            )

      # urlまたはcontactが記入されているか
      if pd.isna(program_df.loc["url", program_name]):
          if pd.isna(program_df.loc["contact", program_name]):
              error_messages.append(f'留学プログラムリストページの「{program_name}」のプログラムで、urlとcontactどちらにも記入があります。URLがある場合はURLのみ、URLがない場合は、contact欄に連絡先を書いてください')

      if not pd.isna(program_df.loc["url", program_name]):
          if not pd.isna(program_df.loc["contact", program_name]):
              error_messages.append(f"留学プログラムリストページの「{program_name}」のプログラムで、URLとcontactどちらも記入されていません。URLまたはcontactどちらかの情報を入れてください")

      # image_nameが正しいか
      if program_df.loc["image_name", program_name] != f"image{i+1}":
        error_messages.append(
          f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
          f"image_nameの名前が間違っています。「image{i+1}」にしてください。"
        )

      # image_nameの写真ががimageフォルダに入ってるか
      base_dir = os.path.dirname(file_path)
      image_path = os.path.join(base_dir, "../image", f"image{i+1}.png")
      print(image_path)

      if not os.path.exists(image_path):
          error_messages.append(
              f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
              f"image{i+1}がimageフォルダに入っていません。"
          )

      #その他の項目が空白でないか
      cols = ["outline", "period", "schedule", "cost", "applicable_grade"]

      nan_cols = program_df.loc[cols, program_name].isna()

      missing = nan_cols[nan_cols].index.tolist()

      if missing:
          error_messages.append(
              f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
              f"{', '.join(missing)}が空白です。"
          )
      program_names_set.add(program_name)

  # 「期間で決める」ページの項目チェック
  period_judge_df = pd.read_excel(file_path,sheet_name=0)
  expected_period_option_count = len(period_judge_df)
  check_valid_program_name(period_judge_df, "期間で決める", start_index=2)

  # 「語学レベルで決める」ページの項目チェック
  language_judge_df = pd.read_excel(file_path,sheet_name=1)
  check_valid_period_option_num(language_judge_df, "語学レベルで決める", expected_period_option_count)
  check_valid_program_name(language_judge_df, df_name="語学レベルで決める", start_index=3)

  # 「次の長期休みにすぐ行きたい」ページの項目チェック
  season_judge_df = pd.read_excel(file_path,sheet_name=2)
  check_valid_period_option_num(season_judge_df, "次の長期休みにすぐ行きたい", expected_period_option_count)
  check_valid_program_name(season_judge_df, "次の長期休みにすぐ行きたい", start_index=3)

  # 「留学目的で選ぶ」ページの項目チェック
  purpose_judge_df = pd.read_excel(file_path,sheet_name=3)
  check_valid_period_option_num(purpose_judge_df, "留学目的で選ぶ", expected_period_option_count)
  check_valid_program_name(purpose_judge_df, "留学目的で選ぶ", start_index=3)

  # 「留学スタイルで選ぶ」ページの項目チェック
  style_judge_df = pd.read_excel(file_path,sheet_name=4)
  check_valid_period_option_num(style_judge_df, "留学目的で選ぶ", expected_period_option_count)
  check_valid_program_name(style_judge_df, "留学目的で選ぶ", start_index=3)

  # 「学内でできること」ページの項目チェック
  style_judge_df = pd.read_excel(file_path,sheet_name=5)
  check_valid_program_name(style_judge_df, "学内でできること", start_index=2)

  return error_messages