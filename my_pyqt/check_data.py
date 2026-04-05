import pandas as pd
import json
import os

def check_data(excel_file_path):

  program_df = pd.read_excel(excel_file_path,sheet_name=6)

  program_df = program_df.set_index(program_df.columns[0])

  # プログラムリストにあるプログラム
  program_names_set = set()

  # 選択肢にあるプログラム
  exsisting_programs_set = set()

  error_messages = []

  # プログラム名が有効か
  def check_valid_program_name(judge_df, df_name, start_index):
    for i in range(len(judge_df)):
      program_list = judge_df.iloc[i, start_index:].dropna()
      for index, program_name in enumerate(program_list):
        if program_name not in program_names_set:
          error_messages.append(
                f"「{df_name}」ページの {i+2}行目, {chr(ord('A') + start_index + index)}列目の「{program_name}」のプログラムはプログラムリストに入っていません。"
                "「プログラムリスト」ページのプログラム名と一致させてください。"
            )

        exsisting_programs_set.add(program_name)


  # 期間の選択肢の数が有効か
  # 期間の選択肢の値が有効か
  def check_valid_period_option(judge_df, df_name, expected_period_option_count, expected_period_option_values):
    i = 0
    while i < len(judge_df):
        block = judge_df.iloc[i:i+expected_period_option_count]
        if len(block) != expected_period_option_count:
          error_messages.append(
              f" 「{df_name}」ページの{i+1}行目の選択肢の数が間違っています。"
                "期間の選択肢の数を、「期間で決める」ページの選択肢の数と揃えてください。"
                "例：期間で決めるページの選択肢が4つあれば、同じ言語レベルの選択肢を4つにしてください"
          )
          break
        if (block.iloc[:, 1] != block.iloc[0, 1]).any():
          error_messages.append(
                f" 「{df_name}」ページの{i+1}行目の選択肢の数が間違っています。"
                "期間の選択肢の数を、「期間で決める」ページの選択肢の数と揃えてください。"
                "例：期間で決めるページの選択肢が4つあったら、同じ言語レベルの選択肢を4つにしてください"
            )
          break
        i += expected_period_option_count

    for i, period_name in enumerate(judge_df["期間"]):
      expected_period = expected_period_option_values[i % expected_period_option_count]
      if period_name != expected_period:
        error_messages.append(
            f"「{df_name}」ページのの{i+2}行目の期間の名前が"
            f"「期間を決める」ページの「{expected_period}」と一致してません。"
        )


  # 留学プログラムリストの記入のチェック
  for i, program_name in enumerate(program_df.columns):
      col_data = program_df[program_name]

      # 空のプログラムはスキップ
      if col_data.isna().all():
        base_dir = os.path.dirname(excel_file_path)
        image_path = os.path.join(base_dir, "../image", f"image{i+1}.png")

        if os.path.exists(image_path):
          error_messages.append(f"削除されたプログラムの写真、image{i+1}がまだimageフォルダにあります。削除してください")
          return error_messages
        continue

      #　空のプログラムでない場合
      base_dir = os.path.dirname(excel_file_path)
      image_path = os.path.join(base_dir, "../image", f"image{i+1}.png")

      if not os.path.exists(image_path):
          error_messages.append(
              f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
              f"image{i+1}がimageフォルダに入っていません。"
          )
          return error_messages

      # urlまたはcontactが記入されているか
      if pd.isna(program_df.loc["url", program_name]):
          if pd.isna(program_df.loc["contact", program_name]):
              error_messages.append(f'留学プログラムリストページの「{program_name}」のプログラムで、urlとcontactどちらにも記入があります。URLがある場合はURLのみ、URLがない場合は、contact欄に連絡先を書いてください')
              return error_messages

      if not pd.isna(program_df.loc["url", program_name]):
          if not pd.isna(program_df.loc["contact", program_name]):
              error_messages.append(f"留学プログラムリストページの「{program_name}」のプログラムで、URLとcontactどちらも記入されていません。URLまたはcontactどちらかの情報を入れてください")
              return error_messages

      # image_nameが正しいか
      if program_df.loc["image_name", program_name] != f"image{i+1}":
        error_messages.append(
          f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
          f"image_nameの名前が間違っています。「image{i+1}」にしてください。"
        )
        return error_messages

      # image_nameの写真ががimageフォルダに入ってるか
      base_dir = os.path.dirname(excel_file_path)
      image_path = os.path.join(base_dir, "../image", f"image{i+1}.png")

      if not os.path.exists(image_path):
          error_messages.append(
              f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
              f"image{i+1}がimageフォルダに入っていません。"
          )
          return error_messages

      #その他の項目が空白でないか
      cols = ["outline", "period", "schedule", "cost", "applicable_grade"]

      nan_cols = program_df.loc[cols, program_name].isna()

      missing = nan_cols[nan_cols].index.tolist()

      if missing:
          error_messages.append(
              f"「留学プログラムリスト」ページの「{program_name}」のプログラムの、"
              f"{', '.join(missing)}が空白です。"
          )
          return error_messages
      program_names_set.add(program_name)

  # 「期間で決める」ページの項目チェック
  period_judge_df = pd.read_excel(excel_file_path,sheet_name=0)
  expected_period_option_count = len(period_judge_df)
  expected_period_option_values = period_judge_df["期間"].tolist()
  check_valid_program_name(period_judge_df, "期間で決める", start_index=2)

  # 「語学レベルで決める」ページの項目チェック
  language_judge_df = pd.read_excel(excel_file_path,sheet_name=1)
  check_valid_period_option(language_judge_df, "語学レベルで決める", expected_period_option_count, expected_period_option_values)
  check_valid_program_name(language_judge_df, df_name="語学レベルで決める", start_index=3)

  # 「次の長期休みにすぐ行きたい」ページの項目チェック
  season_judge_df = pd.read_excel(excel_file_path,sheet_name=2)
  check_valid_period_option(season_judge_df, "次の長期休みにすぐ行きたい", expected_period_option_count, expected_period_option_values)
  check_valid_program_name(season_judge_df, "次の長期休みにすぐ行きたい", start_index=3)

  # 「留学目的で選ぶ」ページの項目チェック
  purpose_judge_df = pd.read_excel(excel_file_path,sheet_name=3)
  check_valid_period_option(purpose_judge_df, "留学目的で選ぶ", expected_period_option_count, expected_period_option_values)
  check_valid_program_name(purpose_judge_df, "留学目的で選ぶ", start_index=3)

  # 「留学スタイルで選ぶ」ページの項目チェック
  style_judge_df = pd.read_excel(excel_file_path,sheet_name=4)
  check_valid_period_option(style_judge_df, "留学スタイルで選ぶ", expected_period_option_count, expected_period_option_values)
  check_valid_program_name(style_judge_df, "留学スタイルで選ぶ", start_index=3)

  # 「学内でできること」ページの項目チェック
  style_judge_df = pd.read_excel(excel_file_path,sheet_name=5)
  check_valid_program_name(style_judge_df, "学内でできること", start_index=2)

  # プログラムリストの中で選択肢に出現しなかった場合
  not_exsisting_program_set = program_names_set - exsisting_programs_set
  if len(not_exsisting_program_set) > 0:
    error_messages.append(f"{not_exsisting_program_set}はプログラムリストにありますが、選択肢にありません。選択肢に追加するか、プログラムリストから削除してください。")

  return error_messages