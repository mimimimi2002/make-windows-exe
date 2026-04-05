import json
import re

def check_json(judge_data, option_data, option_count):
  errors = []

  # judge jsonのkeyの整合性
  expected_keys = set([
      'program_name', 'url', 'outline', 'period',
      'schedule', 'cost', 'application_grade',
      'contact', 'image_name'
  ])

  for key in judge_data.keys():
    for program in judge_data[key]:
      if set(program.keys()) != expected_keys:
        errors.append(f"{program}データのkeyが{expected_keys}と一致していません。")

  for key in judge_data.keys():
    # パターン抽出
    matches = re.findall(r'(PP|Style|P|L|S|E)(\d+)', key)

    for prefix, num in matches:
        num = int(num)
        if prefix in option_count:
            if num > option_count[prefix]:
                errors.append(f"{key}: {prefix}{num} exceeds {option_count[prefix]}")
        else:
            errors.append(f"{key}: unknown prefix {prefix}")

  return errors