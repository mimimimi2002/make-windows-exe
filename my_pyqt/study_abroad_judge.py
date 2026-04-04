import sys
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QMessageBox
from data_read import data_read
from check_data import check_data
import json
import os
import shutil
from datetime import datetime
from check_json import check_json

class UploadApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('ファイルアップロード')
        self.layout = QVBoxLayout()

        # 1. アップロードボタンの作成
        self.btn_upload = QPushButton('ファイルを選択', self)
        self.btn_upload.clicked.connect(self.open_file_dialog)
        self.layout.addWidget(self.btn_upload)

        # ファイル名表示用ラベル
        self.label_file = QLabel('ファイルが選択されていません', self)
        self.layout.addWidget(self.label_file)

        self.setLayout(self.layout)
        self.msg = QMessageBox()

    # 2. クリック時の処理（ファイルダイアログ）
    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "ファイルを開く", "", "All Files (*);;Text Files (*.txt)"
        )

        try:
          if file_name:
              self.label_file.setText(f'選択: {file_name}')
              print(f'アップロード予定ファイル: {file_name}')
              error_messages = check_data(file_name)

              if len(error_messages) > 0:
                  self.show_error(error_messages)
                  return

              judge_data, option_data, option_count = data_read(file_name)

              error_messages = check_json(judge_data, option_data, option_count)

              if len(error_messages) > 0:
                  self.show_error(error_messages)
                  return

              save_judge_dir = os.path.join(os.path.expanduser("~"), "Downloads", "data", "judge_data")
              save_image_dir = os.path.join(os.path.expanduser("~"), "Downloads", "data", "image")

              os.makedirs(save_judge_dir, exist_ok=True)
              os.makedirs(save_image_dir, exist_ok=True)

              file_basename = os.path.basename(file_name)
              save_path = os.path.join(save_judge_dir, file_basename)
              shutil.copy(file_name, save_path)

              with open(f"{save_judge_dir}/judge_data.json", "w", encoding="utf-8") as f:
                  json.dump(judge_data, f, ensure_ascii=False, indent=4)

              with open(f"{save_judge_dir}/option_data.json", "w", encoding="utf-8") as f:
                  json.dump(option_data, f, ensure_ascii=False, indent=4)

              with open(f"{save_judge_dir}/option_count.json", "w", encoding="utf-8") as f:
                  json.dump(option_count, f, ensure_ascii=False, indent=4)

              # save images
              image_dir = os.path.join(os.path.dirname(file_name), "../image")
              if os.path.exists(save_image_dir):
                shutil.rmtree(save_image_dir)
              shutil.copytree(image_dir, save_image_dir)

              date_str = datetime.now().strftime("%Y-%m-%d-%H-%M")

              update_file_path = os.path.join(save_judge_dir, "update_at.txt")

              with open(update_file_path, "w", encoding="utf-8") as f:
                  f.write(date_str)

              self.msg.setWindowTitle("完了")
              save_dir = os.path.join(os.path.expanduser("~"), "Downloads", "data")
              self.msg.setText(f"保存が完了しました！\n{save_dir}")
              self.msg.exec()

              self.close()
        except Exception as e:
          QMessageBox.critical(self, "エラー", str(e))

    def show_error(self, error_messages):
        if len(error_messages) > 5:
            error_text = "\n".join(error_messages[:5])
        else:
            error_text = "\n".join(error_messages)
        self.msg.setWindowTitle("エラー")
        self.msg.setText(f"エクセルデータにエラーがあります\n\n{error_text}")
        self.msg.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = UploadApp()
    window.show()
    sys.exit(app.exec())
