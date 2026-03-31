import sys
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QMessageBox
from data_read import data_read
import json
import os

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

    # 2. クリック時の処理（ファイルダイアログ）
    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "ファイルを開く", "", "All Files (*);;Text Files (*.txt)"
        )

        try:
          if file_name:
              self.label_file.setText(f'選択: {file_name}')
              print(f'アップロード予定ファイル: {file_name}')
              judge_data, option_data, option_count = data_read(file_name)

              save_dir = os.path.join(os.path.expanduser("~"), "Downloads", "data")

              os.makedirs(save_dir, exist_ok=True)

              if save_dir:
                with open(f"{save_dir}/judge_data.json", "w", encoding="utf-8") as f:
                  json.dump(judge_data, f, ensure_ascii=False, indent=4)

                with open(f"{save_dir}/option_data.json", "w", encoding="utf-8") as f:
                    json.dump(option_data, f, ensure_ascii=False, indent=4)

                with open(f"{save_dir}/option_count.json", "w", encoding="utf-8") as f:
                    json.dump(option_count, f, ensure_ascii=False, indent=4)

                msg = QMessageBox()
                msg.setWindowTitle("完了")
                msg.setText(f"保存が完了しました！\n{save_dir}")
                msg.exec()

                self.close()
        except Exception as e:
          QMessageBox.critical(self, "エラー", str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = UploadApp()
    window.show()
    sys.exit(app.exec())
