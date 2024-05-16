import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, 
    QFileDialog, QMessageBox, QTextEdit
)

class ExcelSplitter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.excel_label = QLabel('엑셀 파일 경로:')
        self.excel_input = QLineEdit(self)
        self.excel_button = QPushButton('엑셀 파일 선택', self)
        self.excel_button.clicked.connect(self.select_excel_file)

        self.count_label = QLabel('그룹 개수:')
        self.count_input = QLineEdit(self)

        self.display_button = QPushButton('데이터 표시', self)
        self.display_button.clicked.connect(self.display_data)

        self.submit_button = QPushButton('그룹 나누기', self)
        self.submit_button.clicked.connect(self.split_excel)

        self.result_display = QTextEdit(self)

        self.group_labels = []  # 각 그룹을 나타내는 QLabel을 저장할 리스트

        layout.addWidget(self.excel_label)
        layout.addWidget(self.excel_input)
        layout.addWidget(self.excel_button)
        layout.addWidget(self.count_label)
        layout.addWidget(self.count_input)
        layout.addWidget(self.display_button)
        layout.addWidget(self.submit_button)
        layout.addWidget(self.result_display)

        # 각 그룹을 나타내는 QLabel을 추가
        for i in range(10):  # 최대 10개의 그룹을 표시할 수 있도록 함
            label = QLabel(f'그룹 {i+1}:')
            self.group_labels.append(label)
            layout.addWidget(label)

        self.setLayout(layout)
        self.setWindowTitle('엑셀 인원 나누기')
        self.show()

    def select_excel_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', '', 'Excel Files (*.xlsx)', options=options)
        if file_path:
            self.excel_input.setText(file_path)

    def display_data(self):
        excel_path = self.excel_input.text()
        try:
            df = pd.read_excel(excel_path, usecols=[0])  # A열만 읽기
        except Exception as e:
            QMessageBox.warning(self, '파일 오류', f'엑셀 파일을 읽는 중 오류가 발생했습니다: {e}')
            return

        data_text = '\n'.join(df.iloc[:, 0].astype(str).tolist())
        self.result_display.setPlainText(data_text)
        QMessageBox.information(self, '완료', '데이터를 표시했습니다.')

    def split_excel(self):
        try:
            group_count = int(self.count_input.text())
        except ValueError:
            QMessageBox.warning(self, '입력 오류', '올바른 숫자를 입력하세요.')
            return

        data_text = self.result_display.toPlainText()
        data_list = data_text.split('\n')
        df = pd.DataFrame(data_list, columns=['Name'])

        try:
            self.split_and_display(df, group_count)
        except Exception as e:
            QMessageBox.warning(self, '처리 오류', f'인원 나누는 중 오류가 발생했습니다: {e}')

    def split_and_display(self, df, group_count):
        if group_count <= 0:
            QMessageBox.warning(self, '입력 오류', '그룹 개수는 0보다 커야 합니다.')
            return

        # 인원을 랜덤으로 섞기
        df = df.sample(frac=1).reset_index(drop=True)

        # 각 그룹의 크기 계산
        group_size = len(df) // group_count
        remainder = len(df) % group_count

        groups = []
        start_idx = 0

        for i in range(group_count):
            end_idx = start_idx + group_size + (1 if i < remainder else 0)
            group_df = df.iloc[start_idx:end_idx].copy()  # 그룹 DataFrame 복사

            groups.append(group_df)
            start_idx = end_idx

        # 각 그룹의 데이터를 해당하는 QLabel에 표시
        for i, group_df in enumerate(groups):
            self.group_labels[i].setText(f'그룹 {i + 1}:\n{group_df["Name"].to_string(index=False, header=False)}\n')

        QMessageBox.information(self, '완료', '인원을 나누어 표시했습니다.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelSplitter()
    sys.exit(app.exec_())
