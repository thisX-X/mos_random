import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, 
    QFileDialog, QMessageBox, QTextEdit, QGridLayout
)

class ExcelSplitter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        hbox = QHBoxLayout()  # 수평 상자 레이아웃

        self.excel_label = QLabel('엑셀 파일 경로:')
        self.excel_input = QLineEdit(self)
        self.excel_button = QPushButton('엑셀 파일 선택', self)
        self.excel_button.setShortcut('Ctrl+O')  # 단축키 설정
        self.excel_button.clicked.connect(self.select_excel_file)

        self.count_label = QLabel('그룹 개수:')
        self.count_input = QLineEdit(self)
        self.count_input.textChanged.connect(self.update_leader_inputs)

        self.display_button = QPushButton('데이터 표시', self)
        self.display_button.setShortcut('Ctrl+D')  # 단축키 설정
        self.display_button.clicked.connect(self.display_data)

        self.submit_button = QPushButton('그룹 나누기', self)
        self.submit_button.setShortcut('Ctrl+S')  # 단축키 설정
        self.submit_button.clicked.connect(self.split_excel)

        self.result_display = QTextEdit(self)
        
        self.leader_layout = QGridLayout()
        self.leader_inputs = []

        layout.addWidget(self.excel_label)
        layout.addWidget(self.excel_input)
        layout.addWidget(self.excel_button)
        layout.addWidget(self.count_label)
        layout.addWidget(self.count_input)
        layout.addWidget(self.display_button)
        layout.addWidget(self.submit_button)
        layout.addWidget(QLabel('그룹 조장 입력:'))
        layout.addLayout(self.leader_layout)
        layout.addWidget(self.result_display)

        hbox.addLayout(layout)  # 수직 레이아웃을 수평 상자 레이아웃에 추가

        self.setLayout(hbox)  # 전체 레이아웃을 수평 상자 레이아웃으로 설정
        self.setWindowTitle('엑셀 인원 나누기')

        self.group_labels = []  # group_labels 속성 초기화
        self.show()

    def update_leader_inputs(self):
        # 기존 조장 입력란 제거
        for leader_input in self.leader_inputs:
            self.leader_layout.removeWidget(leader_input)
            leader_input.deleteLater()
        self.leader_inputs = []

        try:
            group_count = int(self.count_input.text())
        except ValueError:
            return

        for i in range(group_count):
            leader_input = QLineEdit(self)
            leader_input.setPlaceholderText(f'그룹 {i + 1} 조장')
            self.leader_layout.addWidget(leader_input, i // 2, i % 2)
            self.leader_inputs.append(leader_input)

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

        leaders = [input.text().strip() for input in self.leader_inputs if input.text().strip()]
        if len(leaders) != group_count:
            QMessageBox.warning(self, '입력 오류', '모든 그룹에 대해 조장을 입력하세요.')
            return

        for leader in leaders:
            if leader not in df['Name'].values:
                QMessageBox.warning(self, '입력 오류', f'조장 {leader}이(가) 명단에 없습니다.')
                return

        all_names = pd.concat([df, pd.DataFrame(leaders, columns=['Name'])])
        all_names = all_names.sample(frac=1).reset_index(drop=True)

        groups = []
        start_idx = 0
        for i in range(group_count):
            group = [leaders[i]] 
            end_idx = start_idx + (len(all_names) // group_count) + (1 if i < len(all_names) % group_count else 0) - 1
            group.extend(all_names.iloc[start_idx:end_idx]['Name'].tolist())
            groups.append(group)
            start_idx = end_idx

        for i, group in enumerate(groups):
            group_df = pd.DataFrame(group, columns=['Name'])
            if i < len(self.group_labels):
                self.group_labels[i].setText(f'그룹 {i + 1} (조장: {leaders[i]}):\n{group_df["Name"].to_string(index=False, header=False)}\n')
            else:
                new_label = QLabel(f'그룹 {i + 1} (조장: {leaders[i]}):\n{group_df["Name"].to_string(index=False, header=False)}\n')
                self.group_labels.append(new_label)
                self.layout().addWidget(new_label)

        QMessageBox.information(self, '완료', '인원을 나누어 표시했습니다.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelSplitter()
    sys.exit(app.exec_())
