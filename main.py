# PyQt5 기반 GUI: 사용자에게 이름 입력 + CSV 파일 선택 + 결과 엑셀 저장

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QLabel,
    QMessageBox,
    QLineEdit,
    QTextEdit,
)
from PyQt5.QtCore import Qt
import pandas as pd
from datetime import datetime, date
import sys
import os
import re


class ScheduleApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("근무시간 병합기")
        self.setGeometry(100, 100, 1000, 800)
        self.setAcceptDrops(True)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(40, 20, 40, 20)
        main_layout.setSpacing(20)

        # 근무자 이름
        name_label = QLabel("근무자 이름:")
        name_label.setStyleSheet(
            "font-size: 30px; font-weight: bold; margin-bottom: 2px;"
        )
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("예: 권혁준")
        self.name_input.setStyleSheet("font-size: 20px; padding: 6px;")
        self.name_input.setFixedHeight(70)

        # 파일 선택
        file_label = QLabel("CSV 파일 경로:")
        file_label.setStyleSheet(
            "font-size: 30px; font-weight: bold; margin-top: 10px; margin-bottom: 2px;"
        )

        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        self.file_path.setPlaceholderText("드래그하거나 '찾아보기' 버튼 클릭")
        self.file_path.setStyleSheet("font-size: 20px; padding: 6px;")
        self.file_path.setFixedHeight(120)

        file_button = QPushButton("찾아보기")
        file_button.setFixedWidth(120)
        file_button.setFixedHeight(70)
        file_button.setStyleSheet("font-size: 20px; padding: 8px; font-weight: bold")
        file_button.clicked.connect(self.select_file)

        # 파일 경로와 버튼을 가로로 배치
        file_row = QHBoxLayout()
        file_row.addWidget(self.file_path)
        file_row.addWidget(file_button)

        # 실행 버튼
        self.run_btn = QPushButton("엑셀로 저장")
        self.run_btn.setStyleSheet("font-size: 18px; padding: 12px; font-weight: bold;")
        self.run_btn.setFixedWidth(200)
        self.run_btn.clicked.connect(self.run_process)

        # 버튼 가운데 정렬
        button_row = QHBoxLayout()
        button_row.addStretch()
        button_row.addWidget(self.run_btn)
        button_row.addStretch()

        # 전체 배치
        main_layout.addWidget(name_label)
        main_layout.addWidget(self.name_input)
        main_layout.addWidget(file_label)
        main_layout.addLayout(file_row)
        main_layout.addSpacing(40)
        main_layout.addLayout(button_row)

        self.setLayout(main_layout)

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "CSV 파일 선택", "", "CSV Files (*.csv)"
        )
        if path:
            self.file_path.setText(path)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(".csv"):
                self.file_path.setText(file_path)
            else:
                QMessageBox.warning(self, "형식 오류", "CSV 파일만 허용됩니다.")

    def run_process(self):
        name = self.name_input.text().strip()
        path = self.file_path.text().strip()

        if not name:
            QMessageBox.warning(self, "오류", "근무자 이름을 입력하세요.")
            return
        if not path or not path.endswith(".csv"):
            QMessageBox.warning(self, "오류", "유효한 CSV 파일을 선택하세요.")
            return

        try:
            self.process_csv(path, name)
            QMessageBox.information(self, "완료", f"{name}.xlsx 파일이 생성되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류 발생", str(e))

    def process_csv(self, file_path, target_name):
        df = pd.read_csv(file_path)
        df.dropna(subset=["Unnamed: 0", "Unnamed: 1"], how="all", inplace=True)
        df.rename(columns={"Unnamed: 0": "날짜", "Unnamed: 1": "근무자"}, inplace=True)
        df["날짜"] = df["날짜"].fillna(method="ffill")

        mask = df.applymap(lambda x: target_name in str(x).strip())
        positions = mask.stack()[mask.stack()].index.tolist()

        results = [(df.at[row, "날짜"], time) for row, time in positions]

        date_order = []
        for d, _ in results:
            if d not in date_order:
                date_order.append(d)

        grouped = {}
        for d, t in results:
            grouped.setdefault(d, []).append(t)

        def sort_key(t):
            start = t.split("~")[0].strip()
            return datetime.strptime(start, "%H:%M")

        for d in grouped:
            grouped[d].sort(key=sort_key)

        def merge_time_ranges(time_ranges):
            def to_range(t):
                start_str, end_str = [
                    s.strip() for s in re.sub(r"\s*~\s*", "-", t).split("-")
                ]
                return (
                    datetime.strptime(start_str, "%H:%M"),
                    datetime.strptime(end_str, "%H:%M"),
                )

            ranges = [to_range(t) for t in time_ranges]
            ranges.sort()

            merged = []
            current_start, current_end = ranges[0]

            for start, end in ranges[1:]:
                if start == current_end:
                    current_end = end
                else:
                    merged.append(
                        f"{current_start.strftime('%H:%M')}-{current_end.strftime('%H:%M')}"
                    )
                    current_start, current_end = start, end
            merged.append(
                f"{current_start.strftime('%H:%M')}-{current_end.strftime('%H:%M')}"
            )

            return merged

        start_year = 2025
        start_month = 4
        weekday_kor = [
            "월요일",
            "화요일",
            "수요일",
            "목요일",
            "금요일",
            "토요일",
            "일요일",
        ]
        cur_year, cur_month, prev_day = start_year, start_month, 0

        records = []
        for d_str in date_order:
            day_int = int(d_str.replace("일", ""))
            if prev_day and day_int < prev_day:
                cur_month += 1
                if cur_month > 12:
                    cur_month = 1
                    cur_year += 1
            prev_day = day_int

            real_date = date(cur_year, cur_month, day_int)
            weekday = weekday_kor[real_date.weekday()]
            merged_times = merge_time_ranges(grouped[d_str])
            times = ",".join(merged_times)

            records.append(
                {"월": cur_month, "일": day_int, "요일": weekday, "근무시간": times}
            )

        output_df = pd.DataFrame(records, columns=["월", "일", "요일", "근무시간"])
        output_df.to_excel(f"{target_name}.xlsx", index=False)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = ScheduleApp()
    window.show()
    sys.exit(app.exec_())
