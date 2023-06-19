# from github import Github
# import urllib.request
# 設置你的 GitHub 身份驗證信息
# g = Github('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')

# # 指定你的存儲庫名稱和發布版本名稱
# repo_name = 'Yi-Feng-Liu/Excel-Filter-Tool'
# release_name = 'v2.0.0'

# # 獲取存儲庫對象
# repo = g.get_repo(repo_name)

# # 獲取最新的發布版本對象
# release = repo.get_release(release_name)

# # 獲取執行檔的下載鏈接
# download_url = release.get_assets()[0].browser_download_url

# 下載執行檔
# urllib.request.urlretrieve(download_url, 'Widget.rar')

# import sys
# from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QComboBox

# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()

#         self.initUI()

#     def initUI(self):
#         self.setWindowTitle("下拉式選單範例")
#         self.setGeometry(100, 100, 300, 200)

#         # 创建一个标签
#         label = QLabel("選擇:", self)
#         label.move(50, 50)

#         # 创建一个下拉式菜单
#         combobox = QComboBox(self)
#         combobox.addItem("選項1")
#         combobox.addItem("選項2")
#         combobox.addItem("選項3")
#         combobox.move(160, 50)

#         # 设置下拉式菜单的默认选择
#         combobox.setCurrentIndex(-1)

#         # 绑定下拉式菜单的选择变化事件
#         combobox.currentIndexChanged.connect(self.onComboBoxChanged)

#         self.show()

#     def onComboBoxChanged(self, index):
#         # 当下拉式菜单的选择发生变化时触发的事件处理函数
#         selected_value = self.sender().currentText()
#         print(f"選項索引: {selected_value}")

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     window = MainWindow()
#     sys.exit(app.exec())


# import sys
# from PyQt6.QtWidgets import QApplication, QMainWindow, QRadioButton, QLabel

# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Radio 按鈕範例")
        
#         # 建立 Radio 按鈕
#         self.radio_button1 = QRadioButton("選項1", self)
#         self.radio_button1.setChecked(False)  # 預設選取第一個按鈕
#         self.radio_button1.setGeometry(50, 50, 100, 30)
        
#         self.radio_button2 = QRadioButton("選項2", self)
#         self.radio_button2.setGeometry(50, 80, 100, 30)
        
#         # 建立標籤，用來顯示選取的結果
#         self.label = QLabel(self)
#         self.label.setGeometry(50, 120, 200, 30)
        
#         # 連接信號和槽函式
#         self.radio_button1.clicked.connect(self.radio_button_clicked)
#         self.radio_button2.clicked.connect(self.radio_button_clicked)
        
#     def radio_button_clicked(self):
#         selected_option = ""
#         if self.radio_button1.isChecked():
#             selected_option = "選項1"
#         elif self.radio_button2.isChecked():
#             selected_option = "選項2"
        
#         self.label.setText(f"選取結果：{selected_option}")


# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = MainWindow()
#     window.setGeometry(300, 300, 300, 200)
#     window.show()
#     sys.exit(app.exec())

