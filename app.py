import sys
import os
import time
import tempfile
from pathlib import Path
from datetime import datetime
from typing import Generator
from decimal import Decimal
import sqlite3
# import pyodbc

# PyQt5
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QColor, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QMessageBox, QFileDialog, QTableWidgetItem
)
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QCompleter
# from PyQt5.QtMultimedia import QSound
from PyQt5.uic import loadUiType
# from PyQt5.QtWebEngineWidgets import QWebEngineView

# SQL & Database
import pyodbc
from sqlalchemy import create_engine, VARCHAR, NVARCHAR, INTEGER, DATE, DECIMAL
from sqlalchemy.engine import URL, Engine
from sqlalchemy.orm import sessionmaker, declarative_base

# Excel & Pandas
import pandas as pd
from pandas import DataFrame
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Load UI
import qdarkstyle
import resources_rc

def get_resource_path(relative_path):
    """Trả về đường dẫn đầy đủ đến tài nguyên."""
    if getattr(sys, 'frozen', False):  # Kiểm tra nếu đang chạy file .exe
        base_path = sys._MEIPASS
    else:  # Nếu đang chạy bằng Python gốc
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def connect_to_db(): 
    try:
        db_file = r"D:\nqhuy\VA\ICOP\IC.sqlite"
        # password = "huyie"
        connection = sqlite3.connect(db_file)
        return connection

        # conn_str = (
        #     r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        #     rf'DBQ={db_file};'
        #     rf'PWD={password};'
        # )

        # conn = pyodbc.connect(conn_str)
        # return conn
    except pyodbc.Error as e:
        print(f"Lỗi khi kết nối tới cơ sở dữ liệu: {e}")
        return None

def table_to_dataframe(table_widget,headers):
        rows = table_widget.rowCount()
        columns = table_widget.columnCount()
        
        # Lấy tiêu đề cột
        # headers = [table_widget.horizontalHeaderItem(i).text() for i in range(columns)]
        
        # Lấy dữ liệu từ bảng
        data = []
        for row in range(rows):
            row_data = []
            for column in range(columns):
                item = table_widget.item(row, column)
                row_data.append(item.text() if item else '')  # Lấy text từ ô, nếu không có thì gán chuỗi rỗng
            data.append(row_data)
        
        # Tạo DataFrame
        df = pd.DataFrame(data, columns=headers)
        return df
    
ui, _ = loadUiType(get_resource_path('app.ui'))


#config URL cho engine
# BASE_DIR = Path(__file__).resolve().parent
# env_file = get_resource_path(".env")
# load_dotenv(env_file)

# class Settings():
#     API_PREFIX = ''
#     DATABASE_1_URL = URL.create(
#         "mssql+pyodbc",
#         username=os.getenv("UID"),
#         password=os.getenv("PASSWORD"),
#         host=os.getenv("SERVER"),
#         port=1433,
#         database=os.getenv("DB"),
#         query={
#            "driver": "ODBC Driver 17 for SQL Server",
#            "TrustServerCertificate": "yes" 
#         }
#     )

# settings = Settings()
# #tạo engine để kêt nối database
# engine_1 = create_engine(settings.DATABASE_1_URL, pool_pre_ping=True)
# SessionLocal_1 = sessionmaker(autocommit=False, autoflush=False, bind=engine_1)

# Base = declarative_base()

# def get_db_1() -> Generator:
#     try:
#         db = SessionLocal_1()
#         yield db
#     finally:
#         db.close()
# #hàm import to sql       
# def import_to_sql(df: DataFrame, table_name: str, dtype: dict, engine: Engine):
#     # Show processing message
#     processing_message = QMessageBox()
#     processing_message.setWindowTitle("Đang xử lý")
#     processing_message.setText("Đang xử lý dữ liệu, vui lòng chờ...")
#     # processing_message.setStandardButtons(QMessageBox.NoButton)
#     processing_message.setModal(True)
#     processing_message.show()
#     QApplication.processEvents()  # Ensure UI updates during processing
#     time.sleep(0.01)  # Simulate processing time
#     try:
#         with engine.connect() as connection:
#             df.to_sql(name=table_name, con=connection, if_exists="append", index=False, dtype=dtype)
            
#         processing_message.close()
#     except Exception as e:
#         processing_message.close()
#         # raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

class MainApp(QMainWindow,ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        
        self.tabWidget.setCurrentIndex(0)
        self.tabWidget.tabBar().setVisible(False)
        self.menuBar.setVisible(False)
        
        # self.toolBar.setVisible(False)
        self.bt001.clicked.connect(self.login)
        self.menu11.triggered.connect(self.show_ETS_tab)
        self.menu13.triggered.connect(self.show_tab_2)
        self.menu14.triggered.connect(self.show_QCO_tab)
        self.menu51.triggered.connect(self.show_login_tab)

        self.de101.setDate(QDate.currentDate().addDays(-365))
        self.de102.setDate(QDate.currentDate())
        # self.de301.setDate(QDate.currentDate().addDays(-365))
        # self.de302.setDate(QDate.currentDate())
        # ####
        # self.bt204.clicked.connect(self.tai_xuong_file_mau_Checker)
        # ####
        self.cb202.currentIndexChanged.connect(self.load_lb202)
        self.cb203.currentIndexChanged.connect(self.load_cb204)
        self.cb204.currentIndexChanged.connect(self.load_lb204)
        self.cb204.currentIndexChanged.connect(self.load_cb205)
        self.cb205.currentIndexChanged.connect(self.load_lb205)
        self.cb205.currentIndexChanged.connect(self.load_cb206)
        self.cb206.currentIndexChanged.connect(self.load_lb206)
        self.bt201.clicked.connect(self.add_du_an)
        self.bt207.clicked.connect(self.add_cong_viec)
        self.rd201.toggled.connect(self.rd201_change)

    def rd201_change(self):
        if self.rd201.isChecked():
            self.cb200.clear()
        else:
            self.load_cb200()

    def msgbox(self, message):
        QtWidgets.QMessageBox.information(self, "Thông báo", message)

    def login(self):
        # fty = self.cb001.currentText()
        un = self.tb001.text()
        pw = self.tb002.text()
        
        connection = connect_to_db()
        if connection is None:
            self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
            return
        cursor = connection.cursor()
        cursor.execute("""
                SELECT Ma_phong_ban, MNV, Ho_ten , Phan_quyen
                FROM DANH_SACH_CBCNV 
                WHERE MNV = ? AND Mat_khau = ?
            """, (un, pw))
        result = cursor.fetchone()
        if connection:
            connection.close()    

        if result:
            self.menuBar.setVisible(True)
            self.tabWidget.setCurrentIndex(1)
            self.lb003.setText("Danh sách công việc")
            self.lb000.setText(result[0])
            self.lb004.setText(result[1])
            self.lb001.setText(result[2])
            # QSound.play(":/sounds/sounds/success.wav") # Phát âm thanh thành công
            #progress bar
            self.progressBar.setValue(0)  # Khởi tạo giá trị là 0
            self.progressBar.setMinimum(0)
            self.progressBar.setMaximum(100)

            # self.tong_so_dong_ETS()
        else:
            # QSound.play(":/sounds/sounds/error.wav") # Phát âm thanh lỗi
            self.lb002.setText("Tài khoản hoặc mật khẩu không đúng!")

    def add_du_an(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_du_an = self.le201.text().strip()
        ten_du_an = self.te201.toPlainText()

        if ma_du_an and ten_du_an:
            try:
                cursor.execute(
                    "INSERT INTO DANH_SACH_DU_AN (ma_du_an, ten_du_an) VALUES (?, ?)",
                    (ma_du_an, ten_du_an)
                )
                connection.commit()
                self.msgbox("✅ Thêm dự án thành công")
            except Exception as e:
                print("Lỗi khi thêm dự án:", e)
                self.msgbox("⚠️ Mã dự án đã tồn tại hoặc có lỗi khác")
        else:
            self.msgbox("⚠️ Vui lòng nhập đầy đủ mã và tên dự án")

        connection.close()
        self.load_cb202()
        self.le201.setText("")
        self.te201.setText("")

    def add_cong_viec(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        mnv = self.lb004.text().strip()
        ten_cv = self.le207.text()

        if mnv and ten_cv:
            try:
                cursor.execute(
                    "INSERT INTO GHI_NHO_TEN_CONG_VIEC (MNV, Ten_cong_viec) VALUES (?, ?)",
                    (mnv, ten_cv)
                )
                connection.commit()
                self.msgbox("✅ Thêm công việc mới thành công")
            except Exception as e:
                print("Lỗi khi thêm công việc mới:", e)
                self.msgbox("⚠️ Tên công việc đã tồn tại hoặc có lỗi khác")
        else:
            self.msgbox("⚠️ Vui lòng nhập đầy đủ tên công việc mới")

        connection.close()
        self.load_cb207()
        self.le207.setText("")
    def load_cb200(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        cursor.execute("SELECT ID FROM DANH_SACH_CONG_VIEC")
        results = cursor.fetchall()

        self.cb200.clear()  # Xóa các item hiện có
        for row in results:
            self.cb200.addItem(str(row[0])) 

        connection.close()
        self.cb200.setEditable(True)
        self.cb200.completer().setFilterMode(QtCore.Qt.MatchContains)  # Gợi ý theo từ khóa chứa
        self.cb200.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.cb200.setStyleSheet("""
            QComboBox {
                background-color: #2e2e2e;
                color: white;
            }
            QComboBox QAbstractItemView {
                background-color: #3c3c3c;
                color: white;
            }
        """)
    def load_cb201(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_phong_ban = self.lb000.text()
        cursor.execute("SELECT DISTINCT NHOM FROM DANH_SACH_CBCNV WHERE Ma_phong_ban = ?",(ma_phong_ban,))
        results = cursor.fetchall()

        self.cb201.clear()  # Xóa các item hiện có
        for row in results:
            self.cb201.addItem(row[0])  # row[0] 

        connection.close()

    def load_cb204(self):
            connection = connect_to_db()
            cursor = connection.cursor()

            ma_phong_ban = self.lb000.text()
            phan_loai = self.cb203.currentText()
            cursor.execute("SELECT DISTINCT Ma_chuc_nang FROM CHUC_NANG_NHIEM_VU WHERE Ma_phong_ban = ? AND phan_loai = ?",(ma_phong_ban,phan_loai))
            results = cursor.fetchall()

            self.cb204.clear()  # Xóa các item hiện có
            for row in results:
                self.cb204.addItem(row[0])  # row[0] 

            connection.close()
    def load_cb205(self):
            connection = connect_to_db()
            cursor = connection.cursor()

            ma_chuc_nang = self.cb204.currentText()
            cursor.execute("SELECT DISTINCT Ma_nhiem_vu FROM CHUC_NANG_NHIEM_VU WHERE Ma_chuc_nang = ?",(ma_chuc_nang,))
            results = cursor.fetchall()

            self.cb205.clear()  # Xóa các item hiện có
            for row in results:
                self.cb205.addItem(row[0])  # row[0] 

            connection.close()   
    def load_cb206(self):
            connection = connect_to_db()
            cursor = connection.cursor()

            ma_nhiem_vu = self.cb205.currentText()
            cursor.execute("SELECT DISTINCT Ma_nhiem_vu_cu_the FROM CHUC_NANG_NHIEM_VU WHERE Ma_nhiem_vu = ?",(ma_nhiem_vu,))
            results = cursor.fetchall()

            self.cb206.clear()  # Xóa các item hiện có
            for row in results:
                self.cb206.addItem(row[0])  # row[0] 

            connection.close() 
    def load_cb207(self):
            connection = connect_to_db()
            cursor = connection.cursor()

            mnv = self.lb004.text()
            cursor.execute("SELECT Ten_cong_viec FROM GHI_NHO_TEN_CONG_VIEC WHERE MNV = ?",(mnv,))
            results = cursor.fetchall()

            self.cb207.clear()  # Xóa các item hiện có
            for row in results:
                self.cb207.addItem(row[0])  # row[0] 

            connection.close() 
            self.cb207.setEditable(True)
            self.cb207.completer().setFilterMode(QtCore.Qt.MatchContains)  # Gợi ý theo từ khóa chứa
            self.cb207.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
            self.cb207.setStyleSheet("""
                QComboBox {
                    background-color: #2e2e2e;
                    color: white;
                }
                QComboBox QAbstractItemView {
                    background-color: #3c3c3c;
                    color: white;
                }
            """)
    def load_cb202(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        cursor.execute("SELECT ma_du_an FROM danh_sach_du_an")
        results = cursor.fetchall()

        self.cb202.clear()  # Xóa các item hiện có
        for row in results:
            self.cb202.addItem(row[0])  # row[0] là 'ten_du_an'

        connection.close()
        self.cb202.setEditable(True)
        self.cb202.completer().setFilterMode(QtCore.Qt.MatchContains)  # Gợi ý theo từ khóa chứa
        self.cb202.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.cb202.setStyleSheet("""
            QComboBox {
                background-color: #2e2e2e;
                color: white;
            }
            QComboBox QAbstractItemView {
                background-color: #3c3c3c;
                color: white;
            }
        """)
    def load_lb202(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_du_an = self.cb202.currentText()
        cursor.execute("SELECT ten_du_an FROM danh_sach_du_an where ma_du_an = ?",(ma_du_an,))
        results = cursor.fetchone()
        if results:
            self.lb202.setText(results[0])
        else:
            self.lb202.setText("")

    def load_lb204(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_chuc_nang = self.cb204.currentText()
        cursor.execute("SELECT Chuc_nang FROM CHUC_NANG_NHIEM_VU where Ma_chuc_nang = ?",(ma_chuc_nang,))
        results = cursor.fetchone()
        if results:
            self.lb204.setText(results[0])
        else:
            self.lb204.setText("")
    def load_lb205(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_nhiem_vu = self.cb205.currentText()
        cursor.execute("SELECT Nhiem_vu FROM CHUC_NANG_NHIEM_VU where Ma_nhiem_vu = ?",(ma_nhiem_vu,))
        results = cursor.fetchone()
        if results:
            self.lb205.setText(results[0])
        else:
            self.lb205.setText("")
    def load_lb206(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_nhiem_vu_cu_the = self.cb206.currentText()
        cursor.execute("SELECT Nhiem_vu_cu_the FROM CHUC_NANG_NHIEM_VU where Ma_nhiem_vu_cu_the = ?",(ma_nhiem_vu_cu_the,))
        results = cursor.fetchone()
        if results:
            self.lb206.setText(results[0])
        else:
            self.lb206.setText("")
    def search_ETS(self):
        nha_may = self.lb000.text()
        tu_ngay = self.de101.date().toString("yyyy-MM-dd")
        den_ngay = self.de102.date().toString("yyyy-MM-dd")
        chuyen = self.tb101.text()
        mnv = self.tb102.text()
        ho_ten = self.tb103.text()
        vi_tri = self.tb104.text()
        
        connection = connect_to_db()
        if connection is None:
            self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
            return
        cursor = connection.cursor()
        sql = f"""
                SELECT LINE,EMP_CODE,EMP_NAME,POSITION,WORK_DATE,MO,OP_NO,QTY,WORKING_TIME,ID
                FROM TNC_INDIVIDUAL_EFF
                WHERE WORK_DATE BETWEEN '{tu_ngay}' AND '{den_ngay}' 
                AND LINE LIKE '%{chuyen}%'
                AND EMP_CODE LIKE '%{mnv}%'
                AND EMP_NAME LIKE '%{ho_ten}%'
                AND POSITION LIKE '%{vi_tri}%'
                AND FAC_CODE = '{nha_may}'
            """
        cursor.execute(sql)
        results = cursor.fetchall()
        # Xóa dữ liệu cũ trong TableWidget
        self.tableWidget.setRowCount(0)
        
        # Hiển thị dữ liệu trong TableWidget
        for row_idx, row_data in enumerate(results):
            self.tableWidget.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.tableWidget.setItem(row_idx, col_idx, item)
            
            # Cập nhật progress bar
            self.progressBar.setValue(int((row_idx + 1) * 100 / len(results)))
        
        self.tableWidget.resizeColumnsToContents()

        # Gọi hàm tổng số dòng
        self.tong_so_dong_ETS()

    def search_Checker(self):
        tu_ngay = self.de201.date().toString("yyyy-MM-dd")
        den_ngay = self.de202.date().toString("yyyy-MM-dd")
        nhom = self.tb201.text()
        may = self.tb202.text()
        style = self.tb203.text()
        mau = self.tb204.text()
        
        connection = connect_to_db()
        if connection is None:
            self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
            return
        cursor = connection.cursor()
        sql = f"""
                SELECT *
                FROM CHECKER_PRT
                WHERE WORK_DATE BETWEEN '{tu_ngay}' AND '{den_ngay}' 
                AND STYLE LIKE '%{style}%'
                AND GROUP_NAME LIKE '%{nhom}%'
                AND MC_NAME LIKE '%{may}%'
                AND COLOR LIKE '%{mau}%'
            """
        cursor.execute(sql)
        results = cursor.fetchall()
        # Xóa dữ liệu cũ trong TableWidget
        self.tableWidget_2.setRowCount(0)
        
        # Hiển thị dữ liệu trong TableWidget
        for row_idx, row_data in enumerate(results):
            self.tableWidget_2.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.tableWidget_2.setItem(row_idx, col_idx, item)
            
            # Cập nhật progress bar
            self.progressBar_2.setValue(int((row_idx + 1) * 100 / len(results)))
        
        self.tableWidget_2.resizeColumnsToContents()

        # Gọi hàm tổng số dòng
        self.tong_so_dong_checker()
    
    def show_login_tab(self):
        self.tabWidget.setCurrentIndex(0) 
        self.tb001.setText("")
        self.tb002.setText("")
        self.lb000.setText("")
        self.lb001.setText("")
        self.lb002.setText("")
        self.lb003.setText("Phần mềm quản lý kho vải")
        self.menuBar.setVisible(False)
                        
    def show_ETS_tab(self):
        self.tabWidget.setCurrentIndex(1) 
        self.lb003.setText("Danh sách công việc")
     
    def show_tab_2(self):
        self.menuBar.setVisible(True)
        self.tabWidget.setCurrentIndex(2) 
        self.lb003.setText("Thêm công việc mới")  
        self.load_cb201()
        self.load_cb202()
        self.load_cb204()
        self.load_cb207()
        self.rd201.setChecked(True)
    
    def show_QCO_tab(self):
        self.menuBar.setVisible(True)
        self.tabWidget.setCurrentIndex(3) 
        self.lb003.setText("Sửa xóa thông tin công việc")  
        
    def tong_so_dong_ETS(self):
        rows = self.tableWidget.rowCount()
        self.lb101.setText(f"Tổng số dòng dữ liệu: {rows}")
        self.lb101.setStyleSheet("color: rgb(0, 255, 0);")
        
    def tong_so_dong_SAM(self):
        rows = self.tableWidget_2.rowCount()
        self.lb201.setText(f"Tổng số dòng dữ liệu: {rows}")
        self.lb201.setStyleSheet("color: rgb(0, 255, 0);")

    def tong_so_dong_QCO(self):
        rows = self.tableWidget_3.rowCount()
        self.lb301.setText(f"Tổng số dòng dữ liệu: {rows}")
        self.lb301.setStyleSheet("color: rgb(0, 255, 0);")
        
    def import_from_excel_ETS(self):
        # Mở hộp thoại để chọn tệp
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_to_open, _ = QFileDialog.getOpenFileName(self, "Chọn tệp Excel", "", "Excel Files (*.xlsx *.xls)", options=options)

        # Kiểm tra nếu người dùng không chọn tệp
        if not file_to_open:
            QMessageBox.information(self, "Thông báo", "Không có tệp nào được chọn!")
            return

        # Đọc tệp Excel
        try:
            usecols = [0,2,3,4,5,6,8,9,16,18]
            df = pd.read_excel(file_to_open,sheet_name="fmRptEmpMakeIE1",usecols=usecols,header=0,skiprows=5)
            df = df.rename(columns={
                'Factory':'FAC_CODE',
                'Prod. Group' : 'LINE',
                'EmpNo.' : 'EMP_CODE',
                'Emp.Name' : 'EMP_NAME',
                'Position' : 'POSITION',
                'Date ' : 'WORK_DATE',
                'MO ' : 'MO',
                'Op.No.' : 'OP_NO',
                'Qty ' : 'QTY',
                'Time' : 'WORKING_TIME'
            })
            # print(df)
            df['WORK_DATE'] = pd.to_datetime(df['WORK_DATE'], errors='coerce').dt.strftime('%Y-%m-%d')
            df = df.drop_duplicates(subset=['FAC_CODE','EMP_CODE','WORK_DATE','MO','OP_NO'])        
            dtype = {
                'FAC_CODE' : VARCHAR(10),
                'LINE' : VARCHAR(10),
                'EMP_CODE' : NVARCHAR(10),
                'EMP_NAME' : VARCHAR(50),
                'POSITION' : VARCHAR(100),
                'WORK_DATE' : DATE,
                'MO' :  VARCHAR(30),
                'OP_NO' :  VARCHAR(10),
                'QTY' : INTEGER,
                'WORKING_TIME' : DECIMAL(6,2)
            }
            # print(df)
            import_to_sql(df,'TNC_INDIVIDUAL_EFF',dtype,engine_1)
            QMessageBox.information(self, "Thông báo", f"Tải lên thành công {df.shape[0]} dòng dữ liệu!")
            self.search_ETS()
            self.exec_CALCULATE_TNC_EFF()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi đọc tệp Excel: {e}")    

    def import_from_excel_Checker(self):
        # Mở hộp thoại để chọn tệp
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_to_open, _ = QFileDialog.getOpenFileName(self, "Chọn tệp Excel", "", "Excel Files (*.xlsx *.xls)", options=options)

        # Kiểm tra nếu người dùng không chọn tệp
        if not file_to_open:
            QMessageBox.information(self, "Thông báo", "Không có tệp nào được chọn!")
            return

        # Đọc tệp Excel
        try:
            df = pd.read_excel(file_to_open,sheet_name="Checker",usecols=range(31),header=0,skiprows=1)
            df = df.rename(columns={
                'Ngày':'Work_date',
                'Tuần' : 'Week_num',
                'Nhóm' : 'Group_name',
                'Máy' : 'MC_name',
                'Màu' : 'Color',
                'Số kiểm' : 'Lot_size',
                'Mục tiêu' : 'Target_AQL',
                'In sai hình' : 'Wrong_graphic',
                'Lỗi màu in' : 'Printing_color',
                'Rỗ mặt' : 'Pin_hole',
                'Lệch khung' : 'Off_screen',
                'Tắc khung' : 'Screen_blocking',
                'Thủng khung' : 'Screen_broken',
                'Thấm mặt sau' : 'Penetration',
                'Xếp ly' : 'Pleated',
                'Bong tróc' : 'Peel_off',
                'Rạn nứt' : 'Cracked',
                'Chi tiết in' : 'Printing_detail',
                'Bề mặt xấu' : 'Printing_effect',
                'Lệch tâm' : 'Off_centre',
                'Thông số' : 'Measurement_issue',
                'Sai vị trí' : 'Wrong_position',
                'Cong,méo' : 'Slanting',
                'Lem mực' : 'Smear_ink',
                'Bẩn' : 'Dirty',
                'Cục vải,bàn' : 'Lumpy',
                'Khác' : 'Other',
                'Lỗi vải' : 'Fabric_issue',
                'Lỗi chết' : 'Critical_defect'
            })
            df['Work_date'] = pd.to_datetime(df['Work_date'], errors='coerce').dt.strftime('%Y-%m-%d')
            df = df.drop_duplicates(subset=['Work_date','MC_name','Style','Color'])
            
            dtype = {
                'Work_date' : DATE,
                'Week_num' : VARCHAR(10),
                'Group_name' : NVARCHAR(20),
                'MC_name' : VARCHAR(10),
                'Style' : VARCHAR(30),
                'Color' : VARCHAR(30),
                'Lot_size' : INTEGER,
                'Target_AQL' : DECIMAL(3,1),
                'Wrong_graphic' : INTEGER,
                'Printing_color' : INTEGER,
                'Pin_hole' : INTEGER,
                'Off_screen' : INTEGER,
                'Screen_blocking' : INTEGER,
                'Screen_broken' : INTEGER,
                'Penetration' : INTEGER,
                'Pleated' : INTEGER,
                'Peel_off' : INTEGER,
                'Cracked' : INTEGER,
                'Printing_detail' : INTEGER,
                'Printing_effect' : INTEGER,
                'Off_centre' : INTEGER,
                'Measurement_issue' : INTEGER,
                'Wrong_position' : INTEGER,
                'Slanting' : INTEGER,
                'Smear_ink' : INTEGER,
                'Dirty' : INTEGER,
                'Lumpy' : INTEGER,
                'Other' : INTEGER,
                'Fabric_issue' : INTEGER,
                'NG' : INTEGER,
                'Critical_defect' : INTEGER,

            }
            import_to_sql(df,'CHECKER_PRT',dtype,engine_1)
            QMessageBox.information(self, "Thông báo", f"Tải lên thành công {df.shape[0]} dòng dữ liệu!")
            self.search_Checker()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi đọc tệp Excel: {e}")        

    def delete_selected_rows_QA(self):
        # Lấy ID danh sách các hàng được chọn
        selected_IDs  = set(self.tableWidget.item(index.row(),31).text() 
                          for index in self.tableWidget.selectedIndexes()
                          if self.tableWidget.item(index.row(),31) is not None)
        # Kiểm tra nếu không có hàng nào được chọn
        if not selected_IDs:  # Tập hợp rỗng
            QMessageBox.information(self, "Thông báo", "Chưa có dòng nào được chọn")
            return
        # Hiển thị cảnh báo xác nhận
        reply = QMessageBox.question(
            self,
            "Xác nhận xóa",
            f"Bạn có chắc chắn muốn xóa {len(selected_IDs)} dòng đã chọn không?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.No:
            return  # Không làm gì nếu người dùng chọn "No"
        try:
            connection = connect_to_db()
            if connection is None:
                self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
                return
            cursor = connection.cursor()
            # Chuyển danh sách ID thành chuỗi tham số
            placeholders = ", ".join(["?"] * len(selected_IDs))
            query = f"DELETE FROM QA_PRT WHERE ID IN ({placeholders})"
            # Thực thi câu lệnh với tham số
            cursor.execute(query, tuple(selected_IDs))
            connection.commit()
            
            QMessageBox.information(self, "Thông báo", f"Xóa thành công {len(selected_IDs)} dữ liệu!")
            self.search_QA()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi xóa dữ liệu: {e}")    
        finally:
            if connection:
                connection.close()   

    def delete_selected_rows_Checker(self):
        # Lấy ID danh sách các hàng được chọn
        selected_IDs  = set(self.tableWidget_2.item(index.row(),31).text() 
                          for index in self.tableWidget_2.selectedIndexes()
                          if self.tableWidget_2.item(index.row(),31) is not None)
        # Kiểm tra nếu không có hàng nào được chọn
        if not selected_IDs:  # Tập hợp rỗng
            QMessageBox.information(self, "Thông báo", "Chưa có dòng nào được chọn")
            return
        # Hiển thị cảnh báo xác nhận
        reply = QMessageBox.question(
            self,
            "Xác nhận xóa",
            f"Bạn có chắc chắn muốn xóa {len(selected_IDs)} dòng đã chọn không?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.No:
            return  # Không làm gì nếu người dùng chọn "No"
        try:
            connection = connect_to_db()
            if connection is None:
                self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
                return
            cursor = connection.cursor()
            # Chuyển danh sách ID thành chuỗi tham số
            placeholders = ", ".join(["?"] * len(selected_IDs))
            query = f"DELETE FROM CHECKER_PRT WHERE ID IN ({placeholders})"
            # Thực thi câu lệnh với tham số
            cursor.execute(query, tuple(selected_IDs))
            connection.commit()
            
            QMessageBox.information(self, "Thông báo", f"Xóa thành công {len(selected_IDs)} dữ liệu!")
            self.search_Checker()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi xóa dữ liệu: {e}")    
        finally:
            if connection:
                connection.close()   

    def tai_xuong_file_mau_QA(self):
        headers = ['Ngày','Tuần','Nhóm','Máy','Style','Màu','Số kiểm','Mục tiêu','In sai hình','Lỗi màu in','Rỗ mặt','Lệch khung','Tắc khung','Thủng khung','Thấm mặt sau',
        'Xếp ly','Bong tróc','Rạn nứt','Chi tiết in','Bề mặt xấu','Lệch tâm','Thông số','Sai vị trí','Cong,méo','Lem mực','Bẩn','Cục vải,bàn','Khác','Lỗi vải','NG','Lỗi chết','ID']
        df = table_to_dataframe(self.tableWidget,headers)
        df = df.drop(columns=['ID'])
        # Create a new Excel workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "QA"

        # Add the title to A1
        sheet['A1'] = "MẪU FILE LỖI QA PRT NHẬP VÀO HỆ THỐNG"
        # Merge cells from A1 to J1
        sheet.merge_cells('A1:AE1')

        # Center-align the merged cells
        sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
        sheet['A1'].font = Font(color="FFFFFF", bold=True,size=14)
        sheet['A1'].fill = PatternFill(start_color="349eeb", end_color="349eeb", fill_type="solid")
        # Apply bold style to headers (A2:J2)
        for row in sheet.iter_rows(min_row=2, max_row=2, min_col=1, max_col=31):
            for cell in row:
                cell.font = Font(bold=True)
                # if cell.column <=9:
                #     cell.font = Font(color="FF0000",bold=True)
        # Write DataFrame to the sheet, starting from the second row
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=2):
            for c_idx, value in enumerate(row, start=1):
                sheet.cell(row=r_idx, column=c_idx, value=value)
   
        # Open file dialog to select save location
        options = QFileDialog.Options()
        file_name = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Tải xuống file", 
            f"File lỗi QA-PRT {file_name}", 
            "Excel Files (*.xlsx);;All Files (*)", 
            options=options
        )

        if not file_path:
            # QMessageBox.information(self, "Thông báo", "Bạn đã hủy việc tải xuống.")
            return

        # Save the file to the selected location
        try:
            workbook.save(file_path)
            QMessageBox.information(self, "Thông báo", f"File đã được tải xuống thành công tại:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi lưu file: {e}")
    
    def tai_xuong_file_mau_Checker(self):
        headers = ['Ngày','Tuần','Nhóm','Máy','Style','Màu','Số kiểm','Mục tiêu','In sai hình','Lỗi màu in','Rỗ mặt','Lệch khung','Tắc khung','Thủng khung','Thấm mặt sau',
        'Xếp ly','Bong tróc','Rạn nứt','Chi tiết in','Bề mặt xấu','Lệch tâm','Thông số','Sai vị trí','Cong,méo','Lem mực','Bẩn','Cục vải,bàn','Khác','Lỗi vải','NG','Lỗi chết','ID']
        df = table_to_dataframe(self.tableWidget,headers)
        df = df.drop(columns=['ID'])
        # Create a new Excel workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Checker"

        # Add the title to A1
        sheet['A1'] = "MẪU FILE LỖI CHECKER PRT NHẬP VÀO HỆ THỐNG"
        # Merge cells from A1 to J1
        sheet.merge_cells('A1:AE1')

        # Center-align the merged cells
        sheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
        sheet['A1'].font = Font(color="FFFFFF", bold=True,size=14)
        sheet['A1'].fill = PatternFill(start_color="349eeb", end_color="349eeb", fill_type="solid")
        # Apply bold style to headers (A2:J2)
        for row in sheet.iter_rows(min_row=2, max_row=2, min_col=1, max_col=31):
            for cell in row:
                cell.font = Font(bold=True)
                # if cell.column <=9:
                #     cell.font = Font(color="FF0000",bold=True)
        # Write DataFrame to the sheet, starting from the second row
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=2):
            for c_idx, value in enumerate(row, start=1):
                sheet.cell(row=r_idx, column=c_idx, value=value)
   
        # Open file dialog to select save location
        options = QFileDialog.Options()
        file_name = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Tải xuống file", 
            f"File lỗi Checker-PRT {file_name}", 
            "Excel Files (*.xlsx);;All Files (*)", 
            options=options
        )

        if not file_path:
            # QMessageBox.information(self, "Thông báo", "Bạn đã hủy việc tải xuống.")
            return

        # Save the file to the selected location
        try:
            workbook.save(file_path)
            QMessageBox.information(self, "Thông báo", f"File đã được tải xuống thành công tại:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi lưu file: {e}")
    
def main():
    app = QApplication(sys.argv)
    window = MainApp()
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()

