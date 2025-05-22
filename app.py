import sys
import os
import time
import tempfile
from pathlib import Path
from datetime import datetime,date
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
from PyQt5.QtGui import QDoubleValidator,QIntValidator
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
        self.bt203.clicked.connect(self.insert_cong_viec)
        self.bt300.clicked.connect(self.insert_chi_tiet_cong_viec)
        self.bt204.clicked.connect(self.save_cong_viec)
        self.bt301.clicked.connect(self.save_chi_tiet_cong_viec)
        self.bt205.clicked.connect(self.delete_cong_viec)
        self.bt302.clicked.connect(self.delete_chi_tiet_cong_viec)
        self.menu11.triggered.connect(self.show_tab_1)
        self.menu13.triggered.connect(self.show_tab_2)
        self.menu51.triggered.connect(self.show_login_tab)

        self.de101.setDate(QDate.currentDate().addDays(-365))
        self.de101.setCalendarPopup(True)
        self.de102.setDate(QDate.currentDate())
        self.de102.setCalendarPopup(True)
        self.de301.setDate(QDate.currentDate())
        self.de301.setCalendarPopup(True)  # Bật popup chọn lịch
        self.de302.setDate(QDate.currentDate())  # Gán ngày hiện tại
        self.de302.setCalendarPopup(True)

        self.le304.setValidator(QDoubleValidator(0, 100, 2))
        self.le305.setValidator(QIntValidator(0,100))
        self.le306.setValidator(QIntValidator(1,5))
        self.le307.setValidator(QIntValidator(1,5))
        # self.bt204.clicked.connect(self.tai_xuong_file_mau_Checker)
        # ####
        self.cb200.currentIndexChanged.connect(self.change_cb200)
        self.cb301.currentIndexChanged.connect(self.change_cb301)
        self.cb202.currentIndexChanged.connect(self.load_lb202)
        self.cb203.currentIndexChanged.connect(self.load_cb204)
        self.cb204.currentIndexChanged.connect(self.load_lb204)
        self.cb204.currentIndexChanged.connect(self.load_cb205)
        self.cb205.currentIndexChanged.connect(self.load_lb205)
        self.cb205.currentIndexChanged.connect(self.load_cb206)
        self.cb206.currentIndexChanged.connect(self.load_lb206)
        self.cb304.currentIndexChanged.connect(self.load_lb304)
        self.bt201.clicked.connect(self.add_du_an)
        self.bt207.clicked.connect(self.add_cong_viec)
        self.bt303.clicked.connect(self.add_chi_tiet_cong_viec)
        self.bt304.clicked.connect(self.show_tab_1)
        self.rd201.toggled.connect(self.rd201_change)
        self.rd301.toggled.connect(self.rd301_change)
        self.bt102.clicked.connect(self.search_cong_viec)
        self.tableWidget.itemDoubleClicked.connect(self.on_table_double_click)

    def rd201_change(self):
        if self.rd201.isChecked():
            self.cb200.clear()
            self.bt203.setVisible(True)
            self.bt204.setVisible(False)
            self.bt205.setVisible(False)
            self.cb200.setVisible(False)
            self.lb200.setVisible(False)
            self.lb003.setText("Thêm công việc mới")
        else:
            self.load_cb200()
            self.bt203.setVisible(False)
            self.bt204.setVisible(True)
            self.bt205.setVisible(True)
            self.cb200.setVisible(True)
            self.lb200.setVisible(True)
            self.lb003.setText("Sửa hoặc xóa công việc")
    def rd301_change(self):
        if self.rd301.isChecked():
            self.cb301.clear()
            self.bt300.setVisible(True)
            self.bt301.setVisible(False)
            self.bt302.setVisible(False)
            self.cb301.setVisible(False)
            self.lb301.setVisible(False)
            self.widget_302.setVisible(False)

        else:
            # self.load_cb200()
            self.bt300.setVisible(False)
            self.bt301.setVisible(True)
            self.bt302.setVisible(True)
            self.cb301.setVisible(True)
            self.lb301.setVisible(True)
            if self.lb005.text() == "Quản lý":
                self.widget_302.setVisible(True)
            else:
                self.widget_302.setVisible(False)
            self.load_cb301()

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
            self.lb005.setText(result[3])
            self.cb104.setCurrentIndex(self.cb104.findText(result[0]))

            self.search_cong_viec()
        else:
            self.lb002.setText("Tài khoản hoặc mật khẩu không đúng!")

    def add_du_an(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_du_an = self.le201.text().strip()
        ten_du_an = self.le202.text().strip()
        # ten_du_an = self.te201.toPlainText()

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
        self.le202.setText("")

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
    def add_chi_tiet_cong_viec(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        mnv = self.lb004.text().strip()
        ten_cv = self.le303.text()

        if mnv and ten_cv:
            try:
                cursor.execute(
                    "INSERT INTO GHI_NHO_TEN_CONG_VIEC (MNV, Ten_chi_tiet_cong_viec) VALUES (?, ?)",
                    (mnv, ten_cv)
                )
                connection.commit()
                self.msgbox("✅ Thêm chi tiết công việc mới thành công")
            except Exception as e:
                print("Lỗi khi thêm hi tiết công việc mới:", e)
                self.msgbox("⚠️ Tên hi tiết công việc đã tồn tại hoặc có lỗi khác")
        else:
            self.msgbox("⚠️ Vui lòng nhập đầy đủ tên hi tiết công việc mới")

        connection.close()
        self.load_cb302()
        self.le303.setText("")
    def load_cb200(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        nguoi_tao = self.lb004.text() + '-' + self.lb001.text()
        cursor.execute("SELECT ID FROM DANH_SACH_CONG_VIEC WHERE Nguoi_cap_nhat = ?",(nguoi_tao,))
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
    def change_cb200(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        id = self.cb200.currentText()
        nguoi_cap_nhat = self.lb004.text() + '-' + self.lb001.text()
        cursor.execute("SELECT * FROM DANH_SACH_CONG_VIEC WHERE ID = ? and Nguoi_cap_nhat = ?",(id,nguoi_cap_nhat))
        results = cursor.fetchall()
 
        if results:
            ma_cong_ty = results[0][1]
            nhom = results[0][3]
            ma_du_an = results[0][4]
            ten_cong_viec = results[0][5]
            phan_loai_cv = results[0][6]
            ghi_chu = results[0][12]
            phan_loai_du_an = results[0][17]
            nhiem_vu = results[0][18]
            nhiem_vu_cu_the = results[0][19]
            chuc_nang = results[0][20]

            self.cb208.setCurrentIndex(self.cb208.findText(ma_cong_ty))
            self.cb201.setCurrentIndex(self.cb201.findText(nhom))
            self.cb202.setCurrentIndex(self.cb202.findText(ma_du_an))
            self.cb203.setCurrentIndex(self.cb203.findText(phan_loai_du_an))
            self.cb204.setCurrentIndex(self.cb204.findText(chuc_nang))
            self.cb205.setCurrentIndex(self.cb205.findText(nhiem_vu))
            self.cb206.setCurrentIndex(self.cb206.findText(nhiem_vu_cu_the))
            self.cb207.setCurrentIndex(self.cb207.findText(ten_cong_viec))
            self.cb210.setCurrentIndex(self.cb210.findText(phan_loai_cv))
            self.te201.setPlainText(ghi_chu)

        connection.close()
        self.show_tab_1
    def change_cb301(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        id = self.cb301.currentText()
        cursor.execute("""SELECT Chi_tiet_cong_viec,Ngay_bat_dau,Thoi_han,Thoi_luong,Tien_do,
                       Nguoi_thuc_hien,Ghi_chu,Diem_tien_do,Diem_chat_luong
                       FROM CHI_TIET_CONG_VIEC 
                       WHERE ID = ?""",(id,))
        results = cursor.fetchall()
 
        if results:
            chi_tiet_cong_viec = results[0][0]
            ngay_bat_dau = results[0][1]
            Thoi_han = results[0][2]
            Thoi_luong = results[0][3]
            Tien_do = int(results[0][4])
            Nguoi_thuc_hien = results[0][5]
            Ghi_chu = results[0][6]
            Diem_tien_do = results[0][7]
            Diem_chat_luong = results[0][8]

            self.cb302.setCurrentIndex(self.cb302.findText(chi_tiet_cong_viec))
            date_parts = ngay_bat_dau.split("-")
            year, month, day = map(int, date_parts)
            qdate = QDate(year, month, day)
            self.de301.setDate(qdate)
            date_parts = Thoi_han.split("-")
            year, month, day = map(int, date_parts)
            qdate = QDate(year, month, day)
            self.de302.setDate(qdate)
            self.le304.setText(str(Thoi_luong))
            self.le305.setText(str(Tien_do))
            self.le306.setText(str(Diem_tien_do))
            self.le307.setText(str(Diem_chat_luong))
            self.te302.setPlainText(Ghi_chu)

        connection.close()

    def insert_cong_viec(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_cong_ty = self.cb208.currentText()
        ma_phong_ban = self.lb000.text()
        nhom = self.cb201.currentText()
        ma_du_an = self.cb202.currentText()
        ten_cong_viec = self.cb207.currentText()
        phan_loai_cv = self.cb210.currentText()
        ngay_bat_dau = date.today()
        thoi_luong = 0
        thoi_han = date.today()
        trang_thai = ""
        tien_do = 0
        ghi_chu = self.te201.toPlainText()
        diem_tien_do = 0
        diem_chat_luong = 0
        thoi_diem_cap_nhat = datetime.now()
        nguoi_cap_nhat = self.lb004.text() + '-' + self.lb001.text()
        phan_loai_du_an = self.cb203.currentText()
        nhiem_vu = self.cb205.currentText()
        nhiem_vu_cu_the = self.cb206.currentText()
        chuc_nang = self.cb204.currentText()
        try:
            cursor.execute("""INSERT INTO DANH_SACH_CONG_VIEC (Ma_cong_ty,Ma_phong_ban,Nhom,Ma_du_an,
                        Ten_cong_viec,Phan_loai_cv,Ngay_bat_dau,Thoi_luong,Thoi_han,Trang_thai,Tien_do,
                        Ghi_chu,Diem_tien_do,Diem_chat_luong,Thoi_diem_cap_nhat,Nguoi_cap_nhat,
                        Phan_loai_du_an,Nhiem_vu,Nhiem_vu_cu_the,Chuc_nang) 
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,(ma_cong_ty,ma_phong_ban,nhom,ma_du_an,ten_cong_viec,phan_loai_cv,ngay_bat_dau,
                                thoi_luong,thoi_han,trang_thai,tien_do,ghi_chu,diem_tien_do,diem_chat_luong,
                                thoi_diem_cap_nhat,nguoi_cap_nhat,phan_loai_du_an,nhiem_vu,nhiem_vu_cu_the,chuc_nang))
            connection.commit()
            self.msgbox("✅ Thêm công việc mới thành công")
            self.show_tab_1()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi thêm công việc mới: {e}")  

        connection.close()
    def insert_chi_tiet_cong_viec(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        id_cv = self.le301.text()
        nguoi_thuc_hien = self.cb304.currentText()
        chuc_danh = self.lb304.text()
        chi_tiet_cong_viec = self.cb302.currentText()
        ngay_bat_dau = self.de301.date().toString("yyyy-MM-dd")
        thoi_han = self.de302.date().toString("yyyy-MM-dd")
        try:
            thoi_luong = float(self.le304.text())
        except Exception  as e:
            self.msgbox("⚠️ Vui lòng nhập thời lượng (giờ)")
            return
        try:
            tien_do = int(self.le305.text())
        except Exception  as e:
            self.msgbox("⚠️ Vui lòng nhập tiến độ công việc")
            return
        trang_thai = "Chưa thực hiện" if tien_do == 0 else "Đã hoàn thành" if tien_do == 100 else "Đang thực hiện"
        ghi_chu = self.te302.toPlainText()
        diem_tien_do = 3
        diem_chat_luong = 3
        thoi_diem_cap_nhat = datetime.now()
        nguoi_tao = self.lb004.text() + '-' + self.lb001.text()
        diem_tien_do_x_thoi_luong = diem_tien_do * thoi_luong
        diem_chat_luong_x_thoi_luong = diem_chat_luong * thoi_luong
        
        try:
            cursor.execute("""INSERT INTO CHI_TIET_CONG_VIEC (ID_CV,Nguoi_thuc_hien,Chuc_danh,
                        Chi_tiet_cong_viec,Ngay_bat_dau,Thoi_luong,Thoi_han,Trang_thai,Tien_do,
                        Ghi_chu,Diem_tien_do,Diem_chat_luong,Thoi_diem_cap_nhat,Nguoi_tao,
                        Diem_tien_do_x_thoi_luong,Diem_chat_luong_x_thoi_luong) 
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,(id_cv,nguoi_thuc_hien,chuc_danh,chi_tiet_cong_viec,ngay_bat_dau,
                                thoi_luong,thoi_han,trang_thai,tien_do,ghi_chu,diem_tien_do,diem_chat_luong,
                                thoi_diem_cap_nhat,nguoi_tao,diem_tien_do_x_thoi_luong,diem_chat_luong_x_thoi_luong))
            connection.commit()
            self.msgbox("✅ Thêm chi tiết công việc mới thành công")
            self.search_chi_tiet_cong_viec()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi thêm chi tiết công việc mới: {e}")  

        connection.close()

    def save_cong_viec(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        id = self.cb200.currentText()
        ma_cong_ty = self.cb208.currentText()
        ma_phong_ban = self.lb000.text()
        nhom = self.cb201.currentText()
        ma_du_an = self.cb202.currentText()
        ten_cong_viec = self.cb207.currentText()
        phan_loai_cv = self.cb210.currentText()
        ghi_chu = self.te201.toPlainText()
        thoi_diem_cap_nhat = datetime.now()
        phan_loai_du_an = self.cb203.currentText()
        nhiem_vu = self.cb205.currentText()
        nhiem_vu_cu_the = self.cb206.currentText()
        chuc_nang = self.cb204.currentText()
        try:
            cursor.execute("""UPDATE DANH_SACH_CONG_VIEC SET 
                           Ma_cong_ty = ?,Nhom = ?,Ma_du_an = ?,
                        Ten_cong_viec = ?,Phan_loai_cv = ?,
                        Ghi_chu = ?,Thoi_diem_cap_nhat = ?,
                        Phan_loai_du_an = ?,Nhiem_vu = ?,Nhiem_vu_cu_the = ?,Chuc_nang = ?
                        WHERE ID = ?
                        """,(ma_cong_ty,nhom,ma_du_an,ten_cong_viec,phan_loai_cv,
                                ghi_chu,thoi_diem_cap_nhat,phan_loai_du_an,nhiem_vu,nhiem_vu_cu_the,chuc_nang,id))
            connection.commit()
            self.msgbox("✅ Đã lưu thay đổi thông tin công việc")
            self.show_tab_1()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi lưu thay đổi thông tin công việc: {e}")  

        connection.close()
    def save_chi_tiet_cong_viec(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        id = self.cb301.currentText()
        chi_tiet_cong_viec = self.cb302.currentText()
        ngay_bat_dau = self.de301.date().toString("yyyy-MM-dd")
        thoi_han = self.de302.date().toString("yyyy-MM-dd")
        thoi_luong = float(self.le304.text())
        tien_do = self.le305.text()
        nguoi_thuc_hien = self.cb304.currentText()
        chuc_danh = self.le304.text()
        ghi_chu = self.te302.toPlainText()
        thoi_diem_cap_nhat = datetime.now()
        diem_tien_do = int(self.le306.text())
        diem_chat_luong = int(self.le307.text())
        diem_tien_do_x_thoi_luong = diem_tien_do * thoi_luong
        diem_chat_luong_x_thoi_luong = diem_chat_luong * thoi_luong
        try:
            cursor.execute("""UPDATE CHI_TIET_CONG_VIEC SET 
                           Chi_tiet_cong_viec = ?,Ngay_bat_dau = ?,Thoi_han = ?,
                        Thoi_luong = ?,Tien_do = ?,
                        Nguoi_thuc_hien = ?,Chuc_danh = ?,
                        Ghi_chu = ?,Thoi_diem_cap_nhat = ?,Diem_tien_do = ?,Diem_chat_luong = ?,
                        diem_tien_do_x_thoi_luong = ?, diem_chat_luong_x_thoi_luong = ?
                        WHERE ID = ?
                        """,(chi_tiet_cong_viec,ngay_bat_dau,thoi_han,thoi_luong,tien_do,
                                nguoi_thuc_hien,chuc_danh,ghi_chu,thoi_diem_cap_nhat,diem_tien_do,diem_chat_luong,
                                diem_tien_do_x_thoi_luong,diem_chat_luong_x_thoi_luong,id))
            connection.commit()
            self.msgbox("✅ Đã lưu thay đổi thông tin chi tiết công việc")
            self.search_chi_tiet_cong_viec()
        except Exception  as e:
            QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi lưu thay đổi thông tin chi tiết công việc: {e}")  

        connection.close()

    def delete_cong_viec(self):
        reply = QMessageBox.question(
        self,
        "Xác nhận xóa",
        "Khi bạn xóa công việc, sẽ xóa hết chi tiết công việc liên quan đến mã công việc này.\nBạn có chắc chắn muốn tiếp tục?",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            connection = connect_to_db()
            cursor = connection.cursor()

            id = self.cb200.currentText()
            try:
                cursor.execute("""DELETE FROM DANH_SACH_CONG_VIEC 
                            WHERE ID = ?
                            """,(id,))
                connection.commit()

                cursor.execute("""DELETE FROM CHI_TIET_CONG_VIEC 
                            WHERE ID_CV = ?
                            """,(id,))
                connection.commit()

                self.msgbox("✅ Đã xóa thành công!")
                self.show_tab_1()
            except Exception  as e:
                QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi xóa công việc: {e}")  

            connection.close()
            self.show_tab_1
    def delete_chi_tiet_cong_viec(self):
        reply = QMessageBox.question(
        self,
        "Xác nhận xóa",
        "Bạn có chắc chắn muốn xóa chi tiết công việc này?",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            connection = connect_to_db()
            cursor = connection.cursor()

            id = self.cb301.currentText()
            try:
                cursor.execute("""DELETE FROM CHI_TIET_CONG_VIEC 
                            WHERE ID = ?
                            """,(id,))
                connection.commit()

                self.msgbox("✅ Đã xóa thành công!")
                self.search_chi_tiet_cong_viec()
            except Exception  as e:
                QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi xóa chi tiết công việc: {e}")  

            connection.close()
            self.search_chi_tiet_cong_viec()
            self.load_cb301()

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
            cursor.execute("""SELECT Ten_cong_viec FROM GHI_NHO_TEN_CONG_VIEC WHERE  
                           Ten_cong_viec IS NOT NULL AND MNV = ?""",(mnv,))
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
    def load_cb302(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        mnv = self.lb004.text()
        cursor.execute("""SELECT Ten_chi_tiet_cong_viec FROM GHI_NHO_TEN_CONG_VIEC WHERE  
                Ten_chi_tiet_cong_viec IS NOT NULL AND MNV = ?""",(mnv,))
        results = cursor.fetchall()

        self.cb302.clear()  # Xóa các item hiện có
        for row in results:
            self.cb302.addItem(row[0])  

        connection.close()
        self.cb302.setEditable(True)
        self.cb302.completer().setFilterMode(QtCore.Qt.MatchContains)  # Gợi ý theo từ khóa chứa
        self.cb302.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.cb302.setStyleSheet("""
            QComboBox {
                background-color: #2e2e2e;
                color: white;
            }
            QComboBox QAbstractItemView {
                background-color: #3c3c3c;
                color: white;
                font-size: 13pt;              
            }
        """)
    def load_cb301(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        id_cv = self.le301.text()
        mnv = self.lb004.text() + '-' + self.lb001.text()
        phan_quyen = self.lb005.text()
        if phan_quyen == "Quản lý":
            cursor.execute("""SELECT ID FROM CHI_TIET_CONG_VIEC WHERE  
                ID_CV  = ?""",(id_cv,))
        else:
            cursor.execute("""SELECT ID FROM CHI_TIET_CONG_VIEC WHERE  
                ID_CV  = ? AND Nguoi_tao = ?""",(id_cv,mnv))
        results = cursor.fetchall()

        self.cb301.clear()  # Xóa các item hiện có
        for row in results:
            self.cb301.addItem(str(row[0]))

        connection.close()
        self.cb301.setEditable(True)
        self.cb301.completer().setFilterMode(QtCore.Qt.MatchContains)  # Gợi ý theo từ khóa chứa
        self.cb301.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.cb301.setStyleSheet("""
            QComboBox {
                background-color: #2e2e2e;
                color: white;
            }
            QComboBox QAbstractItemView {
                background-color: #3c3c3c;
                color: white;
                font-size: 13pt;              
            }
        """)
    def load_cb304(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        ma_pb = self.lb000.text()
        cursor.execute("""SELECT MNV, Ho_ten FROM DANH_SACH_CBCNV WHERE  
                Ma_phong_ban  = ?""",(ma_pb,))
        results = cursor.fetchall()

        self.cb304.clear()  # Xóa các item hiện có
        for row in results:
            self.cb304.addItem(str(row[0]) + '-' + str(row[1]))

        connection.close()
        self.cb304.setEditable(True)
        self.cb304.completer().setFilterMode(QtCore.Qt.MatchContains)  # Gợi ý theo từ khóa chứa
        self.cb304.completer().setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.cb304.setStyleSheet("""
            QComboBox {
                background-color: #2e2e2e;
                color: white;
            }
            QComboBox QAbstractItemView {
                background-color: #3c3c3c;
                color: white;
                font-size: 13pt;              
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

    def load_lb304(self):
        connection = connect_to_db()
        cursor = connection.cursor()

        mnv = self.cb304.currentText().split("-")[0].strip()
        cursor.execute("SELECT Chuc_danh FROM DANH_SACH_cbcnv where MNV = ?",(mnv,))
        results = cursor.fetchone()
        if results:
            self.lb304.setText(results[0])
        else:
            self.lb304.setText("")
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
    def search_cong_viec(self):
        ma_cong_ty = self.cb103.currentText()
        if ma_cong_ty == "Tất cả":
            ma_cong_ty = ""  

        ma_phong_ban = self.cb104.currentText()
        if ma_phong_ban == "Tất cả":
            ma_phong_ban = ""  

        tu_ngay = self.de101.date().toString("yyyy-MM-dd")
        den_ngay = self.de102.date().toString("yyyy-MM-dd")

        trang_thai = self.cb101.currentText()
        if trang_thai == "Tất cả":
            trang_thai = ""

        phan_loai_du_an = self.cb105.currentText()
        if phan_loai_du_an == "Tất cả":
            phan_loai_du_an = ""
        
        phan_loai_cong_viec = self.cb102.currentText()
        if phan_loai_cong_viec == "Tất cả":
            phan_loai_cong_viec = ""
       
        connection = connect_to_db()
        if connection is None:
            self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
            return
        cursor = connection.cursor()
        sql = f"""
                SELECT ID,Ma_cong_ty,Ma_phong_ban,Nhom,Phan_loai_du_an,Ma_du_an,Chuc_nang,
                Nhiem_vu,Nhiem_vu_cu_the,Ten_cong_viec,Phan_loai_cv,Ngay_bat_dau,Thoi_luong,
                Thoi_han,Trang_thai,Tien_do,Ghi_chu,Diem_tien_do,Diem_chat_luong,Thoi_diem_cap_nhat,
                Nguoi_cap_nhat
                FROM DANH_SACH_CONG_VIEC WHERE 
                Trang_thai LIKE ? and Ma_cong_ty LIKE ? and Phan_loai_du_an LIKE ? and
                Phan_loai_cv LIKE ? and Ma_phong_ban LIKE ? 
                and (Ngay_bat_dau BETWEEN ? AND ?)
                ORDER BY ID DESC
            """
        cursor.execute(sql,(f"%{trang_thai}%",f"%{ma_cong_ty}%",f"%{phan_loai_du_an}%",
                            f"%{phan_loai_cong_viec}%",f"%{ma_phong_ban}%",tu_ngay,den_ngay))
        results = cursor.fetchall()
        # Xóa dữ liệu cũ trong TableWidget
        self.tableWidget.setRowCount(0)
        
        # Hiển thị dữ liệu trong TableWidget
        for row_idx, row_data in enumerate(results):
            self.tableWidget.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                if col_idx == 19:
                    value = value[:16]
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.tableWidget.setItem(row_idx, col_idx, item)

            # Cập nhật progress bar
            self.progressBar.setValue(int((row_idx + 1) * 100 / len(results)))
        
        self.tableWidget.resizeColumnsToContents()

        # Gọi hàm tổng số dòng
        self.tong_so_dong_tab_1

    def search_chi_tiet_cong_viec(self):
        connection = connect_to_db()
        if connection is None:
            self.lb003.setText("Không thể kết nối tới cơ sở dữ liệu!")
            return
        cursor = connection.cursor()
        id_cv = self.le301.text()
        print(id_cv)
        sql = f"""
                SELECT ID,Chi_tiet_cong_viec,Thoi_luong,Ngay_bat_dau,
                Thoi_han,Tien_do,Trang_thai,Ghi_chu,Nguoi_thuc_hien,Chuc_danh,Diem_tien_do,Diem_chat_luong,
                Thoi_diem_cap_nhat,Nguoi_tao
                FROM CHI_TIET_CONG_VIEC WHERE ID_CV = ?
                ORDER BY ID DESC
            """
        cursor.execute(sql,(id_cv,))
        results = cursor.fetchall()
        # Xóa dữ liệu cũ trong TableWidget
        self.tableWidget_2.setRowCount(0)
        
        # Hiển thị dữ liệu trong TableWidget
        for row_idx, row_data in enumerate(results):
            self.tableWidget_2.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                if col_idx == 12:
                    value = value[:16]
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.tableWidget_2.setItem(row_idx, col_idx, item)
        
        self.tableWidget_2.resizeColumnsToContents()
    
    def show_login_tab(self):
        self.tabWidget.setCurrentIndex(0) 
        self.tb001.setText("")
        self.tb002.setText("")
        self.lb000.setText("")
        self.lb001.setText("")
        self.lb002.setText("")
        self.lb003.setText("Phần mềm quản lý kho vải")
        self.menuBar.setVisible(False)
                        
    def show_tab_1(self):
        self.tabWidget.setCurrentIndex(1) 
        self.lb003.setText("Danh sách công việc")
        self.search_cong_viec()
     
    def show_tab_2(self):
        self.menuBar.setVisible(True)
        self.tabWidget.setCurrentIndex(2) 
        self.lb003.setText("Thêm công việc mới")  
        self.load_cb201()
        self.load_cb202()
        self.load_cb204()
        self.load_cb207()
        self.rd201.setChecked(True)

    def show_tab_3(self,id_cv,ma_pb,ten_cv,ghi_chu):
            self.tabWidget.setCurrentIndex(3) 
            self.lb003.setText("Chi tiết công việc")
            self.le301.setText(id_cv)
            self.le301.setReadOnly(True)
            self.le302.setText(ten_cv)
            self.le302.setReadOnly(True)
            self.te301.setPlainText(ghi_chu)
            self.te301.setReadOnly(True)
            if ma_pb == self.lb000.text():
                self.frame_301.setVisible(True)
                self.rd301.setChecked(True)
            else:
                self.frame_301.setVisible(False)
            self.load_cb302()
            self.load_cb304()
            self.search_chi_tiet_cong_viec()
    
    def on_table_double_click(self, item):
        row = item.row()
        id_item = self.tableWidget.item(row, 0)
        ma_pb_item = self.tableWidget.item(row, 2)
        ten_cv_item = self.tableWidget.item(row, 9)
        ghi_chu_item = self.tableWidget.item(row, 16)
        if id_item:
            id_cv = id_item.text()
            ma_pb = ma_pb_item.text()
            ten_cv = ten_cv_item.text()
            ghi_chu = ghi_chu_item.text()

            self.show_tab_3(id_cv,ma_pb,ten_cv,ghi_chu) 
        
    def tong_so_dong_tab_1(self):
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

