import os, sys,glob,shutil
import pytesseract
from pdf2image import convert_from_path
from PyQt6 import QtCore, QtGui, QtWidgets, uic#, QtWebEngineWidgets #, QWebEngineSettings
from PyQt6.QtCore import Qt
# import numpy as np
from pathlib import Path
from PIL import Image
from pdf2docx import Converter
import matplotlib.image as mpimg
import pyqtgraph as pg
# from PIL import Image
import numpy as np
from pikepdf import Pdf, Permissions, Encryption


    

class MainWindow(QtWidgets.QMainWindow):
 
    def __init__(self):
        super().__init__()
 
        uic.loadUi(r'.\pdf_helper.ui', self)
        self.setWindowTitle('PDF小幫手: ')
 
 
        #Signals
        self.actionExit.triggered.connect(self.fileExit)
        #page1:pdf轉word
        self.pushButton_choosepdf.clicked.connect(self.fileOpen_pdf_convert)
        self.pushButton_choose_word.clicked.connect(self.fileOpen_docx_convert)
        self.pushButton_start_convert.clicked.connect(self.file_convertpdf)

        #page2:圖片文字提取
        self.pushButton_txt_pdf.clicked.connect(self.fileOpen_pdf_text)
        self.pushButton_gain_txt.clicked.connect(self.gain_text)
        self.savetxt_pushButton.clicked.connect(self.saveData_txt)


        #page3:pdf合併
        self.pdf_count = []
        self.all_data = []
        self.fname_merge = ""
        self.graphicsView.setBackground('w')
        self.graphicsView.getAxis('bottom').setTicks('')
        self.graphicsView.getAxis('left').setTicks('')
        self.pushButton_pdfmerge.clicked.connect(self.fileOpen_pdf_merge)
        self.pushButton_pdfview.clicked.connect(self.deal)
        self.listWidget_pdf.itemSelectionChanged.connect(self.loadImage)
        self.pushButton_deletepdf.clicked.connect(self.delete_current_pdf)
        self.pushButton_merge_all.clicked.connect(self.merge_pdf)
        # self.pushButton_mergepdf_output.clicked.connect(self.save_merge_pdf) 
        #拆分所有
        self.pushButton_check_dismantle_all.clicked.connect(self.check_dismantle_pdf_all) 
        self.pushButton_dismantle_all.clicked.connect(self.dismantle_pdf_all) 
        #拆分範圍
        self.pushButton_check_dismantle_sub.clicked.connect(self.check_dismantle_pdf_sub) 
        self.pushButton_check_dismantle_sub_2.clicked.connect(self.check_dismantle_pdf_sub_chose)
        self.pushButton_dismantle_sub.clicked.connect(self.dismantle_pdf_sub)
        #拆分特定
        self.pushButton_check_dismantle_distinct.clicked.connect(self.check_dismantle_pdf_distinct) 
        self.pushButton_check_dismantle_distinct_2.clicked.connect(self.check_dismantle_pdf_distinct_chose)
        self.pushButton_dismantle_distinct.clicked.connect(self.dismantle_pdf_distinct)



    # 共用 Slots:
    def fileExit(self):
        self.close()

    # pdf轉word Slots----
    def fileOpen_pdf_convert(self):
        self.fname_pdf = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', 
            "", ";PDF Files (*.pdf);")
        self.lineEdit_convertpdf_file.setText(self.fname_pdf[0])

    def fileOpen_docx_convert(self):
        self.fname_docx = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file', 
            "", ";Word Files(*.doc *.docx);")
        self.lineEdit_convertword_file.setText(self.fname_docx[0])

    def file_convertpdf(self):
        try:
            if self.fname_pdf[0]:
                password = self.lineEdit_password.text()
                if self.checkBox_password.isChecked():
                    cv = Converter(self.fname_pdf[0], password)
                else:
                    cv = Converter(self.fname_pdf[0])
                cv.convert(self.fname_docx[0])      # all pages by default
                cv.close()
                self.label_convert_warning.setText('轉檔成功!')
        except:
            self.label_convert_warning.setText('檔案路徑有誤!請檢查')
            

    # 圖片文字提取 Slots----
    def fileOpen_pdf_text(self):
        # home_dir = str(Path.home())
        self.fname = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', 
            "", ";PDF Files (*.pdf);;Images (*.png *.xpm *.jpg)")
        self.lineEdit_txt_pdf.setText(self.fname[0])
 
    def gain_text(self):
        # print(fname[0])
        try:
            self.text_data = ''
            if self.fname[0].split('.')[-1] == 'pdf':
                pytesseract.pytesseract.tesseract_cmd = r".\Tesseract-OCR\tesseract.exe"
                pages = convert_from_path(self.fname[0],poppler_path=r'.\poppler-0.68.0\bin')
                current_page = 1
                maxvalue = len(pages)
                self.progressBar_gaintxt.setMaximum(maxvalue)
                for page in pages:
                    self.text = pytesseract.image_to_string(page,lang="chi_tra+eng")
                    # print(self.current_page)
                    self.text_data += '=='*25+f'第{current_page}頁'+'=='*25  +'\n'
                    self.text_data += self.text + '\n'
                    current_page+=1
                    self.progressBar_gaintxt.setValue(current_page)
                self.output_textBrowser.setText(self.text_data)
                self.label_gain_txt.setText('解析成功!')
            
            elif self.fname[0].split('.')[-1] in ['png','xpm','jpg']:
                current_schedule = 0
                self.progressBar_gaintxt.setMaximum(1)
                pytesseract.pytesseract.tesseract_cmd = r".\Tesseract-OCR\tesseract.exe"
                img = Image.open(self.fname[0])
                self.text_data = pytesseract.image_to_string(img,lang="chi_tra+eng")
                current_schedule += 1
                self.progressBar_gaintxt.setValue(current_schedule)
                self.output_textBrowser.setText(self.text_data)
                self.label_gain_txt.setText('解析成功!')
            
            else:
                self.output_textBrowser.setText('檔案非PDF掃描檔or圖片檔(目前支援:.png,.xpm,.jpg)，請再檢查一下!')
        except:    
            self.label_gain_txt.setText('路徑有誤!請檢查')

    def saveData_txt(self):
        save_fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file', 
            "", "Text files (*.txt)")
        if len(save_fname) != 0:
            with open(save_fname,'w',encoding = 'UTF-8') as wtxt:
                wtxt.write(self.text_data)
                self.label_remind.setText('存檔成功!')

             
    # pdf合併----
    def fileOpen_pdf_merge(self):
        self.fname_merge = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', 
            "", ";PDF Files (*.pdf);;")
        if self.fname_merge[0]:
            self.pdf_count.append(self.fname_merge[0])
        self.pdf_count_tmp = [fname1.split('/')[-1].split('.')[0] for fname1 in self.pdf_count]
        self.comboBox_pdfs.clear()
        for fname in self.pdf_count_tmp:
            self.comboBox_pdfs.addItem(fname)

        # widgetpdf = widget_pdf(self.fname_merge)
        # print(self.fname_merge[0])
        # widgetpdf.show()

    def delete_current_pdf(self):
        col = self.comboBox_pdfs.currentText()
        del self.pdf_count[self.pdf_count_tmp.index(col)]
        # self.pdf_count_tmp = [fname1.split('/')[-1].split('.')[0] for fname1 in self.pdf_count]
        self.pdf_count_tmp.remove(col)
        self.comboBox_pdfs.clear()
        for fname in self.pdf_count_tmp:
            self.comboBox_pdfs.addItem(fname)


    def deal(self):
        #刪除先前檔案
        if self.fname_merge != "" and len(self.comboBox_pdfs) != 0:
            self.graphicsView.clear()
            for item in os.listdir('.\\images'):
                item_path = os.path.join('.\\images', item)
                # 如果是檔案，則直接刪除
                if os.path.isfile(item_path):
                    os.remove(item_path)
            # self.current_pdf = self.comboBox_pdfs.currentText()
            # 將當前pdf轉為圖檔储存
            f_name = self.pdf_count[self.comboBox_pdfs.currentIndex()]
            pdf_pages = convert_from_path(f_name,poppler_path=r'.\poppler-0.68.0\bin')
            i_page = 1
            for pdf_page in pdf_pages:
                pdf_page.save(r'.\images\tmp_picture'+str(i_page)+'.jpg','jpeg')
                i_page += 1
            imgs = glob.glob('.\\images\\*.jpg')
            #圖片轉尺寸
            # for i in imgs:
            #     im = Image.open(i)
            #     name = i.split('\\')[::-1][0]        # 取得圖片的名稱
            #     im2 = im.resize((1241, 1754))         # 調整圖片尺寸為 200x200
            #     im2.save(f'.\\images\\{name}')
            self.all_data = imgs

            def get_item_wight(data,i):
                # 读取属性
                ship_photo = data
                # 总Widget
                wight = QtWidgets.QWidget()
                # 总体横向布局
                layout_main = QtWidgets.QVBoxLayout()
                map_l = QtWidgets.QLabel() # 头像显示
                map_l.setFixedSize(200, 300)
                maps = QtGui.QPixmap(ship_photo).scaled(200, 300)
                map_l.setPixmap(maps)
                # # 右边的纵向布局
                # layout_right = QtWidgets.QVBoxLayout()
                # # 右下的的横向布局
                # layout_right_down = QHBoxLayout() # 右下的横向布局
                # layout_right_down.addWidget(QLabel(ship_type))
                # layout_right_down.addWidget(QLabel(ship_country))
                # layout_right_down.addWidget(QLabel(str(ship_star) + "星"))
                # layout_right_down.addWidget(QLabel(ship_index))
                # 按照从左到右, 从上到下布局添加
                layout_main.addWidget(map_l) # 最左边的头像
                Qlabel_text = QtWidgets.QLabel('第'+str(i)+'頁')
                Qlabel_text.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                layout_main.addWidget(Qlabel_text) # 右边的纵向布局
                # layout_right.addLayout(layout_right_down) # 右下角横向布局
                # layout_main.addLayout(layout_right) # 右边的布局
                wight.setLayout(layout_main) # 布局给wight
                return wight # 返回wight
            self.listWidget_pdf.clear()
            i = 1
            for ship_data in self.all_data:
                item = QtWidgets.QListWidgetItem() # 创建QListWidgetItem对象
                item.setSizeHint(QtCore.QSize(200, 350)) # 设置QListWidgetItem大小
                widget = get_item_wight(ship_data,i) # 调用上面的函数获取对应
                self.listWidget_pdf.addItem(item) # 添加item
                self.listWidget_pdf.setItemWidget(item, widget) # 为item设置widget
                i+=1
        else:
            self.listWidget_pdf.clear()
            self.graphicsView.clear()

        
    def loadImage(self):
        self.currentImgIdx = self.listWidget_pdf.currentIndex().row()
        if len(self.all_data) > 0:
            if self.currentImgIdx in range(len(self.all_data)):
                self.currentImg = QtGui.QPixmap(self.all_data[self.currentImgIdx]).scaledToHeight(500)
                # self.label_current_pdf.setPixmap(self.currentImg)
                image = mpimg.imread(self.all_data[self.currentImgIdx])
                # image = Image.open(self.all_data[self.currentImgIdx])
                # img_item = image
                img_item = pg.ImageItem(image, axisOrder='row-major')
                # self.graphicsView.setImage(np.array(image))
                # self.graphicsView.setBackground('w')

                self.graphicsView.addItem(img_item)
                self.graphicsView.invertY(True)
                # self.graphicsView.getAxis('bottom').setTicks('')
                # self.graphicsView.getAxis('left').setTicks('')
                self.graphicsView.setAspectLocked(lock=True, ratio=1)
    #合併所有pdf
    def merge_pdf(self):
        self.output = Pdf.new()
        for fname in self.pdf_count:
            # pdf_tmp = Pdf.open(fname).pages
            self.output.pages.extend(Pdf.open(fname).pages)
        save_fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file', "", "PDF Files (*.pdf)")
        self.output.save(save_fname)
        self.label_merge_save.setText('合併成功!')
    # def save_merge_pdf(self):
    #     save_fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file', "", "PDF Files (*.pdf)")
    #     self.output.save(save_fname)

    # 拆分檔案:所有頁數
    def check_dismantle_pdf_all(self):
        self.folderPath_dismantle_all = QtWidgets.QFileDialog.getExistingDirectory()
        self.lineEdit_check_dismantle_all.setText(self.folderPath_dismantle_all)

    def dismantle_pdf_all(self):
        # folderPath = QtWidgets.QFileDialog.getExistingDirectory()
        # print(folderPath)
        if self.folderPath_dismantle_all[0]:
            try:
                dism_pdf = Pdf.open(self.pdf_count[0]).pages
                # output = Pdf.new() 
                i = 1
                for pdf_ in dism_pdf:
                    output = Pdf.new() 
                    output.pages.append(pdf_)
                    output.save(os.path.join(self.folderPath_dismantle_all,self.comboBox_pdfs.currentText() + '_' + str(i)+'.pdf'))
                    i += 1
                self.label_dismantle_pdf.setText('拆分成功!')
            except:
                self.label_dismantle_pdf.setText('資料夾需新增!')

    # 拆分檔案:範圍頁數
    def check_dismantle_pdf_sub(self):
        self.folderPath_dismantle_sub = QtWidgets.QFileDialog.getExistingDirectory()
        self.lineEdit_check_dismantle_sub.setText(self.folderPath_dismantle_sub)

    def check_dismantle_pdf_sub_chose(self):
        start_page = int(self.lineEdit_dismantle_startpage.text())-1
        end_page = int(self.lineEdit_dismantle_endpage.text())
        self.dism_pdf = Pdf.open(self.pdf_count[0]).pages
        if (start_page < end_page) & ( 0 <= start_page) & (end_page <= len(self.dism_pdf)):
            self.label_check_merge_sub_2.setText('沒問題!小心存放的資料夾')
        else:
            self.label_check_merge_sub_2.setText('有問題!頁數需檢查!')

    def dismantle_pdf_sub(self):
        self.label_dismantle_pdf_2.clear()
        start_page = int(self.lineEdit_dismantle_startpage.text())-1
        end_page = int(self.lineEdit_dismantle_endpage.text())
        # dism_pdf = Pdf.open(self.pdf_count[0]).pages
        # folderPath = QtWidgets.QFileDialog.getExistingDirectory()
        
        if (start_page < end_page) & ( 0 <= start_page) & (end_page <= len(self.dism_pdf)):
            # try:
            # folderPath = QtWidgets.QFileDialog.getExistingDirectory()
            i = 1
            for pdf_ in self.dism_pdf[start_page:end_page]:
                output = Pdf.new() 
                output.pages.append(pdf_)
                output.save(os.path.join(self.folderPath_dismantle_sub,self.comboBox_pdfs.currentText() + '_' + str(i)+'.pdf'))
                i += 1
            self.label_dismantle_pdf_2.setText('拆分成功!')
            # except:
            #     self.label_dismantle_pdf_2.setText('檔案本身有問題!需檢查!')
        else:
            self.label_dismantle_pdf_2.setText('拆分選擇的頁數有誤!請檢查!')

    # 拆分檔案:特定頁數
    def check_dismantle_pdf_distinct(self):
        self.folderPath_dismantle_distinct = QtWidgets.QFileDialog.getExistingDirectory()
        self.lineEdit_check_dismantle_distinct.setText(self.folderPath_dismantle_distinct)

    def check_dismantle_pdf_distinct_chose(self):
        start_pages_list = [ int(str_)-1 for str_ in self.lineEdit_dismantle_distinctpage.text().split()]
        start_page = min(start_pages_list)
        end_page = max(start_pages_list)
        self.dism_pdf = Pdf.open(self.pdf_count[0]).pages
        if (start_page < end_page) & ( 0 <= start_page) & (end_page <= len(self.dism_pdf)):
            self.label_check_merge_distinct_2.setText('沒問題!小心存放的資料夾')
        else:
            self.label_check_merge_distinct_2.setText('有問題!頁數需檢查!')

    def dismantle_pdf_distinct(self):
        # folderPath = QtWidgets.QFileDialog.getExistingDirectory()
        # print(folderPath)
        # start_pages_list = [ int(str_)-1 for str_ in self.lineEdit_dismantle_distinctpage.text().split()]
        start_pages_list = [ int(str_)-1 for str_ in self.lineEdit_dismantle_distinctpage.text().split()]
        start_page = min(start_pages_list)
        end_page = max(start_pages_list)
        if (start_page < end_page) & ( 0 <= start_page) & (end_page <= len(self.dism_pdf)):
            try:
                # dism_pdf = Pdf.open(self.pdf_count[0]).pages
                # output = Pdf.new() 
                i = 1
                for page in start_pages_list:
                    output = Pdf.new() 
                    output.pages.append(self.dism_pdf[page])
                    output.save(os.path.join(self.folderPath_dismantle_distinct,self.comboBox_pdfs.currentText() + '_' + str(i)+'.pdf'))
                    i += 1
                self.label_dismantle_pdf_3.setText('拆分成功!')
            except:
                self.label_dismantle_pdf_3.setText('拆分選擇的頁數有誤!請檢查!')



def main():
    app = QtWidgets.QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec())
 
if __name__ == '__main__':
    main()