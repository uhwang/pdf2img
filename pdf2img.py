'''
    Extract Images from PDF and sub images
    
    03/18/2025 
'''

import sys, os, io
import fitz  # PyMuPDF
from pathlib import Path
from PIL import Image
import numpy as np
from skimage.morphology import binary_closing

from PyQt5.QtCore import Qt, pyqtSignal, QObject, QProcess, QSize, QBasicTimer
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import ( 
        QApplication, QWidget    , QStyleFactory , QGroupBox, 
        QPushButton , QLineEdit  , QPlainTextEdit, QLineEdit,
        QComboBox   , QGridLayout, QVBoxLayout   , QFileDialog,
        QHBoxLayout , QFormLayout, QFileDialog   , 
        QMessageBox , QLabel     , QCheckBox     ,
        )

import msg
from icons import (
        icon_folder_open  , icon_capture, icon_refresh, 
        icon_copy_src_path, icon_start  , icon_stop,
        icon_setting      , icon_pdf    , icon_copy,
        icon_delete       , icon_copy_src, icon_ocr
        )
        
# https://stackoverflow.com/questions/66636777/how-to-extract-subimages-from-an-image        

def extract_from_table(image, std_thr, kernel_x, kernel_y):

    # Threshold on mean standard deviation in x and y direction
    std_x = np.mean(np.std(image, axis=1), axis=1) > std_thr
    std_y = np.mean(np.std(image, axis=0), axis=1) > std_thr

    # Binary closing to close small whitespaces, e.g. around captions
    std_xx = binary_closing(std_x, np.ones(kernel_x))
    std_yy = binary_closing(std_y, np.ones(kernel_y))

    # Find start and stop positions of each subimage
    start_y = np.where(np.diff(np.int8(std_xx)) == 1)[0]
    stop_y = np.where(np.diff(np.int8(std_xx)) == -1)[0]
    start_x = np.where(np.diff(np.int8(std_yy)) == 1)[0]
    stop_x = np.where(np.diff(np.int8(std_yy)) == -1)[0]

    # Extract subimages
    #return [image[y1:y2, x1:x2, :]
    #        for y1, y2 in zip(start_y, stop_y)
    #        for x1, x2 in zip(start_x, stop_x)]
    return [(y1,y2,x1,x2)
            for y1, y2 in zip(start_y, stop_y)
            for x1, x2 in zip(start_x, stop_x)]

# pdf_path : files or folder        
def extract_images_from_pdf(pdf_path, 
                            output_dir, 
                            start_num, 
                            prefix, 
                            msg,
                            sub_img=False,
                            std_thr=5, 
                            kernel_x=21, 
                            kernel_y=11):
                            
    pdf_name = str(Path(pdf_path).stem)
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        pdf_image_list = page.get_images(full=True)

        for img_index, img_info in enumerate(pdf_image_list):
            xref = img_info[0]
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image_filename = f"{pdf_name}_P{page_num+1}{prefix}{img_index+1}.{image_ext}"
        
            with open(f"{output_dir}/{image_filename}", "wb") as f:
                f.write(image_bytes)
            msg.appendPlainText(f"... Save: {image_filename}")

            if sub_img:
                image_array = np.array(Image.open(io.BytesIO(image_bytes)))
                sub_img_list = extract_from_table(image_array, std_thr, kernel_x, kernel_y)
                msg.appendPlainText("... Start: %d sub image(s) found"%len(sub_img_list))

                #for sub_img_j, sub_img_array in enumerate(sub_img_list):
                for sub_img_j, sub_img_size in enumerate(sub_img_list):
                    filename = f"{pdf_name}_P{page_num+1}{prefix}{img_index+1}_sub_{sub_img_j}.{image_ext}"
                    try:
                        y1, y2 = sub_img_size[0], sub_img_size[1]
                        x1, x2 = sub_img_size[2], sub_img_size[3]
                        Image.fromarray(image_array[y1:y2,x1:x2,:]).save(f"{output_dir}/{filename}")
                    except Exception as e:
                        msg.appendPlainText("... Error(extrac_images_from_pdf): %s"%str(e))
                    msg.appendPlainText(f"... Save:\n    {filename}")
                    
    pdf_document.close()

class PdfToImg(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_list = None
        self.initUI()

    def initUI(self):
        self.form_layout = QFormLayout()
        paper = QGridLayout()
        
        paper.addWidget(QLabel("PDF"), 0,0)
        self.pdf_source = QLineEdit(os.getcwd())
        paper.addWidget(self.pdf_source, 0,1)
        
        self.pdf_source_btn = QPushButton()
        self.pdf_source_btn.setIcon(QIcon(QPixmap(icon_folder_open.table)))
        self.pdf_source_btn.setIconSize(QSize(16,16))
        self.pdf_source_btn.setToolTip("PDF source")
        self.pdf_source_btn.clicked.connect(self.get_pdf_source)
        paper.addWidget(self.pdf_source_btn, 0,2)
        
        paper.addWidget(QLabel("Save"), 1,0)
        self.save_folder = QLineEdit(os.getcwd())
        paper.addWidget(self.save_folder, 1,1)
        
        self.save_folder_btn = QPushButton()
        self.save_folder_btn.setIcon(QIcon(QPixmap(icon_folder_open.table)))
        self.save_folder_btn.setIconSize(QSize(16,16))
        self.save_folder_btn.setToolTip("PDF source")
        self.save_folder_btn.clicked.connect(self.get_new_save_folder)
        paper.addWidget(self.save_folder_btn, 1,2)

        self.copy_source_folder_btn = QPushButton()
        self.copy_source_folder_btn.setIcon(QIcon(QPixmap(icon_copy_src_path.table)))
        self.copy_source_folder_btn.setIconSize(QSize(16,16))
        self.copy_source_folder_btn.setToolTip("Copy PDF source path")
        self.copy_source_folder_btn.clicked.connect(self.copy_pdf_source_path)
        paper.addWidget(self.copy_source_folder_btn, 1,3)
        
        paper.addWidget(QLabel("Prefix"), 2,0)
        self.prefix = QLineEdit("_image_")
        paper.addWidget(self.prefix, 2, 1)
        
        paper.addWidget(QLabel("Start#"), 3,0)
        self.start_number = QLineEdit("0")
        paper.addWidget(self.start_number, 3,1)
        
        crop_group = QGroupBox('Extract Sub Images')
        crop_layout = QGridLayout()
        self.sub_img = QCheckBox()
        crop_layout.addWidget(QLabel("Run"), 0, 0)
        crop_layout.addWidget(self.sub_img, 0, 1)
        
        crop_layout.addWidget(QLabel("Standard Threshold"), 1, 0)
        self.sub_img_std_thr= QLineEdit("5") 
        crop_layout.addWidget(self.sub_img_std_thr, 1, 1)
        
        crop_layout.addWidget(QLabel("Kernel X"), 2, 0)
        self.sub_img_kernel_x = QLineEdit("21") 
        crop_layout.addWidget(self.sub_img_kernel_x, 2, 1)
        
        crop_layout.addWidget(QLabel("Kernel Y"), 3, 0)
        self.sub_img_kernel_y = QLineEdit("11") 
        crop_layout.addWidget(self.sub_img_kernel_y, 3, 1)
        crop_group.setLayout(crop_layout)
        
        bv = QHBoxLayout()
        
        self.start_extract_btn = QPushButton('Start')
        self.start_extract_btn.clicked.connect(self.start_extract)
        
        self.stop_extract_btn = QPushButton('Stop')
        self.stop_extract_btn.clicked.connect(self.stop_extract)

        bv.addWidget(self.start_extract_btn)
        bv.addWidget(self.stop_extract_btn)

        self.message = QPlainTextEdit()
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_message)
        
        self.form_layout.addRow(paper)
        self.form_layout.addWidget(crop_group)
        self.form_layout.addRow(bv)
        self.form_layout.addWidget(self.message)
        self.form_layout.addWidget(self.clear_btn)
        
        self.setLayout(self.form_layout)
        self.setWindowTitle("Pdf2Img")
        self.setWindowIcon(QIcon(QPixmap(icon_capture.table)))
        self.show()
        
    def sub_img_setting_callback(self):
        pass
        
    def copy_pdf_source_path(self):
        self.save_folder.setText(str(Path(self.pdf_source.text()).parent))
        
    def get_pdf_source(self):
        files = QFileDialog.getOpenFileNames(self, "Select PDF", 
        directory=os.getcwd(), 
        filter="PDF (*.pdf);;Images (*.jpg *.jpeg *.png);;All files (*.*)")
        
        npdf = len(files[0])
        
        if npdf:
            self.pdf_list = files[0]
            if npdf == 1:
                self.pdf_source.setText(self.pdf_list[0])
            else:
                self.pdf_source.setText(str(Path(self.pdf_list[0]).parent))
            self.save_folder.setText(str(Path(self.pdf_list[0]).parent))
        else:
            self.pdf_list = None
        
    def start_extract(self):
        if self.pdf_list:
            for pdf in self.pdf_list:
                if pdf.lower().endswith(".pdf"):
                    try:
                        extract_images_from_pdf(
                            pdf, 
                            self.save_folder.text(), 
                            int(self.start_number.text()), 
                            self.prefix.text(),
                            self.message,
                            sub_img=self.sub_img.isChecked(),
                            std_thr  = int(self.sub_img_std_thr.text()),
                            kernel_x = int(self.sub_img_kernel_x.text()),
                            kernel_y = int(self.sub_img_kernel_y.text()))
                    except Exception as e:
                        self.message.appendPlainText("... Error(start_extract:PDF) : %s"%(str(e)))
                elif (self.sub_img.isChecked() and (
                      pdf.lower().endswith(".jpg") or 
                      pdf.lower().endswith(".jpeg") or 
                      pdf.lower().endswith(".png"))):
                    try:
                        self.message.appendPlainText("... Convert image to Numpy array")
                        pf = Path(pdf)
                        image_array = np.array(Image.open(pdf))
                        std_thr  = int(self.sub_img_std_thr.text())
                        kernel_x = int(self.sub_img_kernel_x.text())
                        kernel_y = int(self.sub_img_kernel_y.text())
                        prefix   = self.prefix.text()
                        image_ext= pf.suffix
                        image_name= pf.stem
                        output_dir= self.save_folder.text()
                        sub_img_list = extract_from_table(image_array, std_thr, kernel_x, kernel_y)
                        self.message.appendPlainText("... %d sub image(s) found"%len(sub_img_list))
        
                        if len(sub_img_list) == 0: return

                        #for sub_img_j, sub_img_array in enumerate(sub_img_list):
                        for sub_img_j, sub_img_size in enumerate(sub_img_list):
                            #filename = f"{image_name}{prefix}_sub_{sub_img_j}{image_ext}"
                            filename = f"{image_name}_sub_{sub_img_j}{image_ext}"
                            y1, y2 = sub_img_size[0], sub_img_size[1]
                            x1, x2 = sub_img_size[2], sub_img_size[3]
                            Image.fromarray(image_array[y1:y2,x1:x2,:]).save(f"{output_dir}/{filename}")
                            self.message.appendPlainText("... Save(start_extract:IMG):\n   %s"%filename)
                    except Exception as e:
                        self.message.appendPlainText("... Error(start_extract:IMG): %s"%str(e))
        
    def stop_extract(self):
        pass
        
    def clear_message(self):
        self.message.clear()
        
    def get_new_save_folder(self):
        startingDir = os.getcwd() 
        path = QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, 
        QFileDialog.ShowDirsOnly)
        if not path: return
        self.save_folder.setText(path)
        os.chdir(path)    
        
def run_pdf_to_img():
    
    app = QApplication(sys.argv)

    # --- PyQt4 Only
    #app.setStyle(QStyleFactory.create(u'Motif'))
    #app.setStyle(QStyleFactory.create(u'CDE'))
    #app.setStyle(QStyleFactory.create(u'Plastique'))
    #app.setStyle(QStyleFactory.create(u'Cleanlooks'))
    # --- PyQt4 Only
    
    app.setStyle(QStyleFactory.create("Fusion"))
    pdf = PdfToImg()
    sys.exit(app.exec_())
    
if __name__ == '__main__':
    run_pdf_to_img()    
        