#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
MIT License

arcdump

Copyright (c) 2022 Control-F Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''

__description__ = 'Control-F - arcdump - Archive Dumper'
__contact__ = 'mike.bangham@controlf.co.uk'

from PyQt5.QtWidgets import (QApplication, QWidget, QMainWindow, QGroupBox,
                             QGridLayout, QTextEdit, QProgressBar, QCheckBox, QDialog, QVBoxLayout,
                             QLabel, QPushButton, QFileDialog)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QThread
import os
from os.path import join as pj
from os.path import abspath, basename
from datetime import datetime
import sys
import tarfile
import zipfile
import time
import math
import tkinter
from tkinter.filedialog import askopenfilename

if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)

if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = pj(abspath("."), 'res')
    return pj(base_path, relative_path)


def button_config(widg, icon_fn):
    widg.setIcon(QIcon(resource_path(icon_fn)))
    widg.setIconSize(QSize(14, 14))
    widg.setFixedWidth(20)
    widg.setFixedHeight(20)


def human_bytes(size):
    units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB']
    power = 0 if size <= 0 else math.floor(math.log(size, 1024))
    return '{} {}'.format(round(size / 1024 ** power, 2), units[int(power)])


class ArchiveExtractionThread(QThread):
    progressSignal = pyqtSignal(str)
    progressbarSignal = pyqtSignal(int)
    finishedSignal = pyqtSignal(str)

    def __init__(self, parent, *args):
        QThread.__init__(self, parent)
        self.params_dict = args[0]
        self.count = 0

    def run(self):
        if zipfile.is_zipfile(self.params_dict['zip_in']):
            with zipfile.ZipFile(self.params_dict['zip_in'], 'r') as zip_in_obj:
                with zipfile.ZipFile(self.params_dict['zip_out'], 'w') as zip_out_obj:
                    if self.params_dict['case']:
                        file_paths = [(f.filename, f.date_time) for f in zip_in_obj.infolist() if
                                      any(kw in f.filename for kw in self.params_dict['keywords'])]
                    else:
                        file_paths = [(f.filename, f.date_time) for f in zip_in_obj.infolist() if
                                      any(kw.lower() in f.filename for kw in self.params_dict['keywords'])]
                    total_files = len(file_paths)
                    for f_tuple in file_paths:
                        self.count += 1
                        info = zipfile.ZipInfo(filename=f_tuple[0], date_time=f_tuple[1])
                        zip_out_obj.writestr(info, zip_in_obj.read(f_tuple[0]))
                        self.progressSignal.emit(basename(f_tuple[0]))
                        self.progressbarSignal.emit(int((self.count / total_files) * 100))

        else:
            with tarfile.open(self.params_dict['zip_in'], 'r') as tar_in_obj:
                with zipfile.ZipFile(self.params_dict['zip_out'], 'w') as zip_out_obj:
                    if self.params_dict['case']:
                        file_paths = [(f, f.name, datetime.fromtimestamp(f.mtime).timetuple()) for f in tar_in_obj.getmembers() if
                                      any(kw in f.name for kw in self.params_dict['keywords'])]
                    else:
                        file_paths = [(f, f.name, datetime.fromtimestamp(f.mtime).timetuple()) for f in tar_in_obj.getmembers() if
                                      any(kw.lower() in f.name for kw in self.params_dict['keywords'])]
                    total_files = len(file_paths)
                    for f_tuple in file_paths:
                        self.count += 1
                        if f_tuple[0].isreg():
                            info = zipfile.ZipInfo(filename=f_tuple[1], date_time=f_tuple[2])
                            zip_out_obj.writestr(info, tar_in_obj.extractfile(f_tuple[1]).read())
                            self.progressSignal.emit(basename(f_tuple[1]))
                            self.progressbarSignal.emit(int((self.count / total_files) * 100))

        self.finishedSignal.emit('Done! Extracted {} objects ({})'.format(total_files,
                                                                          human_bytes(os.stat(
                                                                              self.params_dict['zip_out']).st_size)))


class GUI(QMainWindow):
    def __init__(self):
        super().__init__()
        __version__ = open(resource_path('version'), 'r').readlines()[0]
        self.setWindowTitle('arcdump')
        self.setWindowIcon(QIcon(resource_path('controlF.ico')))
        self.setFixedSize(500, 325)
        window_widget = QWidget(self)
        self.setCentralWidget(window_widget)

        control_f_emblem = QLabel()
        emblem_pixmap = QPixmap(resource_path('ControlF_R_RGB.png')).scaled(200, 200, Qt.KeepAspectRatio,
                                                                            Qt.SmoothTransformation)
        control_f_emblem.setPixmap(emblem_pixmap)

        info_btn = QPushButton()
        button_config(info_btn, "info.png")
        info_btn.clicked.connect(lambda: ShowInfo().exec_())

        self._maingrid = QGridLayout()
        self._maingrid.setContentsMargins(0, 0, 0, 0)
        self._maingrid.addWidget(ExtractFromArchive(self),      0, 0, 1, 4)
        self._maingrid.addWidget(info_btn,                      1, 0, 1, 1, alignment=(Qt.AlignBottom | Qt.AlignLeft))
        self._maingrid.addWidget(control_f_emblem,              1, 1, 1, 3, alignment=(Qt.AlignBottom | Qt.AlignRight))

        window_widget.setLayout(self._maingrid)


class ExtractFromArchive(QWidget):
    def __init__(self, maingui):
        super(ExtractFromArchive, self).__init__(maingui)
        self.maingui = maingui
        self.params_dict = {}

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.addWidget(self._extraction_panel(), 1, 0, 1, 1)
        self.setLayout(grid)

    def _extraction_panel(self):
        groupBox = QGroupBox()
        layout = QGridLayout()

        self.select_archive_btn = QPushButton()
        button_config(self.select_archive_btn, "folder.png")
        layout.addWidget(self.select_archive_btn,       0, 0, 1, 1, alignment=Qt.AlignLeft)
        self.select_archive_btn.clicked.connect(lambda: self.get_file())

        self.archive_btn_lbl = QLabel("Archive")
        layout.addWidget(self.archive_btn_lbl,          0, 1, 1, 1, alignment=Qt.AlignLeft)

        self.extract_btn = QPushButton('Extract')
        layout.addWidget(self.extract_btn,              0, 9, 1, 1, alignment=Qt.AlignRight)
        self.extract_btn.clicked.connect(lambda: self.init_extraction_thread())
        self.extract_btn.hide()

        self.archive_file_lbl = QLabel("Keywords (separated by newlines):")
        layout.addWidget(self.archive_file_lbl,         1, 0, 1, 5, alignment=Qt.AlignLeft)

        self.case_sensitive_cb = QCheckBox('Case Sensitive')
        layout.addWidget(self.case_sensitive_cb,        1, 8, 1, 2, alignment=Qt.AlignRight)

        self.search_input = QTextEdit()
        layout.addWidget(self.search_input,             2, 0, 1, 10, alignment=Qt.AlignLeft)
        self.search_input.setFixedHeight(150)
        self.search_input.setFixedWidth(480)
        self.search_input.setStyleSheet('background-color: #262626; border: 0;')

        self.progress_lbl = QLabel("")
        layout.addWidget(self.progress_lbl, 6, 0, 1, 10)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        layout.addWidget(self.progress_bar, 7, 0, 1, 10)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()

        groupBox.setLayout(layout)
        return groupBox

    def get_file(self):
        self.progress_lbl.setText('')
        tkinter.Tk().withdraw()  # PYQT5 dialog freezes when selecting large zips; tkinter does not
        archive = askopenfilename(title='Select archive', initialdir=os.getcwd())
        if archive:
            if zipfile.is_zipfile(archive) or tarfile.is_tarfile(archive):
                self.params_dict['zip_in'] = archive
                self.params_dict['zip_in_size'] = human_bytes(os.path.getsize(self.params_dict['zip_in']))
                self.archive_btn_lbl.setText('{}  | {}'.format(basename(self.params_dict['zip_in']), self.params_dict['zip_in_size']))
                self.extract_btn.show()
            else:
                self.progress_lbl.setText('[!] File is not a zip/tar archive')

    def get_save_dir(self):
        save_dir = str(QFileDialog.getExistingDirectory(self, "Select Save Directory"))
        if save_dir:
            return save_dir
        else:
            return False

    def init_extraction_thread(self):
        self.progress_lbl.setText('')
        self.params_dict['keywords'] = self.search_input.toPlainText()
        if self.params_dict['keywords']:
            if '\n' in self.params_dict['keywords']:
                self.params_dict['keywords'] = self.search_input.toPlainText().splitlines()

            for widg in [self.select_archive_btn, self.extract_btn, self.search_input, self.case_sensitive_cb]:
                widg.setEnabled(False)

            if self.case_sensitive_cb.isChecked():
                self.params_dict['case'] = True
            else:
                self.params_dict['case'] = True
        else:
            self.progress_lbl.setText('[!] A keyword or list of keywords is missing')

        save_dir = self.get_save_dir()
        if save_dir:
            self.params_dict['zip_out'] = pj(save_dir, 'out_{}.zip'.format(int(time.time())))
            self.progress_bar.show()
            self.extract_archive_thread = ArchiveExtractionThread(self, self.params_dict)
            self.extract_archive_thread.progressSignal.connect(self._archive_thread_progress)
            self.extract_archive_thread.progressbarSignal.connect(self._update_progress_bar)
            self.extract_archive_thread.finishedSignal.connect(self._archive_thread_completed)
            self.extract_archive_thread.start()
        else:
            self.progress_lbl.setText('[!] You must select an output')

    def _update_progress_bar(self, val):
        self.progress_bar.setValue(val)

    def _archive_thread_progress(self, txt):
        self.progress_lbl.setText(txt)

    def _archive_thread_completed(self, s):
        for widg in [self.select_archive_btn, self.extract_btn, self.search_input, self.case_sensitive_cb]:
            widg.setEnabled(True)
        self.extract_btn.hide()
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.progress_lbl.setText(s)


class ShowInfo(QDialog):
    def __init__(self):
        super(ShowInfo, self).__init__()
        self.setFixedSize(550, 300)
        self.about_info()

        layout = QVBoxLayout()
        layout.addWidget(self.gb)
        self.setLayout(layout)

        self.setWindowTitle("About")
        self.setWindowIcon(QIcon(resource_path("controlF.ico")))

    def about_info(self):
        self.gb = QGroupBox("")
        self.layout = QGridLayout()

        chisel_img = QLabel()
        chisel_img_pixmap = QPixmap(resource_path('archive_extract.png')).scaled(150, 150, Qt.KeepAspectRatio,
                                                                                     Qt.SmoothTransformation)
        chisel_img.setPixmap(chisel_img_pixmap)
        self.layout.addWidget(chisel_img, 0, 0, 1, 1)

        self.instructions = QTextEdit()
        self.instructions.setStyleSheet('background-color: #404040; border: 0;')
        self.instructions.setReadOnly(True)
        self.instructions.insertPlainText('arcdump\n\nControl-F \xa9 2022\n\nAuthor: mike.bangham@controlf.co.uk\n\n')
        self.instructions.insertPlainText('arcdump is a forensic forensic extraction tool developed by Control-F.\n\n'
                                          'arcdump will use keywords to extract files from a zip or tar archive; '
                                          'outputting those files to another archive, thereby forensically preserving '
                                          'timestamps and attributes.'
                                          '\n\nIf you would like to report a bug, please contact:\n\ninfo@controlf.net')
        self.layout.addWidget(self.instructions, 0, 1, 1, 1)
        self.gb.setLayout(self.layout)


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(open(resource_path('dark_style.qss')).read())
    ex = GUI()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
