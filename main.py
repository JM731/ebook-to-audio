from PyQt6.QtWidgets import (QMainWindow,
                             QApplication,
                             QLabel,
                             QWidget,
                             QGridLayout,
                             QPushButton,
                             QFileDialog,
                             QMessageBox,
                             QComboBox,
                             QSlider,
                             QSpinBox)
from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, pyqtSlot, QTimer
from pypdf import PdfReader
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
import pyttsx3


def get_pdf_page_count(pdf_file_path):
    with open(pdf_file_path, "rb") as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        page_count = len(pdf_reader.pages)
        return page_count


class Worker(QObject):
    finished = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.engine = pyttsx3.init()

    @pyqtSlot(dict)
    def do_work(self, data: dict):

        text = ""

        if data["file_extension"] == ".pdf":
            reader = PdfReader(data["file"])
            for i in range(data["initial_page"] - 1, data["final_page"]):
                text += reader.pages[i].extract_text()
        else:
            book = epub.read_epub(data["file"])
            for item in book.get_items():
                if item.get_type() == ebooklib.ITEM_DOCUMENT:
                    content = item.get_content()
                    soup = BeautifulSoup(content, 'html.parser')
                    text += soup.get_text()

        self.engine.setProperty('voice', data["voice"].id)
        self.engine.setProperty('rate', data["rate"])
        self.engine.save_to_file(text, data["file_path"])
        self.engine.runAndWait()

        self.finished.emit()


class Window(QMainWindow):

    work_requested = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        self.file = None
        self.file_extension = None
        self.engine = pyttsx3.init()
        # noinspection PyTypeChecker
        self.voices = {getattr(voice, "name"): voice for voice in self.engine.getProperty("voices")}
        self.timer_ind = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.onTimerTimeout)

        self.setWindowTitle("PDF, EPUB to audio converter")
        self.setContentsMargins(20, 0, 20, 20)

        self.convert_button = QPushButton("Convert")
        self.convert_button.clicked.connect(self.convert)
        self.convert_button.setDisabled(True)
        self.upload_button = QPushButton("Upload")
        self.upload_button.clicked.connect(self.uploadFile)

        self.file_label = QLabel("Upload a PDF file")
        self.file_label.setStyleSheet("margin: 20px;")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.voices_combobox = QComboBox()
        # noinspection PyTypeChecker
        for voice_name in self.voices:
            self.voices_combobox.addItem(voice_name)
        combobox_label = QLabel("Voice")

        self.speech_rate_slider = QSlider()
        self.speech_rate_slider.setOrientation(Qt.Orientation.Horizontal)
        self.speech_rate_slider.setRange(100, 500)
        self.speech_rate_slider.setValue(180)
        self.speech_rate_slider.valueChanged.connect(self.changeSliderLabel)
        self.slider_label = QLabel("Speech rate (180 WPM)")

        self.initial_page_spinbox = QSpinBox()
        self.final_page_spinbox = QSpinBox()
        self.initial_page_spinbox.valueChanged.connect(self.onInitialValueChanged)
        self.initial_page_label = QLabel()
        self.final_page_label = QLabel()
        self.initial_page_label.setText("From page")
        self.final_page_label.setText("To")
        self.initial_page_spinbox.hide()
        self.initial_page_label.hide()
        self.final_page_spinbox.hide()
        self.final_page_label.hide()

        central_widget = QWidget()
        layout = QGridLayout()

        layout.addWidget(QWidget(), 0, 0, 1, 2)
        layout.addWidget(self.file_label, 1, 0, 1, 2)
        layout.addWidget(self.upload_button, 2, 0, 1, 1)
        layout.addWidget(self.convert_button, 2, 1, 1, 1)
        layout.addWidget(combobox_label, 3, 0, 1, 1)
        layout.addWidget(self.slider_label, 3, 1, 1, 1)
        layout.addWidget(self.voices_combobox, 4, 0, 1, 1)
        layout.addWidget(self.speech_rate_slider, 4, 1, 1, 1)
        layout.addWidget(self.initial_page_label, 5, 0, 1, 1)
        layout.addWidget(self.final_page_label, 5, 1, 1, 1)
        layout.addWidget(self.initial_page_spinbox, 6, 0, 1, 1)
        layout.addWidget(self.final_page_spinbox, 6, 1, 1, 1)
        layout.addWidget(QWidget(), 7, 1, 1, 2)

        layout.setRowStretch(0, 1)
        layout.setRowStretch(7, 1)

        self.worker = Worker()
        self.worker_thread = QThread()
        self.work_requested.connect(self.worker.do_work)
        self.worker.finished.connect(self.onConversionFinished)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.start()

        central_widget.setLayout(layout)

        self.setCentralWidget(central_widget)

        self.show()

    def uploadFile(self):
        file_name, _ = QFileDialog.getOpenFileName(self,
                                                   "Select a file",
                                                   "",
                                                   "PDF, EPUB file (*.pdf *.epub);;All Files (*)")
        if file_name:
            file_extension = file_name[file_name.rfind("."):].lower()

            if file_extension == ".pdf" or file_extension == ".epub":
                self.file = file_name
                self.file_extension = file_extension
                self.file_label.setText(file_name)

                if self.file_extension == ".pdf":
                    num_pages = get_pdf_page_count(self.file)
                    self.initial_page_spinbox.setRange(1, num_pages)
                    self.onInitialValueChanged()
                    self.initial_page_label.show()
                    self.initial_page_spinbox.show()
                    self.final_page_label.show()
                    self.final_page_spinbox.show()

                else:
                    self.initial_page_label.hide()
                    self.initial_page_spinbox.hide()
                    self.final_page_label.hide()
                    self.final_page_spinbox.hide()
                    QApplication.processEvents()

                self.convert_button.setDisabled(False)

            else:
                self.invalidFileMessage("Please select a valid file.")

        self.adjustSize()

    def invalidFileMessage(self, text):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setWindowTitle("Invalid File Format")
        msg_box.setText(text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    def convert(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Audio", "", "WAV (*.wav)")

        if file_path:

            data = {
                "file_extension": self.file_extension,
                "file": self.file,
                "file_path": file_path,
                "voice": self.voices[self.voices_combobox.currentText()],
                "rate": self.speech_rate_slider.value(),
                "initial_page": self.initial_page_spinbox.value(),
                "final_page": self.final_page_spinbox.value()
            }

            self.work_requested.emit(data)
            self.timer.start(500)
            self.convert_button.setDisabled(True)
            self.upload_button.setDisabled(True)

    def onConversionFinished(self):
        self.timer.stop()
        self.timer_ind = 1
        self.upload_button.setDisabled(False)
        self.convert_button.setDisabled(False)
        self.file_label.setText(f"{self.file}\nDone!")

    def changeSliderLabel(self):
        self.slider_label.setText(f"Speech Rate ({self.speech_rate_slider.value()} WPM)")

    def onInitialValueChanged(self):
        self.final_page_spinbox.setRange(self.initial_page_spinbox.value(), self.initial_page_spinbox.maximum())

    def onTimerTimeout(self):
        self.file_label.setText(f"{self.file}\nConverting, please wait{'.'*self.timer_ind}")
        self.timer_ind += 1
        if self.timer_ind == 4:
            self.timer_ind = 0


if __name__ == "__main__":
    main_event_thread = QApplication([])
    window = Window()
    main_event_thread.exec()
