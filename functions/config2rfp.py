from PyQt5.QtWidgets import QApplication, QSplashScreen, QVBoxLayout, QWidget, QTextEdit
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
import sys
import time
from PyQt5.QtGui import QTextCursor

class WriteStream:
    def __init__(self, text_written_signal):
        self.text_written_signal = text_written_signal

    def write(self, text):
        self.text_written_signal.emit(text)

    def flush(self):
        pass

class MessageWindow(QWidget):
    text_written_signal = pyqtSignal(str)

    def __init__(self, title="2RFP", icon_path="assets/icon.png"):
        super().__init__()

        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(icon_path))

        self.message_label = QTextEdit(self)
        self.message_label.setReadOnly(True)
        self.message_label.setStyleSheet(
            "background-color: black; color: white; font-size: 16px;"
        )

        layout = QVBoxLayout(self)
        layout.addWidget(self.message_label)

        self.is_closed = False

        self.text_written_signal.connect(self.write_text)
        
    def closeEvent(self, event):
        self.is_closed = True
        event.accept() 

    def write_text(self, text):
        if not self.is_closed:
            cursor = self.message_label.textCursor()
            cursor.movePosition(QTextCursor.End)
            cursor.insertText(text)
            self.message_label.setTextCursor(cursor)
            self.message_label.ensureCursorVisible()
        self.setFixedSize(720, 720)

def redirect_output(text_written_signal):
    sys.stdout = WriteStream(text_written_signal)
    sys.stderr = WriteStream(text_written_signal)

def initSplash():
    app = QApplication(sys.argv)

    splash_pix = QPixmap('assets/splash-screen.png')
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.show()

    message_window = MessageWindow(title="2RFP", icon_path="assets/icon.png")

    write_stream = WriteStream(message_window.text_written_signal)
    sys.stdout = write_stream
    sys.stderr = write_stream

    app.processEvents()
    time.sleep(3)
    splash.close()
    message_window.show()


    return app, message_window