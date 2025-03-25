import sys
import os
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QMessageBox, QFrame, QComboBox, QSpacerItem, QSizePolicy, QTextEdit,
    QProgressBar, QDialog, QFormLayout, QSpinBox, QCheckBox, QHBoxLayout
)
from PyQt5.QtGui import QIcon, QPalette, QColor, QFont, QPixmap
from PyQt5.QtCore import Qt, QPropertyAnimation, QAbstractAnimation, QSize, QThread, pyqtSignal, QSettings
import subprocess

# Globale Farbdefinitionen (Instagram-inspiriert)
BACKGROUND_COLOR = "#FFFFFF"  # Weiß
TEXT_COLOR = "#262626"  # Dunkelgrau
SECONDARY_TEXT_COLOR = "#8E8E8E"  # Grau
ACCENT_COLOR = "#0095F6"  # Instagram Blau
DIVIDER_COLOR = "#DBDBDB"  # Sehr helles Grau

class ConversionThread(QThread):
    message_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    progress_signal = pyqtSignal(int)  # Signal für Fortschritt

    def __init__(self, source_folder, target_folder, file_format, settings):
        super().__init__()
        self.source_folder = source_folder
        self.target_folder = target_folder
        self.file_format = file_format
        self.settings = settings

    def run(self):
        try:
            word = win32com.client.Dispatch("Word.Application")
            files_to_convert = []
            for root, _, files in os.walk(self.source_folder):
                for filename in files:
                    if filename.endswith(".docx") or filename.endswith(".doc"):
                        files_to_convert.append(os.path.join(root, filename))

            total_files = len(files_to_convert)
            for i, file_path in enumerate(files_to_convert):
                rel_path = os.path.relpath(file_path, self.source_folder)
                file_name_without_extension = os.path.splitext(rel_path)[0]

                if self.file_format == 17:  # PDF
                    target_extension = '.pdf'
                    target_format = 17
                elif self.file_format == 0:  # DOC
                    target_extension = '.doc'
                    target_format = 0
                elif self.file_format == 6:  # RTF
                    target_extension = '.rtf'
                    target_format = 6
                elif self.file_format == 2:  # TXT
                    target_extension = '.txt'
                    target_format = 2
                    #ojniubuhbn
                else:
                    target_extension = '.docx'  # DOCX als Standard
                    target_format = 12

                target_path = os.path.join(self.target_folder, file_name_without_extension + target_extension)
                target_dir = os.path.dirname(target_path)
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir)

                self.message_signal.emit(f"Versuche umzuwandeln: {file_path} zu {target_path}")

                try:
                    if os.path.exists(file_path):
                        doc = word.Documents.Open(file_path)

                        # Wenn TXT, zuerst in RTF konvertieren
                        if target_format == 2:
                            rtf_path = os.path.join(self.target_folder, file_name_without_extension + '.rtf')
                            doc.SaveAs(rtf_path, FileFormat=6)  # RTF-Format
                            doc.Close()

                            doc = word.Documents.Open(rtf_path)
                            doc.SaveAs(target_path, FileFormat=target_format)  # TXT-Format
                            doc.Close()
                            os.remove(rtf_path)  # RTF-Datei löschen
                        else:
                            doc.SaveAs(target_path, FileFormat=target_format)
                            doc.Close()
                    else:
                        self.message_signal.emit(f"Nicht gefunden: {file_path}")
                except Exception as e:
                    self.message_signal.emit(f"Fehler beim Konvertieren von {file_path}: {e}")

                # Fortschritt aktualisieren
                progress = int((i + 1) / total_files * 100)
                self.progress_signal.emit(progress)

            word.Quit()
            self.message_signal.emit("Konvertierung abgeschlossen!")
        except Exception as e:
            self.message_signal.emit(f"Unerwarteter Fehler: {e}")
        finally:
            self.finished_signal.emit()

class AnimatedButton(QPushButton):
    def __init__(self, text, button_color, text_color=TEXT_COLOR, parent=None):
        super().__init__(text, parent)
        self.default_stylesheet = f"background-color: {BACKGROUND_COLOR}; color: {button_color}; border-radius: 4px; padding: 8px 12px; font-size: 14px; font-family: 'System-UI'; border: 1px solid {DIVIDER_COLOR};"
        self.hover_stylesheet = f"background-color: {button_color}; color: {BACKGROUND_COLOR}; border-radius: 4px; padding: 8px 12px; font-size: 14px; font-family: 'System-UI'; border: 1px solid {button_color};"
        self.setStyleSheet(self.default_stylesheet)
        self.setCursor(Qt.PointingHandCursor)

    def enterEvent(self, event):
        self.setStyleSheet(self.hover_stylesheet)

    def leaveEvent(self, event):
        self.setStyleSheet(self.default_stylesheet)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Einstellungen")
        self.settings = QSettings("YourCompany", "WordConverter")  # Organisiere Einstellungen

        layout = QFormLayout(self)
        layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        # PDF-Qualitätseinstellung
        self.pdf_quality_label = QLabel("PDF-Qualität:")
        self.pdf_quality_spinbox = QSpinBox()
        self.pdf_quality_spinbox.setRange(50, 100)
        self.pdf_quality_spinbox.setSuffix("%")
        self.pdf_quality_spinbox.setValue(self.settings.value("pdf_quality", 90, type=int))  # Standardwert: 90
        layout.addRow(self.pdf_quality_label, self.pdf_quality_spinbox)

        # TXT-Kodierungseinstellung
        self.txt_encoding_label = QLabel("TXT-Kodierung:")
        self.txt_encoding_combo = QComboBox()
        self.txt_encoding_combo.addItem("UTF-8")
        self.txt_encoding_combo.addItem("Latin-1")
        self.txt_encoding_combo.setCurrentText(self.settings.value("txt_encoding", "UTF-8", type=str))  # Standardwert: UTF-8
        layout.addRow(self.txt_encoding_label, self.txt_encoding_combo)

        # Option zum Überschreiben vorhandener Dateien
        self.overwrite_checkbox = QCheckBox("Vorhandene Dateien überschreiben")
        self.overwrite_checkbox.setChecked(self.settings.value("overwrite_files", False, type=bool))  # Standardwert: False
        layout.addRow(self.overwrite_checkbox)

        # Buttons
        button_layout = QHBoxLayout()
        self.save_button = AnimatedButton("Speichern", ACCENT_COLOR)
        self.cancel_button = AnimatedButton("Abbrechen", ACCENT_COLOR)
        self.save_button.clicked.connect(self.save_settings)
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)
        layout.addRow(button_layout)

        self.setLayout(layout)

    def save_settings(self):
        self.settings.setValue("pdf_quality", self.pdf_quality_spinbox.value())
        self.settings.setValue("txt_encoding", self.txt_encoding_combo.currentText())
        self.settings.setValue("overwrite_files", self.overwrite_checkbox.isChecked())
        self.accept()

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Word Konverter")
        self.setMinimumSize(400, 600)
        self.setStyleSheet(f"background-color: {BACKGROUND_COLOR}; color: {TEXT_COLOR}; font-family: 'System-UI';")

        # Einstellungen laden
        self.settings = QSettings("YourCompany", "WordConverter")

        # Icon
        script_dir = os.path.dirname(os.path.realpath(__file__))
        icon_path = os.path.join(script_dir, "pdf_icon.png")
        self.setWindowIcon(QIcon(icon_path))

        # Layout
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)  # Noch weniger Abstand
        main_layout.setContentsMargins(15, 15, 15, 15)  # Noch weniger Ränder

        # Hinweis-Label
        self.hint_label = QLabel("Nur Word-Dateien (.docx, .doc) werden unterstützt.")
        self.hint_label.setFont(QFont("System-UI", 12))
        self.hint_label.setAlignment(Qt.AlignCenter)
        self.hint_label.setWordWrap(True)
        self.hint_label.setStyleSheet(f"color: {SECONDARY_TEXT_COLOR};")
        main_layout.addWidget(self.hint_label)

        # Buttons
        self.button_select_source = AnimatedButton("Quellordner wählen", ACCENT_COLOR)
        self.button_select_source.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.button_select_source.setFont(QFont("System-UI", 14))
        main_layout.addWidget(self.button_select_source)

        # Pfad Anzeige Label
        self.source_folder_label = QLabel("Kein Ordner ausgewählt")
        self.source_folder_label.setFont(QFont("System-UI", 10))
        self.source_folder_label.setAlignment(Qt.AlignCenter)
        self.source_folder_label.setWordWrap(True)
        self.source_folder_label.setStyleSheet(f"color: {SECONDARY_TEXT_COLOR};")
        main_layout.addWidget(self.source_folder_label)

        # ComboBox
        self.label_select_format = QLabel("Format")
        self.label_select_format.setFont(QFont("System-UI", 14))
        self.label_select_format.setAlignment(Qt.AlignCenter)
        self.label_select_format.setStyleSheet(f"color: {TEXT_COLOR};")
        main_layout.addWidget(self.label_select_format)

        self.format_combo = QComboBox()
        self.format_combo.addItem("PDF", 17)
        self.format_combo.addItem("DOC", 0)
        self.format_combo.addItem("DOCX", 12)
        self.format_combo.addItem("RTF", 6)
        self.format_combo.addItem("TXT", 2)
        self.format_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.format_combo.setFont(QFont("System-UI", 14))
        self.format_combo.setStyleSheet(f"background-color: {BACKGROUND_COLOR}; color: {TEXT_COLOR}; border-radius: 4px; padding: 6px 10px; font-size: 14px; font-family: 'System-UI'; border: 1px solid {DIVIDER_COLOR};")
        main_layout.addWidget(self.format_combo)

        # Fortschrittsbalken
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)  # Initialisiere auf 0 %
        self.progress_bar.setStyleSheet(f"QProgressBar {{ border: 1px solid {DIVIDER_COLOR}; border-radius: 4px; text-align: center; color: {TEXT_COLOR}; }}"
                                         f"QProgressBar::chunk {{ background-color: {ACCENT_COLOR}; border-radius: 4px; }}")
        main_layout.addWidget(self.progress_bar)

        # Status-Textfeld
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setStyleSheet(f"background-color: {BACKGROUND_COLOR}; color: {TEXT_COLOR}; font-size: 12px; font-family: 'System-UI'; border: 1px solid {DIVIDER_COLOR}; border-radius: 4px;")
        main_layout.addWidget(self.status_text)

        # Spacer
        spacer = QSpacerItem(20, 10, QSizePolicy.Minimum, QSizePolicy.Fixed) # Noch weniger Abstand
        main_layout.addItem(spacer)

        # Buttons
        self.start_conversion_button = AnimatedButton("Konvertieren", ACCENT_COLOR)
        self.start_conversion_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.start_conversion_button.setFont(QFont("System-UI", 14))
        main_layout.addWidget(self.start_conversion_button)

        self.settings_button = AnimatedButton("Einstellungen", ACCENT_COLOR)
        self.settings_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.settings_button.setFont(QFont("System-UI", 14))
        main_layout.addWidget(self.settings_button)

        self.setLayout(main_layout)

        self.source_folder = ''
        self.selected_format = 17

        # Verbindungen
        self.button_select_source.clicked.connect(self.select_source_folder)
        self.start_conversion_button.clicked.connect(self.start_conversion)
        self.settings_button.clicked.connect(self.open_settings)

    def select_source_folder(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_filter = "Word Dateien (*.docx *.doc)"
        self.source_folder = QFileDialog.getExistingDirectory(self, "Quellordner auswählen", "", options=options)

        if self.source_folder:
            self.source_folder_label.setText(self.source_folder)
            print(f"Ausgewählter Quellordner: {self.source_folder}")

    def start_conversion(self):
        if not self.source_folder:
            QMessageBox.warning(self, "Warnung", "Bitte Quellordner wählen")
            return

        self.selected_format = self.format_combo.currentData()

        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        target_folder_name = "WordConverter_Output"
        target_folder = os.path.join(downloads_path, target_folder_name)
        os.makedirs(target_folder, exist_ok=True)

        # Konvertierung im Thread starten
        self.conversion_thread = ConversionThread(self.source_folder, target_folder, self.selected_format, self.settings)
        self.conversion_thread.message_signal.connect(self.update_status)
        self.conversion_thread.finished_signal.connect(self.conversion_finished)
        self.conversion_thread.progress_signal.connect(self.update_progress)  # Fortschrittssignal verbinden
        self.conversion_thread.start()

        self.start_conversion_button.setEnabled(False)

    def update_status(self, message):
        self.status_text.append(message)
        QApplication.processEvents()

    def update_progress(self, value):
        self.progress_bar.setValue(value)  # Fortschrittsbalken aktualisieren

    def conversion_finished(self):
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        target_folder_name = "WordConverter_Output"
        target_folder = os.path.join(downloads_path, target_folder_name)

        if sys.platform.startswith('darwin'):  # macOS
            subprocess.Popen(['open', target_folder])
        elif sys.platform.startswith('win'):  # Windows
            os.startfile(target_folder)
        elif sys.platform.startswith('linux'):  # Linux
            subprocess.Popen(['xdg-open', target_folder])
        else:
            QMessageBox.warning(self, "Warnung", "Download-Ordner konnte nicht automatisch geöffnet werden.")

        QMessageBox.information(self, "Info", "Konvertierung abgeschlossen!\nDownload-Ordner geöffnet.")
        QApplication.instance().quit()

        self.start_conversion_button.setEnabled(True)

    def open_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            print("Einstellungen gespeichert")

app = QApplication(sys.argv)

# Dunkles Theme global setzen
app.setStyle("Fusion")
palette = QPalette()
palette.setColor(QPalette.Window, QColor(BACKGROUND_COLOR))
palette.setColor(QPalette.WindowText, QColor(TEXT_COLOR))
palette.setColor(QPalette.Base, QColor(BACKGROUND_COLOR))  # Weiß
palette.setColor(QPalette.Text, QColor(TEXT_COLOR))
palette.setColor(QPalette.Button, QColor(BACKGROUND_COLOR))  # Weiß
palette.setColor(QPalette.ButtonText, QColor(TEXT_COLOR))
palette.setColor(QPalette.Highlight, QColor(ACCENT_COLOR))
palette.setColor(QPalette.HighlightedText, QColor(BACKGROUND_COLOR))
app.setPalette(palette)

# Schriftart global setzen
font = QFont("System-UI")
QApplication.instance().setFont(font)

window = MyWindow()
window.show()

sys.exit(app.exec_())