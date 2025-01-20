import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QListWidget, QScrollArea,
    QTextEdit, QTableWidget, QTableWidgetItem, QAbstractItemView,
    QHeaderView
)
from PyQt6.QtCore import Qt, QPoint, QUrl
from PyQt6.QtGui import QImage, QPixmap, QPainter, QPen, QMouseEvent, QDragEnterEvent, QDropEvent
import fitz  # PyMuPDF
import pandas as pd

class PDFSelector(QMainWindow):
    """
    A PyQt6 PDF extraction tool with:
      - Drag-and-drop PDF files to load.
      - Sync to All PDFs (copies current PDF's selections to others).
      - Clear All Selections from memory.
      - Table-based selection display (with per-selection Delete).
      - Export with each PDF in one row, selections in columns.
      - Fixed scale=1.0 for consistent coords.
      - A smaller default window size + thinner middle panel.
    """

    def __init__(self):
        super().__init__()
        # Reduced default size: narrower width
        self.setWindowTitle("PDF Extractor (Eason)")
        self.setGeometry(100, 100, 1000, 600)

        # Accept drag-and-drop for PDFs
        self.setAcceptDrops(True)

        # Fixed scale (no zoom)
        self.zoom_factor = 1.0

        # Current PDF info
        self.current_pdf = None
        self.current_pdf_path = ""
        self.current_page_idx = 0

        # PDF file mapping: {filename -> full_path}
        self.pdf_paths = {}

        # Selections across all PDFs:
        #   pdf_selections[full_path] = [ { 'page': int, 'coords': (x1,y1,x2,y2) }, ... ]
        self.pdf_selections = {}

        # For drawing new rectangles
        self.selection_start = None
        self.selection_end = None
        self.temp_selection_rect = None

        # Pixmap of the current page
        self.current_pixmap = None

        # ---------------------- UI Layout ----------------------
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # Left panel: file list + basic nav
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)

        self.pdf_list = QListWidget()
        self.pdf_list.itemClicked.connect(self.handle_pdf_list_click)
        left_layout.addWidget(QLabel("PDF Files:"))
        left_layout.addWidget(self.pdf_list)

        self.load_folder_btn = QPushButton("Open PDF Folder")
        self.load_folder_btn.clicked.connect(self.load_folder)
        left_layout.addWidget(self.load_folder_btn)

        self.prev_page_btn = QPushButton("Previous Page")
        self.prev_page_btn.clicked.connect(self.prev_page)
        left_layout.addWidget(self.prev_page_btn)

        self.next_page_btn = QPushButton("Next Page")
        self.next_page_btn.clicked.connect(self.next_page)
        left_layout.addWidget(self.next_page_btn)

        # Buttons to sync or clear
        self.sync_all_btn = QPushButton("Sync to All PDFs")
        self.sync_all_btn.clicked.connect(self.sync_to_all_pdfs)
        left_layout.addWidget(self.sync_all_btn)

        self.clear_all_btn = QPushButton("Clear All Selections")
        self.clear_all_btn.clicked.connect(self.clear_all_selections)
        left_layout.addWidget(self.clear_all_btn)

        # Export
        self.export_btn = QPushButton("Export Selections")
        self.export_btn.clicked.connect(self.export_all_pdfs)
        left_layout.addWidget(self.export_btn)

        # Middle panel: PDF preview + table
        middle_panel = QWidget()
        middle_layout = QVBoxLayout(middle_panel)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(False)
        self.pdf_label = QLabel()
        self.pdf_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area.setWidget(self.pdf_label)

        self.selections_table = QTableWidget()
        self.selections_table.setColumnCount(3)
        self.selections_table.setHorizontalHeaderLabels(["Page", "Coords", "Action"])
        self.selections_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.selections_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.selections_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.selections_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        middle_layout.addWidget(self.scroll_area, stretch=4)
        middle_layout.addWidget(self.selections_table, stretch=2)

        # Right side: text extraction preview
        self.extraction_preview = QTextEdit()
        self.extraction_preview.setReadOnly(True)

        # --- Set up the layout with adjusted stretch factors ---
        # Suppose we want:
        #   Left panel  = 1
        #   Middle panel = 4  (slightly less wide than before)
        #   Right preview = 1 (thin)
        main_layout.addWidget(left_panel, 1)
        main_layout.addWidget(middle_panel, 4)
        main_layout.addWidget(self.extraction_preview, 1)

        # Mouse events on pdf_label
        self.pdf_label.mousePressEvent = self.mouse_press_event
        self.pdf_label.mouseMoveEvent = self.mouse_move_event
        self.pdf_label.mouseReleaseEvent = self.mouse_release_event

    # --------------------------------------------------------------------------
    # Drag-and-Drop: accept PDF files
    # --------------------------------------------------------------------------
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(".pdf"):
                    file_name = os.path.basename(file_path)
                    if file_name not in self.pdf_paths:
                        self.pdf_list.addItem(file_name)
                    self.pdf_paths[file_name] = file_path
            event.acceptProposedAction()
        else:
            event.ignore()

    # --------------------------------------------------------------------------
    # Folder Loading
    # --------------------------------------------------------------------------
    def load_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.pdf_list.clear()
            self.pdf_paths.clear()
            for file_name in os.listdir(folder_path):
                if file_name.lower().endswith(".pdf"):
                    full_path = os.path.join(folder_path, file_name)
                    self.pdf_list.addItem(file_name)
                    self.pdf_paths[file_name] = full_path

    def handle_pdf_list_click(self, item):
        # Save old PDF's selections
        if self.current_pdf_path:
            self.pdf_selections[self.current_pdf_path] = self.get_current_selections()

        # Load new PDF
        pdf_basename = item.text()
        full_path = self.pdf_paths[pdf_basename]
        self.current_pdf_path = full_path
        self.current_pdf = fitz.open(self.current_pdf_path)
        self.current_page_idx = 0

        # Restore or create an empty list for the new PDF
        if self.current_pdf_path in self.pdf_selections:
            self.set_current_selections(self.pdf_selections[self.current_pdf_path])
        else:
            self.set_current_selections([])

        self.display_page()
        self.update_extraction_preview()

    # --------------------------------------------------------------------------
    # Page Display & Navigation
    # --------------------------------------------------------------------------
    def display_page(self):
        if not self.current_pdf:
            return
        page = self.current_pdf[self.current_page_idx]

        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_factor, self.zoom_factor))
        qimg = QImage(pix.samples, pix.width, pix.height,
                      pix.stride, QImage.Format.Format_RGB888)
        self.current_pixmap = QPixmap.fromImage(qimg)

        self.pdf_label.setPixmap(self.current_pixmap)
        self.pdf_label.resize(self.current_pixmap.size())

        self.update_selection_display()
        self.refresh_selections_table()

    def prev_page(self):
        if self.current_pdf and self.current_page_idx > 0:
            self.current_page_idx -= 1
            self.display_page()
            self.update_extraction_preview()

    def next_page(self):
        if self.current_pdf and self.current_page_idx < len(self.current_pdf) - 1:
            self.current_page_idx += 1
            self.display_page()
            self.update_extraction_preview()

    # --------------------------------------------------------------------------
    # Selections
    # --------------------------------------------------------------------------
    def get_current_selections(self):
        if not hasattr(self, "_current_selections"):
            self._current_selections = []
        return self._current_selections

    def set_current_selections(self, selections):
        self._current_selections = selections

    def clear_all_selections(self):
        self.pdf_selections.clear()
        self.set_current_selections([])
        self.update_selection_display()
        self.update_extraction_preview()
        self.refresh_selections_table()

    def sync_to_all_pdfs(self):
        if not self.current_pdf:
            return

        current_sels = self.get_current_selections()
        if not current_sels:
            return

        self.pdf_selections[self.current_pdf_path] = current_sels
        for pdf_basename, pdf_path in self.pdf_paths.items():
            if pdf_path == self.current_pdf_path:
                continue

            if not os.path.isfile(pdf_path):
                continue

            if pdf_path not in self.pdf_selections:
                self.pdf_selections[pdf_path] = []

            try:
                other_doc = fitz.open(pdf_path)
            except Exception:
                continue

            max_pages = len(other_doc)
            other_sels = self.pdf_selections[pdf_path]

            for sel in current_sels:
                pg = sel['page']
                if pg < max_pages:
                    new_sel = {
                        'page': pg,
                        'coords': sel['coords']
                    }
                    other_sels.append(new_sel)

            other_doc.close()
            self.pdf_selections[pdf_path] = other_sels

        self.display_page()
        self.update_extraction_preview()

    # --------------------------------------------------------------------------
    # Mouse Events
    # --------------------------------------------------------------------------
    def mouse_press_event(self, event: QMouseEvent):
        if not self.current_pixmap:
            return
        if event.button() == Qt.MouseButton.LeftButton:
            x, y = self.get_image_coordinates(event.pos())
            self.selection_start = (x, y)
            self.selection_end = (x, y)
            self.temp_selection_rect = (x, y, x, y)
            self.update_selection_display()

    def mouse_move_event(self, event: QMouseEvent):
        if event.buttons() & Qt.MouseButton.LeftButton and self.selection_start:
            x, y = self.get_image_coordinates(event.pos())
            x1, y1 = self.selection_start
            self.selection_end = (x, y)
            self.temp_selection_rect = (
                min(x1, x), min(y1, y),
                max(x1, x), max(y1, y)
            )
            self.update_selection_display()

    def mouse_release_event(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton and self.selection_start:
            x1, y1 = self.selection_start
            x2, y2 = self.selection_end
            new_sel = {
                'page': self.current_page_idx,
                'coords': (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
            }

            all_sels = self.get_current_selections()
            all_sels.append(new_sel)
            self.set_current_selections(all_sels)
            self.pdf_selections[self.current_pdf_path] = all_sels

            self.selection_start = None
            self.selection_end = None
            self.temp_selection_rect = None

            self.update_selection_display()
            self.update_extraction_preview()
            self.refresh_selections_table()

    # --------------------------------------------------------------------------
    # Table of Selections
    # --------------------------------------------------------------------------
    def refresh_selections_table(self):
        if not self.current_pdf:
            self.selections_table.clearContents()
            self.selections_table.setRowCount(0)
            return

        all_sels = self.get_current_selections()
        self.selections_table.setRowCount(len(all_sels))

        for row_idx, sel in enumerate(all_sels):
            page_str = str(sel['page'] + 1)
            (x1, y1, x2, y2) = sel['coords']
            coords_str = f"({x1},{y1}) - ({x2},{y2})"

            page_item = QTableWidgetItem(page_str)
            page_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            coords_item = QTableWidgetItem(coords_str)
            coords_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.selections_table.setItem(row_idx, 0, page_item)
            self.selections_table.setItem(row_idx, 1, coords_item)

            btn_delete = QPushButton("Delete")
            btn_delete.clicked.connect(lambda _, r=row_idx: self.delete_selection_at_row(r))
            self.selections_table.setCellWidget(row_idx, 2, btn_delete)

        self.selections_table.resizeRowsToContents()

    def delete_selection_at_row(self, row_idx: int):
        all_sels = self.get_current_selections()
        if 0 <= row_idx < len(all_sels):
            all_sels.pop(row_idx)
            self.set_current_selections(all_sels)
            self.pdf_selections[self.current_pdf_path] = all_sels

            self.update_selection_display()
            self.update_extraction_preview()
            self.refresh_selections_table()

    # --------------------------------------------------------------------------
    # Rendering Selections
    # --------------------------------------------------------------------------
    def update_selection_display(self):
        if not self.current_pixmap:
            return

        pixmap = QPixmap(self.current_pixmap)
        painter = QPainter(pixmap)
        painter.setPen(QPen(Qt.GlobalColor.red, 2))

        for sel in self.get_current_selections():
            if sel['page'] == self.current_page_idx:
                x1, y1, x2, y2 = sel['coords']
                painter.drawRect(x1, y1, x2 - x1, y2 - y1)

        if self.temp_selection_rect:
            painter.setPen(QPen(Qt.GlobalColor.blue, 2, Qt.PenStyle.DashLine))
            rx1, ry1, rx2, ry2 = self.temp_selection_rect
            painter.drawRect(rx1, ry1, rx2 - rx1, ry2 - ry1)

        painter.end()
        self.pdf_label.setPixmap(pixmap)

    # --------------------------------------------------------------------------
    # Extraction Preview (Current PDF)
    # --------------------------------------------------------------------------
    def update_extraction_preview(self):
        if not self.current_pdf:
            self.extraction_preview.clear()
            return

        text_preview = ""
        for sel in self.get_current_selections():
            pg = sel['page']
            (x1, y1, x2, y2) = sel['coords']
            rect = fitz.Rect(x1, y1, x2, y2)
            text = self.current_pdf[pg].get_text("text", clip=rect).strip()
            if text:
                text_preview += f"--- Page {pg + 1} ---\n{text}\n\n"

        self.extraction_preview.setPlainText(text_preview.strip())

    # --------------------------------------------------------------------------
    # Coordinate Conversion
    # --------------------------------------------------------------------------
    def get_image_coordinates(self, event_pos: QPoint):
        if not self.current_pixmap:
            return (0, 0)

        lbl_w = self.pdf_label.width()
        lbl_h = self.pdf_label.height()
        pix_w = self.current_pixmap.width()
        pix_h = self.current_pixmap.height()

        offset_x = max((lbl_w - pix_w) // 2, 0)
        offset_y = max((lbl_h - pix_h) // 2, 0)

        x = event_pos.x() - offset_x
        y = event_pos.y() - offset_y

        x = max(0, min(x, pix_w))
        y = max(0, min(y, pix_h))
        return (x, y)

    # --------------------------------------------------------------------------
    # Export Selections in Columns
    # --------------------------------------------------------------------------
    def export_all_pdfs(self):
        if not self.pdf_selections:
            return

        if self.current_pdf_path:
            self.pdf_selections[self.current_pdf_path] = self.get_current_selections()

        pdf_to_texts = {}
        for pdf_path, sel_list in self.pdf_selections.items():
            if not os.path.isfile(pdf_path):
                continue
            try:
                doc = fitz.open(pdf_path)
            except Exception:
                continue

            texts_for_pdf = []
            for sel in sel_list:
                pg = sel['page']
                x1, y1, x2, y2 = sel['coords']
                if pg < len(doc):
                    rect = fitz.Rect(x1, y1, x2, y2)
                    extracted = doc[pg].get_text("text", clip=rect).strip()
                    texts_for_pdf.append(extracted)
                else:
                    texts_for_pdf.append("")

            doc.close()
            pdf_to_texts[pdf_path] = texts_for_pdf

        if not pdf_to_texts:
            return

        max_sel_count = max(len(lst) for lst in pdf_to_texts.values())
        columns = ["PDF Name"] + [f"Selection {i+1}" for i in range(max_sel_count)]

        data_rows = []
        for pdf_path, texts in pdf_to_texts.items():
            pdf_name = os.path.basename(pdf_path)
            row = [pdf_name] + texts
            # pad if needed
            if len(texts) < max_sel_count:
                row.extend([""] * (max_sel_count - len(texts)))
            data_rows.append(row)

        df = pd.DataFrame(data_rows, columns=columns)

        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_name:
            df.to_excel(file_name, index=False)


# ------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFSelector()
    window.show()
    sys.exit(app.exec())
