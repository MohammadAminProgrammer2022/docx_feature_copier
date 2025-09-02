# main.py
# Dark-mode GUI (PySide6) for copying layout/styles between DOCX files using Word 2013+ (COM).
# Features: Template/Report pickers, Save As…, live log, progress bar, Help, section-map & Show Word UI toggles.
# Stable: Fusion style applied AFTER QApplication creation, no deprecated HDPI attribute, dark dialogs.

import os
import sys
import tempfile
import shutil
import zipfile
import traceback

from PySide6 import QtCore, QtGui, QtWidgets

# ---------- Word/COM + OpenXML logic ----------
import pythoncom
from win32com.client import Dispatch
import xml.etree.ElementTree as ET

ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ---- Helpers ----
def save_docx(doc, path):
    """Save a Document to .docx, trying SaveCopyAs -> SaveAs2 -> SaveAs (Word 2013-safe)."""
    dir_ = os.path.dirname(path)
    if dir_ and not os.path.isdir(dir_):
        os.makedirs(dir_, exist_ok=True)
    try:
        doc.SaveCopyAs(FileName=path, FileFormat=12)  # .docx
        return
    except Exception:
        pass
    try:
        doc.SaveAs2(FileName=path, FileFormat=12, AddToRecentFiles=False)
        return
    except Exception:
        pass
    doc.SaveAs(FileName=path, FileFormat=12, AddToRecentFiles=False)

def copy_page_setup(src_sec, dst_sec):
    sps, dps = src_sec.PageSetup, dst_sec.PageSetup
    dps.TopMargin      = sps.TopMargin
    dps.BottomMargin   = sps.BottomMargin
    dps.LeftMargin     = sps.LeftMargin
    dps.RightMargin    = sps.RightMargin
    dps.Gutter         = sps.Gutter
    dps.HeaderDistance = sps.HeaderDistance
    dps.FooterDistance = sps.FooterDistance
    dps.PageWidth      = sps.PageWidth
    dps.PageHeight     = sps.PageHeight
    dps.Orientation    = sps.Orientation
    dps.DifferentFirstPageHeaderFooter = sps.DifferentFirstPageHeaderFooter
    dps.OddAndEvenPagesHeaderFooter    = sps.OddAndEvenPagesHeaderFooter
    for prop in ("MirrorMargins", "TwoPagesOnOne"):
        try: setattr(dps, prop, getattr(sps, prop))
        except Exception: pass

def copy_page_borders_basic(src_sec, dst_sec):
    """Basic line borders via COM. Artistic borders are patched via OpenXML later."""
    sbd, dbd = src_sec.Borders, dst_sec.Borders
    for prop in ("Enable", "DistanceFrom", "SurroundHeader", "SurroundFooter",
                 "JoinBorders", "AlwaysInFront"):
        try: setattr(dbd, prop, getattr(sbd, prop))
        except Exception: pass
    for prop in ("ArtStyle", "ArtWidth"):
        try: setattr(dbd, prop, getattr(sbd, prop))
        except Exception: pass
    for idx in (1, 2, 3, 4):  # 1=Top, 2=Left, 3=Bottom, 4=Right
        try:
            sb, db = sbd(idx), dbd(idx)
            db.LineStyle = sb.LineStyle
            db.LineWidth = sb.LineWidth
            db.Color     = sb.Color
            for p in ("DistanceFromTop","DistanceFromBottom","DistanceFromLeft","DistanceFromRight"):
                try: setattr(db, p, getattr(sb, p))
                except Exception: pass
        except Exception:
            pass

def _copy_single_hf(src_hf, dst_hf):
    try: dst_hf.LinkToPrevious = False
    except Exception: pass
    rng = dst_hf.Range
    rng.Delete()
    try:
        rng.FormattedText = src_hf.Range.FormattedText
    except Exception:
        rng.Text = src_hf.Range.Text

def copy_headers_footers(src_sec, dst_sec):
    for t in (1, 2, 3):  # 1=Primary, 2=First, 3=Even
        try: _copy_single_hf(src_sec.Headers(t), dst_sec.Headers(t))
        except Exception: pass
        try: _copy_single_hf(src_sec.Footers(t), dst_sec.Footers(t))
        except Exception: pass

def try_organizer_copy_all_styles(src_doc, dst_doc_path):
    WD_ORGANIZER_OBJECT_STYLES = 3
    app = src_doc.Application
    moved = 0
    for st in src_doc.Styles:
        name = None
        for attr in ("NameLocal","Name"):
            try:
                name = getattr(st, attr)
                if name: break
            except Exception:
                pass
        if not name: continue
        try:
            app.OrganizerCopy(Source=src_doc.FullName,
                              Destination=dst_doc_path,
                              Name=name,
                              Object=WD_ORGANIZER_OBJECT_STYLES)
            moved += 1
        except Exception:
            pass
    return moved

def copy_styles_via_template(work_doc, src_doc):
    tmp_dotx = os.path.join(tempfile.gettempdir(), "style_source_tmp.dotx")
    try:
        try:
            src_doc.SaveCopyAs(FileName=tmp_dotx, FileFormat=16)  # .dotx
        except Exception:
            save_docx(src_doc, tmp_dotx)  # ensure exists
        work_doc.CopyStylesFromTemplate(tmp_dotx)
    finally:
        try: os.remove(tmp_dotx)
        except Exception: pass

# --- OpenXML helpers for artistic page borders ---
def _deepcopy(elem):
    new = ET.Element(elem.tag, elem.attrib)
    for child in list(elem):
        new.append(_deepcopy(child))
    new.text = elem.text
    new.tail = elem.tail
    return new

def _extract_pgBorders_from_source(source_docx):
    with zipfile.ZipFile(source_docx, "r") as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    # Last body-level sectPr preferred
    sectprs = root.findall(f".//{W_NS}body/{W_NS}sectPr")
    if sectprs:
        pg = sectprs[-1].find(f"{W_NS}pgBorders")
        if pg is not None:
            return pg
    # Fallback: last paragraph-level sectPr
    sectprs = root.findall(f".//{W_NS}p/{W_NS}pPr/{W_NS}sectPr")
    if sectprs:
        pg = sectprs[-1].find(f"{W_NS}pgBorders")
        if pg is not None:
            return pg
    return None

def _insert_pgBorders_schema_order(sectpr, pgBorders_elem):
    """
    Insert <w:pgBorders> in schema-valid position:
      AFTER  last of:  w:type, w:pgSz, w:pgMar, w:paperSrc
      BEFORE first of: w:lnNumType, w:pgNumType, w:cols, w:docGrid
    """
    existing = sectpr.find(f"{W_NS}pgBorders")
    if existing is not None:
        sectpr.remove(existing)

    children = list(sectpr)

    after_tags  = {f"{W_NS}type", f"{W_NS}pgSz", f"{W_NS}pgMar", f"{W_NS}paperSrc"}
    before_tags = {f"{W_NS}lnNumType", f"{W_NS}pgNumType", f"{W_NS}cols", f"{W_NS}docGrid"}

    after_idx = -1
    for i, ch in enumerate(children):
        if ch.tag in after_tags and i > after_idx:
            after_idx = i

    before_idx = len(children)
    for i, ch in enumerate(children):
        if ch.tag in before_tags:
            before_idx = i
            break

    insert_idx = min(after_idx + 1, before_idx)
    sectpr.insert(insert_idx, _deepcopy(pgBorders_elem))

def _set_pgBorders_in_all_sections(target_docx, pgBorders_elem):
    with zipfile.ZipFile(target_docx, "r") as zin:
        xml = zin.read("word/document.xml")
    root = ET.fromstring(xml)

    changed = False
    for sectpr in root.findall(f".//{W_NS}body/{W_NS}sectPr"):
        _insert_pgBorders_schema_order(sectpr, pgBorders_elem)
        changed = True
    for sectpr in root.findall(f".//{W_NS}p/{W_NS}pPr/{W_NS}sectPr"):
        _insert_pgBorders_schema_order(sectpr, pgBorders_elem)
        changed = True

    if not changed:
        return False

    new_docx = target_docx + ".tmp"
    with zipfile.ZipFile(target_docx, "r") as zin, \
         zipfile.ZipFile(new_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            zout.writestr(item, data)
    os.replace(new_docx, target_docx)
    return True

def patch_art_page_borders(source_docx, output_docx):
    pg = _extract_pgBorders_from_source(source_docx)
    if pg is None:
        return False
    return _set_pgBorders_in_all_sections(output_docx, pg)

def transfer_layout(source_docx, target_docx, output_docx, visible=False, section_map=False, log=lambda m: None):
    # Normalize & validate
    source_docx = os.path.abspath(os.path.expanduser(source_docx))
    target_docx = os.path.abspath(os.path.expanduser(target_docx))
    output_docx = os.path.abspath(os.path.expanduser(output_docx))

    if not os.path.isfile(source_docx): raise FileNotFoundError(f"Source not found: {source_docx}")
    if not os.path.isfile(target_docx): raise FileNotFoundError(f"Target not found: {target_docx}")

    log(f"[INFO] SOURCE: {source_docx}")
    log(f"[INFO] TARGET: {target_docx}")
    log(f"[INFO] OUTPUT: {output_docx}")

    pythoncom.CoInitialize()
    word = Dispatch("Word.Application")
    word.Visible = bool(visible)
    try:
        try:
            word.DisplayAlerts = 0
            word.ScreenUpdating = False
            word.EnableEvents = False
        except Exception: pass

        log("[INFO] Opening documents in Word…")
        src = word.Documents.Open(FileName=source_docx, ReadOnly=True,
                                  AddToRecentFiles=False, ConfirmConversions=False,
                                  Revert=False, Visible=False, OpenAndRepair=False,
                                  NoEncodingDialog=True)
        tgt = word.Documents.Open(FileName=target_docx, ReadOnly=False,
                                  AddToRecentFiles=False, ConfirmConversions=False,
                                  Revert=False, Visible=False, OpenAndRepair=False,
                                  NoEncodingDialog=True)

        # Save target as output (so Organizer has a real path)
        log("[INFO] Creating working copy…")
        save_docx(tgt, output_docx)

        tgt.Close(False)
        work = word.Documents.Open(FileName=output_docx, ReadOnly=False,
                                   AddToRecentFiles=False, ConfirmConversions=False,
                                   Revert=False, Visible=False, OpenAndRepair=False,
                                   NoEncodingDialog=True)

        try:
            # Styles first
            log("[INFO] Copying styles (Organizer)…")
            moved = try_organizer_copy_all_styles(src, output_docx)
            log(f"[INFO] Styles moved via Organizer: {moved}")
            if moved == 0:
                log("[WARN] Organizer moved 0 styles. Using CopyStylesFromTemplate fallback…")
                copy_styles_via_template(work, src)
                log("[INFO] Styles copied via template.")

            # Layout
            log("[INFO] Copying layout…")
            src_count, tgt_count = src.Sections.Count, work.Sections.Count
            for i in range(1, tgt_count + 1):
                s_idx = i if (section_map and i <= src_count) else 1
                log(f"  - Applying SOURCE Section({s_idx}) -> OUTPUT Section({i})")
                copy_page_setup(src.Sections(s_idx), work.Sections(i))
                copy_page_borders_basic(src.Sections(s_idx), work.Sections(i))
                copy_headers_footers(src.Sections(s_idx), work.Sections(i))

            log("[INFO] Saving before XML patch…")
            try:
                work.Save()
            except Exception:
                tmp = os.path.join(tempfile.gettempdir(), "docx_pre_patch.docx")
                save_docx(work, tmp)
                shutil.copy2(tmp, output_docx)

        finally:
            log("[INFO] Closing COM docs…")
            try: src.Close(False)
            except Exception: pass
            try: work.Close(False)
            except Exception: pass

    finally:
        log("[INFO] Quitting Word COM.")
        try: word.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()

    # OpenXML artistic border patch
    log("[INFO] Patching artistic page borders via OpenXML…")
    patched = patch_art_page_borders(source_docx, output_docx)
    if patched:
        log("[SUCCESS] Artistic page borders patched.")
    else:
        log("[WARN] No w:pgBorders found in source; nothing to patch.")

    return output_docx

# ---------- Dark MessageBox helper ----------
class DarkMessageBox(QtWidgets.QMessageBox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.setStyleSheet("""
            QMessageBox {
                background-color: #24272D;
                color: #E8E8E8;
            }
            QLabel { color: #E8E8E8; }
            QPushButton {
                background: #3C4048;
                color: #E8E8E8;
                border: 1px solid #4A4F58;
                border-radius: 8px;
                padding: 6px 12px;
            }
            QPushButton:hover { background: #4A4F58; }
            QPushButton:pressed { background: #2F3340; }
        """)

def show_info(parent, title, text):
    dlg = DarkMessageBox(parent); dlg.setIcon(DarkMessageBox.Information)
    dlg.setWindowTitle(title); dlg.setText(text); return dlg.exec()

def show_warning(parent, title, text):
    dlg = DarkMessageBox(parent); dlg.setIcon(DarkMessageBox.Warning)
    dlg.setWindowTitle(title); dlg.setText(text); return dlg.exec()

def show_error(parent, title, text):
    dlg = DarkMessageBox(parent); dlg.setIcon(DarkMessageBox.Critical)
    dlg.setWindowTitle(title); dlg.setText(text); return dlg.exec()

# ---------- Worker (QThread) ----------
class TransferWorker(QtCore.QThread):
    progressed = QtCore.Signal(str)
    finishedOk = QtCore.Signal(str)
    failed = QtCore.Signal(str)

    def __init__(self, template, report, output, section_map, show_ui, parent=None):
        super().__init__(parent)
        self.template = template
        self.report = report
        self.output = output
        self.section_map = section_map
        self.show_ui = show_ui

    def run(self):
        try:
            def log(msg): self.progressed.emit(msg)
            out = transfer_layout(
                self.template, self.report, self.output,
                visible=self.show_ui,
                section_map=self.section_map,
                log=log
            )
            self.finishedOk.emit(out)
        except Exception as e:
            tb = traceback.format_exc()
            self.failed.emit(f"{e}\n\n{tb}")

# ---------- Main Window (Dark Mode) ----------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DOCX Layout Copier | TF-Dena AI Section | version 1.0")
        self.resize(780, 580)
        self.setMinimumSize(720, 540)

        # Apply dark palette + styles (no global style/attribute calls here)
        self.apply_dark_theme()

        # Central widget
        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)

        # Inputs
        self.templateEdit = QtWidgets.QLineEdit()
        self.templateEdit.setPlaceholderText("Select template (.docx)…")
        self.reportEdit = QtWidgets.QLineEdit()
        self.reportEdit.setPlaceholderText("Select report (.docx)…")

        self.templateBtn = self.make_browse_button("Browse Template")
        self.reportBtn = self.make_browse_button("Browse Report")

        self.sectionMapChk = QtWidgets.QCheckBox("Map sections 1→1 (source Section i → target Section i)")
        self.showUiChk = QtWidgets.QCheckBox("Show Word UI (may be slower)")

        # Help button
        self.helpBtn = QtWidgets.QPushButton("Help")
        self.helpBtn.setCursor(QtCore.Qt.PointingHandCursor)
        self.helpBtn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_MessageBoxInformation))
        self.helpBtn.clicked.connect(self.show_help)

        # Action buttons
        self.runBtn = self.make_run_button("Run")
        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0, 0)  # busy
        self.progress.setVisible(False)

        # Log
        self.logView = QtWidgets.QPlainTextEdit()
        self.logView.setReadOnly(True)
        self.logView.setWordWrapMode(QtGui.QTextOption.NoWrap)
        self.logView.setStyleSheet("font-family: Consolas, Menlo, monospace; font-size: 12px;")

        # Layouts
        grid = QtWidgets.QGridLayout()
        grid.addWidget(self.label("Template (.docx):"), 0, 0)
        grid.addWidget(self.templateEdit, 0, 1)
        grid.addWidget(self.templateBtn, 0, 2)

        grid.addWidget(self.label("Report (.docx):"), 1, 0)
        grid.addWidget(self.reportEdit, 1, 1)
        grid.addWidget(self.reportBtn, 1, 2)

        # Row for checkboxes + help
        chkRow = QtWidgets.QHBoxLayout()
        chkRow.addWidget(self.sectionMapChk)
        chkRow.addSpacing(10)
        chkRow.addWidget(self.showUiChk)
        chkRow.addStretch(1)
        chkRow.addWidget(self.helpBtn)

        topBox = QtWidgets.QGroupBox("Inputs")
        topLay = QtWidgets.QVBoxLayout(topBox)
        topLay.addLayout(grid)
        topLay.addLayout(chkRow)

        btnRow = QtWidgets.QHBoxLayout()
        btnRow.addWidget(self.progress, 1)
        btnRow.addStretch(1)
        btnRow.addWidget(self.runBtn, 0)

        v = QtWidgets.QVBoxLayout(cw)
        v.addWidget(topBox)
        v.addSpacing(10)
        v.addWidget(self.label("Log:"))
        v.addWidget(self.logView, 1)
        v.addLayout(btnRow)

        # Signals
        self.templateBtn.clicked.connect(self.pickTemplate)
        self.reportBtn.clicked.connect(self.pickReport)
        self.runBtn.clicked.connect(self.onRun)

    # ---- UI helpers ----
    def apply_dark_theme(self):
        palette = QtGui.QPalette()
        bg = QtGui.QColor(27, 29, 33)
        card = QtGui.QColor(36, 39, 45)
        text = QtGui.QColor(232, 232, 232)
        mid = QtGui.QColor(60, 64, 72)
        acc = QtGui.QColor(88, 101, 242)  # soft indigo

        palette.setColor(QtGui.QPalette.Window, bg)
        palette.setColor(QtGui.QPalette.Base, card)
        palette.setColor(QtGui.QPalette.AlternateBase, bg)
        palette.setColor(QtGui.QPalette.Button, card)
        palette.setColor(QtGui.QPalette.Text, text)
        palette.setColor(QtGui.QPalette.ButtonText, text)
        palette.setColor(QtGui.QPalette.WindowText, text)
        palette.setColor(QtGui.QPalette.ToolTipBase, card)
        palette.setColor(QtGui.QPalette.ToolTipText, text)
        palette.setColor(QtGui.QPalette.Highlight, acc)
        palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
        palette.setColor(QtGui.QPalette.PlaceholderText, QtGui.QColor(160, 160, 160))

        QtWidgets.QApplication.setPalette(palette)
        self.setPalette(palette)

        self.setStyleSheet(f"""
            QMainWindow {{ background: {bg.name()}; }}
            QDialog {{ background: {bg.name()}; color: {text.name()}; }}
            QMessageBox {{
                background: {bg.name()};
                color: {text.name()};
            }}
            QFileDialog QWidget {{
                background: {bg.name()};
                color: {text.name()};
            }}
            QGroupBox {{
                font-weight: 600;
                border: 1px solid {mid.name()};
                border-radius: 12px;
                margin-top: 8px;
                padding: 12px;
                background: {card.name()};
                color: {text.name()};
            }}
            QLabel {{ color: {text.name()}; }}
            QLineEdit {{
                padding: 8px 10px;
                border: 1px solid {mid.name()};
                border-radius: 10px;
                background: {bg.name()};
                color: {text.name()};
            }}
            QPlainTextEdit {{
                border: 1px solid {mid.name()};
                border-radius: 10px;
                background: {bg.name()};
                color: {text.name()};
            }}
            QCheckBox {{ color: {text.name()}; }}
            QProgressBar {{
                height: 18px;
                border: 1px solid {mid.name()};
                border-radius: 9px;
                background: {bg.name()};
                color: {text.name()};
                text-align: center;
            }}
            QProgressBar::chunk {{ background: {acc.name()}; border-radius: 7px; }}
            QPushButton {{
                border: 1px solid {mid.name()};
                border-radius: 12px;
                padding: 10px 16px;
                background: {card.name()};
                color: {text.name()};
            }}
            QPushButton:hover {{ background: {mid.name()}; }}
            QPushButton:pressed {{ background: {acc.darker(140).name()}; }}
            QToolTip {{
                background: {card.name()};
                color: {text.name()};
                border: 1px solid {mid.name()};
            }}
        """)

    def label(self, text):
        lbl = QtWidgets.QLabel(text)
        f = lbl.font()
        f.setPointSize(10)
        lbl.setFont(f)
        return lbl

    def make_browse_button(self, text):
        btn = QtWidgets.QPushButton(text)
        btn.setCursor(QtCore.Qt.PointingHandCursor)
        btn.setMinimumWidth(150)
        btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DirOpenIcon))
        return btn

    def make_run_button(self, text):
        btn = QtWidgets.QPushButton(text)
        btn.setCursor(QtCore.Qt.PointingHandCursor)
        f = btn.font()
        f.setPointSize(11); f.setBold(True); btn.setFont(f)
        btn.setMinimumWidth(140); btn.setFixedHeight(44)
        btn.setStyleSheet(btn.styleSheet() + """
            QPushButton { background: #5865F2; border: none; }
            QPushButton:hover { background: #6C78F4; }
            QPushButton:pressed { background: #4754C4; }
        """)
        btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_MediaPlay))
        return btn

    # ---- Actions ----
    def pickTemplate(self):
        # Force Qt (non-native) dialog to ensure dark palette is respected
        opts = QtWidgets.QFileDialog.Options()
        opts |= QtWidgets.QFileDialog.DontUseNativeDialog
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Choose Template (.docx)", "", "Word Document (*.docx)", options=opts)
        if path:
            self.templateEdit.setText(path)

    def pickReport(self):
        opts = QtWidgets.QFileDialog.Options()
        opts |= QtWidgets.QFileDialog.DontUseNativeDialog
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Choose Report (.docx)", "", "Word Document (*.docx)", options=opts)
        if path:
            self.reportEdit.setText(path)

    def show_help(self):
        text = (
            "<b>Map sections 1→1</b><br>"
            "Copies layout from <i>Template Section i</i> to <i>Report Section i</i>.<br>"
            "If the report has more sections than the template, the template’s last section is reused.<br>"
            "If unchecked, Template Section(1) is applied to <i>all</i> sections in the report.<br><br>"
            "<b>Show Word UI</b><br>"
            "Makes Microsoft Word visible during the operation. Handy for troubleshooting, but slightly slower. "
            "Avoid clicking around in Word while it runs."
        )
        show_info(self, "Help", text)
    
    
    @staticmethod
    def resource_path(relpath: str) -> str:
        # works both in dev and in PyInstaller onefile
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, relpath)
        return relpath

    def onRun(self):
        template = self.templateEdit.text().strip()
        report = self.reportEdit.text().strip()
        if not template or not os.path.isfile(template):
            show_warning(self, "Missing file", "Please choose a valid Template (.docx).")
            return
        if not report or not os.path.isfile(report):
            show_warning(self, "Missing file", "Please choose a valid Report (.docx).")
            return

        # Ask for output path (force Qt dialog for dark theme)
        opts = QtWidgets.QFileDialog.Options()
        opts |= QtWidgets.QFileDialog.DontUseNativeDialog
        default_name = os.path.splitext(os.path.basename(report))[0] + "_with_layout.docx"
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Result As…",
            os.path.join(os.path.dirname(report), default_name),
            "Word Document (*.docx)",
            options=opts
        )
        if not out_path:
            return

        # Prep UI
        self.logView.clear()
        self.progress.setVisible(True)
        self.runBtn.setEnabled(False)
        self.templateBtn.setEnabled(False)
        self.reportBtn.setEnabled(False)

        # Launch worker
        self.worker = TransferWorker(
            template=template,
            report=report,
            output=out_path,
            section_map=self.sectionMapChk.isChecked(),
            show_ui=self.showUiChk.isChecked()
        )
        self.worker.progressed.connect(self.appendLog)
        self.worker.finishedOk.connect(self.onDone)
        self.worker.failed.connect(self.onFail)
        self.worker.start()

    @QtCore.Slot(str)
    def appendLog(self, msg):
        self.logView.appendPlainText(msg)
        self.logView.verticalScrollBar().setValue(self.logView.verticalScrollBar().maximum())

    @QtCore.Slot(str)
    def onDone(self, out_path):
        self.progress.setVisible(False)
        self.runBtn.setEnabled(True)
        self.templateBtn.setEnabled(True)
        self.reportBtn.setEnabled(True)
        self.appendLog(f"[DONE] Saved: {out_path}")
        if show_info(self, "Success", f"Saved:\n{out_path}\n\nOpen folder?") == QtWidgets.QMessageBox.Ok:
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(os.path.dirname(out_path)))

    @QtCore.Slot(str)
    def onFail(self, err):
        self.progress.setVisible(False)
        self.runBtn.setEnabled(True)
        self.templateBtn.setEnabled(True)
        self.reportBtn.setEnabled(True)
        self.appendLog("[ERROR] " + err)
        show_error(self, "Error", "Operation failed.\n\n" + err)

# ---------- entry ----------
def main():
    # Create the app FIRST
    app = QtWidgets.QApplication(sys.argv)

    # Set Fusion style AFTER app exists (avoids crashes)
    QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create("Fusion"))

    app.setWindowIcon(QtGui.QIcon(MainWindow.resource_path("wca.ico")))
    app.setOrganizationName("DocxTools")
    app.setApplicationName("DOCX Layout Copier | TF-Dena AI Section | version 1.0")

    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
