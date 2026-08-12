import os
from datetime import datetime
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFileDialog, QFrame, QLabel
)

from qfluentwidgets import (
    ScrollArea, CardWidget, LineEdit, SpinBox, SwitchButton,
    PrimaryPushButton, PushButton, InfoBar, InfoBarPosition, TitleLabel,
    SubtitleLabel, CaptionLabel, FluentIcon, IconWidget, BodyLabel,
    StrongBodyLabel
)

from metadata_engine import WordMetadataEngine
from ui.i18n import i18n


class DragDropCard(CardWidget):
    """Card widget that supports Drag & Drop for .docx files."""

    def __init__(self, parent=None, on_file_dropped=None):
        super().__init__(parent)
        self.on_file_dropped = on_file_dropped
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.docx') for url in urls):
                event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.docx'):
                if self.on_file_dropped:
                    self.on_file_dropped(file_path)
                break


class SingleFileInterface(ScrollArea):
    """Single Word Document Metadata Editor Interface."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file_path = None
        self.current_meta = {}

        self.scroll_widget = QWidget()
        self.scroll_widget.setObjectName("scrollWidget")
        self.scroll_widget.setStyleSheet("background-color: transparent;")
        self.setWidget(self.scroll_widget)
        self.setWidgetResizable(True)
        self.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        self.main_layout = QVBoxLayout(self.scroll_widget)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(36, 36, 36, 36)

        self._init_ui()
        i18n.languageChanged.connect(self.retranslate_ui)

    def _init_ui(self):
        # 1. Header Title
        self.title_label = TitleLabel(i18n.t("单文件属性精细编辑", "Single File Fine Metadata Editing"), self)
        self.subtitle_label = CaptionLabel(
            i18n.t(
                "修改 Word (.docx) 文档内部元数据（作者、创建时间、总编辑时间）及操作系统文件记录",
                "Modify Word (.docx) metadata (author, created time, total editing time) & OS file attributes"
            ),
            self
        )
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addWidget(self.subtitle_label)

        # 2. File Selection Drop Zone Card
        self.drop_card = DragDropCard(self, on_file_dropped=self.load_file)
        drop_layout = QHBoxLayout(self.drop_card)
        drop_layout.setContentsMargins(24, 20, 24, 20)

        file_icon = IconWidget(FluentIcon.DOCUMENT, self.drop_card)
        file_icon.setFixedSize(40, 40)

        file_info_vbox = QVBoxLayout()
        self.file_name_label = StrongBodyLabel(
            i18n.t("点击选择 Word 文档 (.docx) 或将文件拖拽至此处", "Click to select Word document (.docx) or drag and drop file here"),
            self.drop_card
        )
        self.file_path_label = CaptionLabel(i18n.t("尚未选择文件", "No file selected"), self.drop_card)
        file_info_vbox.addWidget(self.file_name_label)
        file_info_vbox.addWidget(self.file_path_label)

        self.browse_btn = PushButton(FluentIcon.FOLDER, i18n.t("浏览文件", "Browse File"), self.drop_card)
        self.browse_btn.clicked.connect(self._select_file)

        drop_layout.addWidget(file_icon)
        drop_layout.addLayout(file_info_vbox, 1)
        drop_layout.addWidget(self.browse_btn)

        self.drop_card.mousePressEvent = lambda e: self._select_file() if e.button() == Qt.LeftButton else None
        self.main_layout.addWidget(self.drop_card)

        # 3. Main Form Cards (Hidden until file loaded)
        self.form_container = QWidget(self)
        form_layout = QVBoxLayout(self.form_container)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(16)

        # Card A: Document Author & Identity
        author_card = CardWidget(self.form_container)
        ac_layout = QVBoxLayout(author_card)
        ac_layout.setContentsMargins(24, 20, 24, 20)
        ac_layout.setSpacing(12)

        self.author_card_title = StrongBodyLabel(i18n.t("文档作者与归属信息", "Document Author & Ownership"), author_card)
        ac_layout.addWidget(self.author_card_title)

        # Row 1: Creator & Modifier
        row1 = QHBoxLayout()
        r1_v1 = QVBoxLayout()
        self.lbl_author = CaptionLabel(i18n.t("作者 (Creator)", "Author (Creator)"), author_card)
        r1_v1.addWidget(self.lbl_author)
        self.author_input = LineEdit(author_card)
        self.author_input.setPlaceholderText(i18n.t("例如: 张三", "e.g. John Doe"))
        r1_v1.addWidget(self.author_input)

        r1_v2 = QVBoxLayout()
        self.lbl_last_mod = CaptionLabel(i18n.t("最后修改者 (Last Modified By)", "Last Modified By"), author_card)
        r1_v2.addWidget(self.lbl_last_mod)
        self.last_mod_by_input = LineEdit(author_card)
        self.last_mod_by_input.setPlaceholderText(i18n.t("例如: 李四", "e.g. Jane Doe"))
        r1_v2.addWidget(self.last_mod_by_input)

        row1.addLayout(r1_v1)
        row1.addLayout(r1_v2)
        ac_layout.addLayout(row1)

        # Row 2: Company & Application
        row2 = QHBoxLayout()
        r2_v1 = QVBoxLayout()
        self.lbl_company = CaptionLabel(i18n.t("公司 / 单位名称 (Company)", "Company / Organization"), author_card)
        r2_v1.addWidget(self.lbl_company)
        self.company_input = LineEdit(author_card)
        self.company_input.setPlaceholderText(i18n.t("例如: Microsoft Corporation", "e.g. Acme Corp"))
        r2_v1.addWidget(self.company_input)

        r2_v2 = QVBoxLayout()
        self.lbl_doc_title = CaptionLabel(i18n.t("文档标题 (Title)", "Document Title"), author_card)
        r2_v2.addWidget(self.lbl_doc_title)
        self.title_input = LineEdit(author_card)
        self.title_input.setPlaceholderText(i18n.t("例如: 2026年度财务分析报告", "e.g. Annual Financial Report"))
        r2_v2.addWidget(self.title_input)

        row2.addLayout(r2_v1)
        row2.addLayout(r2_v2)
        ac_layout.addLayout(row2)

        # Row 3: Comments
        row3 = QHBoxLayout()
        r3_v1 = QVBoxLayout()
        self.lbl_comments = CaptionLabel(i18n.t("备注 / 摘要说明 (Comments)", "Comments / Description"), author_card)
        r3_v1.addWidget(self.lbl_comments)
        self.comments_input = LineEdit(author_card)
        self.comments_input.setPlaceholderText(i18n.t("例如: 本文档包含内部保密数据", "e.g. Internal confidential document"))
        r3_v1.addWidget(self.comments_input)
        row3.addLayout(r3_v1)
        ac_layout.addLayout(row3)

        form_layout.addWidget(author_card)

        # Card B: Time & Revision Properties
        time_card = CardWidget(self.form_container)
        tc_layout = QVBoxLayout(time_card)
        tc_layout.setContentsMargins(24, 20, 24, 20)
        tc_layout.setSpacing(12)

        self.time_card_title = StrongBodyLabel(i18n.t("时间戳与统计数据", "Timestamps & Statistics"), time_card)
        tc_layout.addWidget(self.time_card_title)

        # Created & Modified Time
        t_row1 = QHBoxLayout()
        tr1_v1 = QVBoxLayout()
        self.lbl_created_time = CaptionLabel(i18n.t("创建时间 (Created Time)", "Created Time"), time_card)
        tr1_v1.addWidget(self.lbl_created_time)
        t1_hb = QHBoxLayout()
        self.created_time_input = LineEdit(time_card)
        self.created_time_input.setPlaceholderText(i18n.t("格式: YYYY-MM-DD HH:MM:SS", "Format: YYYY-MM-DD HH:MM:SS"))
        self.btn_now_created = PushButton(i18n.t("设为当前", "Set Now"), time_card)
        self.btn_now_created.clicked.connect(lambda: self.created_time_input.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        t1_hb.addWidget(self.created_time_input, 1)
        t1_hb.addWidget(self.btn_now_created)
        tr1_v1.addLayout(t1_hb)

        tr1_v2 = QVBoxLayout()
        self.lbl_modified_time = CaptionLabel(i18n.t("修改时间 (Modified Time)", "Modified Time"), time_card)
        tr1_v2.addWidget(self.lbl_modified_time)
        t2_hb = QHBoxLayout()
        self.modified_time_input = LineEdit(time_card)
        self.modified_time_input.setPlaceholderText(i18n.t("格式: YYYY-MM-DD HH:MM:SS", "Format: YYYY-MM-DD HH:MM:SS"))
        self.btn_now_modified = PushButton(i18n.t("设为当前", "Set Now"), time_card)
        self.btn_now_modified.clicked.connect(lambda: self.modified_time_input.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        t2_hb.addWidget(self.modified_time_input, 1)
        t2_hb.addWidget(self.btn_now_modified)
        tr1_v2.addLayout(t2_hb)

        t_row1.addLayout(tr1_v1)
        t_row1.addLayout(tr1_v2)
        tc_layout.addLayout(t_row1)

        # Total Editing Time & Revision Number
        t_row2 = QHBoxLayout()

        tr2_v1 = QVBoxLayout()
        self.lbl_total_time = CaptionLabel(i18n.t("总编辑时间 (Total Editing Time / 分钟)", "Total Editing Time (Minutes)"), time_card)
        tr2_v1.addWidget(self.lbl_total_time)
        self.total_time_spin = SpinBox(time_card)
        self.total_time_spin.setRange(0, 999999)
        self.total_time_spin.setSingleStep(30)
        tr2_v1.addWidget(self.total_time_spin)
        self.total_time_hint = CaptionLabel(i18n.t("提示: 60 分钟 = 1 小时", "Note: 60 mins = 1 hour"), time_card)
        tr2_v1.addWidget(self.total_time_hint)

        tr2_v2 = QVBoxLayout()
        self.lbl_revision = CaptionLabel(i18n.t("修订版本号 (Revision Number)", "Revision Number"), time_card)
        tr2_v2.addWidget(self.lbl_revision)
        self.revision_spin = SpinBox(time_card)
        self.revision_spin.setRange(1, 99999)
        tr2_v2.addWidget(self.revision_spin)
        self.revision_hint = CaptionLabel(i18n.t("提示: 每次 Word 保存自动增加", "Note: Auto-increments on Word save"), time_card)
        tr2_v2.addWidget(self.revision_hint)

        t_row2.addLayout(tr2_v1)
        t_row2.addLayout(tr2_v2)
        tc_layout.addLayout(t_row2)

        form_layout.addWidget(time_card)

        # Card C: OS File System Sync Switch
        sync_card = CardWidget(self.form_container)
        sc_layout = QHBoxLayout(sync_card)
        sc_layout.setContentsMargins(24, 16, 24, 16)

        sc_info = QVBoxLayout()
        self.sync_card_title = StrongBodyLabel(i18n.t("同步更新 Windows 操作系统文件时间记录", "Sync Windows File System Timestamps"), sync_card)
        self.sync_card_subtitle = CaptionLabel(
            i18n.t(
                "开启后，修改 Word 内部属性的同时，也会改变 NTFS 文件系统的【创建时间】与【修改时间】",
                "Also updates NTFS Created & Modified timestamps on the Windows file system"
            ),
            sync_card
        )
        sc_info.addWidget(self.sync_card_title)
        sc_info.addWidget(self.sync_card_subtitle)

        self.sync_switch = SwitchButton(sync_card)
        self.sync_switch.setChecked(True)

        sc_layout.addLayout(sc_info, 1)
        sc_layout.addWidget(self.sync_switch)
        form_layout.addWidget(sync_card)

        # Card D: Action Bar Buttons
        action_layout = QHBoxLayout()
        action_layout.setContentsMargins(0, 10, 0, 10)

        self.save_btn = PrimaryPushButton(FluentIcon.SAVE, i18n.t("保存所有修改", "Save All Changes"), self.form_container)
        self.save_btn.clicked.connect(self._save_metadata)

        self.scrub_btn = PushButton(FluentIcon.DELETE, i18n.t("一键脱敏 / 抹去隐私", "Anonymize / Clear Metadata"), self.form_container)
        self.scrub_btn.clicked.connect(self._anonymize_metadata)

        self.reload_btn = PushButton(FluentIcon.SYNC, i18n.t("重置 / 恢复原始值", "Reset to Original"), self.form_container)
        self.reload_btn.clicked.connect(lambda: self.load_file(self.current_file_path))

        action_layout.addWidget(self.save_btn)
        action_layout.addWidget(self.scrub_btn)
        action_layout.addWidget(self.reload_btn)
        action_layout.addStretch(1)

        form_layout.addLayout(action_layout)

        self.main_layout.addWidget(self.form_container)
        self.form_container.hide() # Hide until a file is selected

    def retranslate_ui(self):
        self.title_label.setText(i18n.t("单文件属性精细编辑", "Single File Fine Metadata Editing"))
        self.subtitle_label.setText(
            i18n.t(
                "修改 Word (.docx) 文档内部元数据（作者、创建时间、总编辑时间）及操作系统文件记录",
                "Modify Word (.docx) metadata (author, created time, total editing time) & OS file attributes"
            )
        )
        if not self.current_file_path:
            self.file_name_label.setText(i18n.t("点击选择 Word 文档 (.docx) 或将文件拖拽至此处", "Click to select Word document (.docx) or drag and drop file here"))
            self.file_path_label.setText(i18n.t("尚未选择文件", "No file selected"))
        self.browse_btn.setText(i18n.t("浏览文件", "Browse File"))

        self.author_card_title.setText(i18n.t("文档作者与归属信息", "Document Author & Ownership"))
        self.lbl_author.setText(i18n.t("作者 (Creator)", "Author (Creator)"))
        self.author_input.setPlaceholderText(i18n.t("例如: 张三", "e.g. John Doe"))
        self.lbl_last_mod.setText(i18n.t("最后修改者 (Last Modified By)", "Last Modified By"))
        self.last_mod_by_input.setPlaceholderText(i18n.t("例如: 李四", "e.g. Jane Doe"))
        self.lbl_company.setText(i18n.t("公司 / 单位名称 (Company)", "Company / Organization"))
        self.company_input.setPlaceholderText(i18n.t("例如: Microsoft Corporation", "e.g. Acme Corp"))
        self.lbl_doc_title.setText(i18n.t("文档标题 (Title)", "Document Title"))
        self.title_input.setPlaceholderText(i18n.t("例如: 2026年度财务分析报告", "e.g. Annual Financial Report"))
        self.lbl_comments.setText(i18n.t("备注 / 摘要说明 (Comments)", "Comments / Description"))
        self.comments_input.setPlaceholderText(i18n.t("例如: 本文档包含内部保密数据", "e.g. Internal confidential document"))

        self.time_card_title.setText(i18n.t("时间戳与统计数据", "Timestamps & Statistics"))
        self.lbl_created_time.setText(i18n.t("创建时间 (Created Time)", "Created Time"))
        self.created_time_input.setPlaceholderText(i18n.t("格式: YYYY-MM-DD HH:MM:SS", "Format: YYYY-MM-DD HH:MM:SS"))
        self.btn_now_created.setText(i18n.t("设为当前", "Set Now"))

        self.lbl_modified_time.setText(i18n.t("修改时间 (Modified Time)", "Modified Time"))
        self.modified_time_input.setPlaceholderText(i18n.t("格式: YYYY-MM-DD HH:MM:SS", "Format: YYYY-MM-DD HH:MM:SS"))
        self.btn_now_modified.setText(i18n.t("设为当前", "Set Now"))

        self.lbl_total_time.setText(i18n.t("总编辑时间 (Total Editing Time / 分钟)", "Total Editing Time (Minutes)"))
        self.total_time_hint.setText(i18n.t("提示: 60 分钟 = 1 小时", "Note: 60 mins = 1 hour"))
        self.lbl_revision.setText(i18n.t("修订版本号 (Revision Number)", "Revision Number"))
        self.revision_hint.setText(i18n.t("提示: 每次 Word 保存自动增加", "Note: Auto-increments on Word save"))

        self.sync_card_title.setText(i18n.t("同步更新 Windows 操作系统文件时间记录", "Sync Windows File System Timestamps"))
        self.sync_card_subtitle.setText(
            i18n.t(
                "开启后，修改 Word 内部属性的同时，也会改变 NTFS 文件系统的【创建时间】与【修改时间】",
                "Also updates NTFS Created & Modified timestamps on the Windows file system"
            )
        )

        self.save_btn.setText(i18n.t("保存所有修改", "Save All Changes"))
        self.scrub_btn.setText(i18n.t("一键脱敏 / 抹去隐私", "Anonymize / Clear Metadata"))
        self.reload_btn.setText(i18n.t("重置 / 恢复原始值", "Reset to Original"))

    def _select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            i18n.t("选择 Word 文档", "Select Word Document"),
            "",
            i18n.t("Word 文档 (*.docx)", "Word Document (*.docx)")
        )
        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path: str):
        if not file_path or not os.path.exists(file_path):
            return

        self.current_file_path = file_path
        self.file_name_label.setText(os.path.basename(file_path))
        self.file_path_label.setText(file_path)

        # Load metadata
        meta = WordMetadataEngine.read_metadata(file_path)
        self.current_meta = meta

        self.author_input.setText(meta.get('author', ''))
        self.last_mod_by_input.setText(meta.get('last_modified_by', ''))
        self.company_input.setText(meta.get('company', ''))
        self.title_input.setText(meta.get('title', ''))
        self.comments_input.setText(meta.get('comments', ''))

        # Time formatting
        c_time = meta.get('created_time', '')
        if c_time:
            c_time = c_time.replace('T', ' ').replace('Z', '')
        self.created_time_input.setText(c_time)

        m_time = meta.get('modified_time', '')
        if m_time:
            m_time = m_time.replace('T', ' ').replace('Z', '')
        self.modified_time_input.setText(m_time)

        try:
            self.total_time_spin.setValue(int(meta.get('total_editing_time', '0')))
        except ValueError:
            self.total_time_spin.setValue(0)

        try:
            self.revision_spin.setValue(int(meta.get('revision', '1')))
        except ValueError:
            self.revision_spin.setValue(1)

        self.form_container.show()

        InfoBar.success(
            title=i18n.t("成功加载文档", "Document Loaded"),
            content=i18n.t(f"已成功读取: {os.path.basename(file_path)}", f"Successfully read: {os.path.basename(file_path)}"),
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=3000,
            parent=self
        )

    def _save_metadata(self):
        if not self.current_file_path:
            return

        # Prepare new metadata dictionary
        new_meta = {
            'author': self.author_input.text().strip(),
            'last_modified_by': self.last_mod_by_input.text().strip(),
            'company': self.company_input.text().strip(),
            'title': self.title_input.text().strip(),
            'comments': self.comments_input.text().strip(),
            'total_editing_time': str(self.total_time_spin.value()),
            'revision': str(self.revision_spin.value())
        }

        # Format ISO timestamp
        c_str = self.created_time_input.text().strip()
        if c_str:
            if 'T' not in c_str:
                c_str = c_str.replace(' ', 'T')
            if not c_str.endswith('Z'):
                c_str += 'Z'
            new_meta['created_time'] = c_str

        m_str = self.modified_time_input.text().strip()
        if m_str:
            if 'T' not in m_str:
                m_str = m_str.replace(' ', 'T')
            if not m_str.endswith('Z'):
                m_str += 'Z'
            new_meta['modified_time'] = m_str

        success = WordMetadataEngine.write_metadata(
            self.current_file_path,
            new_meta,
            sync_fs_time=self.sync_switch.isChecked()
        )

        if success:
            InfoBar.success(
                title=i18n.t("保存成功", "Saved Successfully"),
                content=i18n.t("Word 文档属性及系统时间已成功更新！", "Word metadata and OS timestamps updated successfully!"),
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
            # Reload fresh state
            self.load_file(self.current_file_path)
        else:
            InfoBar.error(
                title=i18n.t("保存失败", "Save Failed"),
                content=i18n.t("更新文档属性时发生错误，请检查文件是否被 Word 独占打开。", "Error updating metadata. Please check if the file is locked by Word."),
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=5000,
                parent=self
            )

    def _anonymize_metadata(self):
        if not self.current_file_path:
            return

        self.author_input.setText("")
        self.last_mod_by_input.setText("")
        self.company_input.setText("")
        self.comments_input.setText("")
        self.total_time_spin.setValue(0)
        self.revision_spin.setValue(1)

        self._save_metadata()

        InfoBar.info(
            title=i18n.t("隐私脱敏完成", "Anonymization Complete"),
            content=i18n.t("已清空作者、修改者、公司名称并重置总编辑时间。", "Cleared author, modifier, company, and reset editing time."),
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=3000,
            parent=self
        )
