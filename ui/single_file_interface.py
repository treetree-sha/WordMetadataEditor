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

    def _init_ui(self):
        # 1. Header Title
        title_label = TitleLabel("单文件属性精细编辑", self)
        subtitle_label = CaptionLabel("修改 Word (.docx) 文档内部元数据（作者、创建时间、总编辑时间）及操作系统文件记录", self)
        self.main_layout.addWidget(title_label)
        self.main_layout.addWidget(subtitle_label)

        # 2. File Selection Drop Zone Card
        self.drop_card = DragDropCard(self, on_file_dropped=self.load_file)
        drop_layout = QHBoxLayout(self.drop_card)
        drop_layout.setContentsMargins(24, 20, 24, 20)

        file_icon = IconWidget(FluentIcon.DOCUMENT, self.drop_card)
        file_icon.setFixedSize(40, 40)

        file_info_vbox = QVBoxLayout()
        self.file_name_label = StrongBodyLabel("点击选择 Word 文档 (.docx) 或将文件拖拽至此处", self.drop_card)
        self.file_path_label = CaptionLabel("尚未选择文件", self.drop_card)
        file_info_vbox.addWidget(self.file_name_label)
        file_info_vbox.addWidget(self.file_path_label)

        self.browse_btn = PushButton(FluentIcon.FOLDER, "浏览文件", self.drop_card)
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

        ac_layout.addWidget(StrongBodyLabel("文档作者与归属信息", author_card))

        # Row 1: Creator & Modifier
        row1 = QHBoxLayout()
        r1_v1 = QVBoxLayout()
        r1_v1.addWidget(CaptionLabel("作者 (Creator)", author_card))
        self.author_input = LineEdit(author_card)
        self.author_input.setPlaceholderText("例如: 张三")
        r1_v1.addWidget(self.author_input)

        r1_v2 = QVBoxLayout()
        r1_v2.addWidget(CaptionLabel("最后修改者 (Last Modified By)", author_card))
        self.last_mod_by_input = LineEdit(author_card)
        self.last_mod_by_input.setPlaceholderText("例如: 李四")
        r1_v2.addWidget(self.last_mod_by_input)

        row1.addLayout(r1_v1)
        row1.addLayout(r1_v2)
        ac_layout.addLayout(row1)

        # Row 2: Company & Application
        row2 = QHBoxLayout()
        r2_v1 = QVBoxLayout()
        r2_v1.addWidget(CaptionLabel("公司 / 单位名称 (Company)", author_card))
        self.company_input = LineEdit(author_card)
        self.company_input.setPlaceholderText("例如: Microsoft Corporation")
        r2_v1.addWidget(self.company_input)

        r2_v2 = QVBoxLayout()
        r2_v2.addWidget(CaptionLabel("文档标题 (Title)", author_card))
        self.title_input = LineEdit(author_card)
        self.title_input.setPlaceholderText("例如: 2026年度财务分析报告")
        r2_v2.addWidget(self.title_input)

        row2.addLayout(r2_v1)
        row2.addLayout(r2_v2)
        ac_layout.addLayout(row2)

        form_layout.addWidget(author_card)

        # Card B: Time & Revision Properties
        time_card = CardWidget(self.form_container)
        tc_layout = QVBoxLayout(time_card)
        tc_layout.setContentsMargins(24, 20, 24, 20)
        tc_layout.setSpacing(12)

        tc_layout.addWidget(StrongBodyLabel("时间戳与统计数据", time_card))

        # Created & Modified Time
        t_row1 = QHBoxLayout()
        tr1_v1 = QVBoxLayout()
        tr1_v1.addWidget(CaptionLabel("创建时间 (Created Time)", time_card))
        t1_hb = QHBoxLayout()
        self.created_time_input = LineEdit(time_card)
        self.created_time_input.setPlaceholderText("格式: YYYY-MM-DD HH:MM:SS")
        btn_now_created = PushButton("设为当前", time_card)
        btn_now_created.clicked.connect(lambda: self.created_time_input.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        t1_hb.addWidget(self.created_time_input, 1)
        t1_hb.addWidget(btn_now_created)
        tr1_v1.addLayout(t1_hb)

        tr1_v2 = QVBoxLayout()
        tr1_v2.addWidget(CaptionLabel("修改时间 (Modified Time)", time_card))
        t2_hb = QHBoxLayout()
        self.modified_time_input = LineEdit(time_card)
        self.modified_time_input.setPlaceholderText("格式: YYYY-MM-DD HH:MM:SS")
        btn_now_modified = PushButton("设为当前", time_card)
        btn_now_modified.clicked.connect(lambda: self.modified_time_input.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        t2_hb.addWidget(self.modified_time_input, 1)
        t2_hb.addWidget(btn_now_modified)
        tr1_v2.addLayout(t2_hb)

        t_row1.addLayout(tr1_v1)
        t_row1.addLayout(tr1_v2)
        tc_layout.addLayout(t_row1)

        # Total Editing Time & Revision Number
        t_row2 = QHBoxLayout()

        tr2_v1 = QVBoxLayout()
        tr2_v1.addWidget(CaptionLabel("总编辑时间 (Total Editing Time / 分钟)", time_card))
        self.total_time_spin = SpinBox(time_card)
        self.total_time_spin.setRange(0, 999999)
        self.total_time_spin.setSingleStep(30)
        tr2_v1.addWidget(self.total_time_spin)
        self.total_time_hint = CaptionLabel("提示: 60 分钟 = 1 小时", time_card)
        tr2_v1.addWidget(self.total_time_hint)

        tr2_v2 = QVBoxLayout()
        tr2_v2.addWidget(CaptionLabel("修订版本号 (Revision Number)", time_card))
        self.revision_spin = SpinBox(time_card)
        self.revision_spin.setRange(1, 99999)
        tr2_v2.addWidget(self.revision_spin)
        tr2_v2.addWidget(CaptionLabel("提示: 每次 Word 保存自动增加", time_card))

        t_row2.addLayout(tr2_v1)
        t_row2.addLayout(tr2_v2)
        tc_layout.addLayout(t_row2)

        form_layout.addWidget(time_card)

        # Card C: OS File System Sync Switch
        sync_card = CardWidget(self.form_container)
        sc_layout = QHBoxLayout(sync_card)
        sc_layout.setContentsMargins(24, 16, 24, 16)

        sc_info = QVBoxLayout()
        sc_info.addWidget(StrongBodyLabel("同步更新 Windows 操作系统文件时间记录", sync_card))
        sc_info.addWidget(CaptionLabel("开启后，修改 Word 内部属性的同时，也会改变 NTFS 文件系统的【创建时间】与【修改时间】", sync_card))

        self.sync_switch = SwitchButton(sync_card)
        self.sync_switch.setChecked(True)

        sc_layout.addLayout(sc_info, 1)
        sc_layout.addWidget(self.sync_switch)
        form_layout.addWidget(sync_card)

        # Card D: Action Bar Buttons
        action_layout = QHBoxLayout()
        action_layout.setContentsMargins(0, 10, 0, 10)

        self.save_btn = PrimaryPushButton(FluentIcon.SAVE, "保存所有修改", self.form_container)
        self.save_btn.clicked.connect(self._save_metadata)

        self.scrub_btn = PushButton(FluentIcon.DELETE, "一键脱敏 / 抹去隐私", self.form_container)
        self.scrub_btn.clicked.connect(self._anonymize_metadata)

        self.reload_btn = PushButton(FluentIcon.SYNC, "重置 / 恢复原始值", self.form_container)
        self.reload_btn.clicked.connect(lambda: self.load_file(self.current_file_path))

        action_layout.addWidget(self.save_btn)
        action_layout.addWidget(self.scrub_btn)
        action_layout.addWidget(self.reload_btn)
        action_layout.addStretch(1)

        form_layout.addLayout(action_layout)

        self.main_layout.addWidget(self.form_container)
        self.form_container.hide() # Hide until a file is selected

    def _select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 文档", "", "Word 文档 (*.docx)"
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
            title="成功加载文档",
            content=f"已成功读取: {os.path.basename(file_path)}",
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
                title="保存成功",
                content="Word 文档属性及系统时间已成功更新！",
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
                title="保存失败",
                content="更新文档属性时发生错误，请检查文件是否被 Word 独占打开。",
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
        self.total_time_spin.setValue(0)
        self.revision_spin.setValue(1)

        self._save_metadata()

        InfoBar.info(
            title="隐私脱敏完成",
            content="已清空作者、修改者、公司名称并重置总编辑时间。",
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=3000,
            parent=self
        )
