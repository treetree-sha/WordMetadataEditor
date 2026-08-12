import os
import random
from datetime import datetime
from PySide6.QtCore import Qt, QThread, Signal, QDateTime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFileDialog, QHeaderView, QTableWidgetItem
)

from qfluentwidgets import (
    ScrollArea, CardWidget, LineEdit, SpinBox, SwitchButton,
    PrimaryPushButton, PushButton, InfoBar, InfoBarPosition, TitleLabel,
    CaptionLabel, FluentIcon, TableWidget, ProgressBar, CheckBox,
    StrongBodyLabel, DateTimeEdit
)

from metadata_engine import WordMetadataEngine
from ui.i18n import i18n


class BatchWorkerThread(QThread):
    """Background thread for processing batch files without freezing UI."""
    progress_changed = Signal(int, int, str, bool) # (current, total, filename, success)
    finished_batch = Signal(int, int) # (success_count, fail_count)

    def __init__(self, file_list, batch_config, sync_fs_time=True):
        super().__init__()
        self.file_list = file_list
        self.batch_config = batch_config
        self.sync_fs_time = sync_fs_time

    def run(self):
        total = len(self.file_list)
        success_count = 0
        fail_count = 0

        for i, file_path in enumerate(self.file_list):
            if not os.path.exists(file_path):
                self.progress_changed.emit(i + 1, total, os.path.basename(file_path), False)
                fail_count += 1
                continue

            try:
                meta = WordMetadataEngine.read_metadata(file_path)

                if self.batch_config.get('anonymize'):
                    meta['author'] = ''
                    meta['last_modified_by'] = ''
                    meta['company'] = ''
                    meta['comments'] = ''
                    meta['total_editing_time'] = '0'
                    meta['created'] = ''
                else:
                    if self.batch_config.get('set_author'):
                        meta['author'] = self.batch_config['author']
                    if self.batch_config.get('set_modifier'):
                        meta['last_modified_by'] = self.batch_config['modifier']
                    if self.batch_config.get('set_comments'):
                        meta['comments'] = self.batch_config['comments']

                    # Handle Total Editing Time
                    if self.batch_config.get('set_random_time'):
                        min_t = self.batch_config['min_time']
                        max_t = self.batch_config['max_time']
                        low, high = min(min_t, max_t), max(min_t, max_t)
                        random_val = random.randint(low, high)
                        meta['total_editing_time'] = str(random_val)
                    elif self.batch_config.get('set_total_time'):
                        meta['total_editing_time'] = str(self.batch_config['total_time'])

                    # Handle Creation Time
                    if self.batch_config.get('set_random_created_time'):
                        start_dt = self.batch_config['start_created_time']
                        end_dt = self.batch_config['end_created_time']
                        start_ts = int(start_dt.timestamp())
                        end_ts = int(end_dt.timestamp())
                        low_ts, high_ts = min(start_ts, end_ts), max(start_ts, end_ts)
                        rand_ts = random.randint(low_ts, high_ts)
                        rand_dt = datetime.fromtimestamp(rand_ts)
                        meta['created'] = rand_dt.strftime('%Y-%m-%dT%H:%M:%SZ')
                    elif self.batch_config.get('set_created_time'):
                        dt_val = self.batch_config['created_time']
                        meta['created'] = dt_val.strftime('%Y-%m-%dT%H:%M:%SZ')

                ok = WordMetadataEngine.write_metadata(file_path, meta, sync_fs_time=self.sync_fs_time)
                if ok:
                    success_count += 1
                else:
                    fail_count += 1

                self.progress_changed.emit(i + 1, total, os.path.basename(file_path), ok)

            except Exception as e:
                print(f"Error processing {file_path}: {e}")
                fail_count += 1
                self.progress_changed.emit(i + 1, total, os.path.basename(file_path), False)

        self.finished_batch.emit(success_count, fail_count)


class BatchInterface(ScrollArea):
    """Batch Processing Interface for Word Documents."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_paths = []

        self.scroll_widget = QWidget()
        self.scroll_widget.setObjectName("batchScrollWidget")
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
        self.title_label = TitleLabel(i18n.t("批量文件属性修改", "Batch Document Metadata Editor"), self)
        self.subtitle_label = CaptionLabel(
            i18n.t(
                "批量导入 Word 文档，一键统一修改作者、编辑时间、创建时间或进行脱敏处理",
                "Batch import Word documents to unify author, editing time, creation time, or anonymize in bulk"
            ),
            self
        )
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addWidget(self.subtitle_label)

        # 2. File Import Action Bar Card
        bar_card = CardWidget(self)
        bar_layout = QHBoxLayout(bar_card)
        bar_layout.setContentsMargins(20, 16, 20, 16)

        self.btn_add_files = PushButton(FluentIcon.DOCUMENT, i18n.t("添加文件...", "Add Files..."), bar_card)
        self.btn_add_files.clicked.connect(self._add_files)

        self.btn_add_folder = PushButton(FluentIcon.FOLDER, i18n.t("添加文件夹...", "Add Folder..."), bar_card)
        self.btn_add_folder.clicked.connect(self._add_folder)

        self.btn_clear_list = PushButton(FluentIcon.DELETE, i18n.t("清空列表", "Clear List"), bar_card)
        self.btn_clear_list.clicked.connect(self._clear_list)

        bar_layout.addWidget(self.btn_add_files)
        bar_layout.addWidget(self.btn_add_folder)
        bar_layout.addWidget(self.btn_clear_list)
        bar_layout.addStretch(1)

        self.file_count_label = CaptionLabel(i18n.t("共 0 个文档", "0 documents total"), bar_card)
        bar_layout.addWidget(self.file_count_label)

        self.main_layout.addWidget(bar_card)

        # 3. File Table
        self.table = TableWidget(self)
        self.table.setColumnCount(6)
        self._set_table_headers()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setMinimumHeight(240)
        self.main_layout.addWidget(self.table)

        # 4. Batch Operations Config Card
        config_card = CardWidget(self)
        cfg_layout = QVBoxLayout(config_card)
        cfg_layout.setContentsMargins(24, 20, 24, 20)
        cfg_layout.setSpacing(12)

        self.config_card_title = StrongBodyLabel(i18n.t("批量修改规则配置", "Batch Modification Rules"), config_card)
        cfg_layout.addWidget(self.config_card_title)

        # Checkbox 1: Set Author
        row1 = QHBoxLayout()
        self.chk_author = CheckBox(i18n.t("统一设置作者", "Set Unified Author"), config_card)
        self.input_author = LineEdit(config_card)
        self.input_author.setPlaceholderText(i18n.t("填入新作者姓名", "Enter new author name"))
        row1.addWidget(self.chk_author)
        row1.addWidget(self.input_author, 1)
        cfg_layout.addLayout(row1)

        # Checkbox 2: Set Modifier
        row2 = QHBoxLayout()
        self.chk_modifier = CheckBox(i18n.t("统一设置修改人", "Set Unified Modifier"), config_card)
        self.input_modifier = LineEdit(config_card)
        self.input_modifier.setPlaceholderText(i18n.t("填入新修改人姓名", "Enter new modifier name"))
        row2.addWidget(self.chk_modifier)
        row2.addWidget(self.input_modifier, 1)
        cfg_layout.addLayout(row2)

        # Checkbox 2.5: Set Comments
        row2_5 = QHBoxLayout()
        self.chk_comments = CheckBox(i18n.t("统一设置备注", "Set Unified Comments"), config_card)
        self.input_comments = LineEdit(config_card)
        self.input_comments.setPlaceholderText(i18n.t("填入新备注内容", "Enter new comments"))
        row2_5.addWidget(self.chk_comments)
        row2_5.addWidget(self.input_comments, 1)
        cfg_layout.addLayout(row2_5)

        # Checkbox 3: Set Fixed Total Time
        row3 = QHBoxLayout()
        self.chk_total_time = CheckBox(i18n.t("统一固定总编辑时间", "Set Fixed Total Editing Time"), config_card)
        self.spin_total_time = SpinBox(config_card)
        self.spin_total_time.setRange(0, 999999)
        self.spin_total_time.setValue(120)
        self.lbl_minutes = CaptionLabel(i18n.t("分钟", "Minutes"), config_card)
        row3.addWidget(self.chk_total_time)
        row3.addWidget(self.spin_total_time)
        row3.addWidget(self.lbl_minutes)
        row3.addStretch(1)
        cfg_layout.addLayout(row3)

        # Checkbox 3.5: Set Random Total Time Range
        row3_5 = QHBoxLayout()
        self.chk_random_time = CheckBox(i18n.t("随机生成总编辑时间 (指定范围)", "Random Total Time within Range"), config_card)
        self.spin_min_time = SpinBox(config_card)
        self.spin_min_time.setRange(0, 999999)
        self.spin_min_time.setValue(30)
        self.lbl_range_to = CaptionLabel(i18n.t("至", "to"), config_card)
        self.spin_max_time = SpinBox(config_card)
        self.spin_max_time.setRange(0, 999999)
        self.spin_max_time.setValue(180)
        self.lbl_random_minutes = CaptionLabel(i18n.t("分钟", "Minutes"), config_card)

        row3_5.addWidget(self.chk_random_time)
        row3_5.addWidget(self.spin_min_time)
        row3_5.addWidget(self.lbl_range_to)
        row3_5.addWidget(self.spin_max_time)
        row3_5.addWidget(self.lbl_random_minutes)
        row3_5.addStretch(1)
        cfg_layout.addLayout(row3_5)

        # Mutually exclusive checkboxes for total time vs random time
        self.chk_total_time.stateChanged.connect(self._on_fixed_time_checked)
        self.chk_random_time.stateChanged.connect(self._on_random_time_checked)

        # Checkbox 4: Set Fixed Creation Time
        row6 = QHBoxLayout()
        self.chk_created_time = CheckBox(i18n.t("统一固定文件创建时间", "Set Fixed File Creation Time"), config_card)
        self.dt_created_time = DateTimeEdit(config_card)
        self.dt_created_time.setDateTime(QDateTime.currentDateTime())
        row6.addWidget(self.chk_created_time)
        row6.addWidget(self.dt_created_time)
        row6.addStretch(1)
        cfg_layout.addLayout(row6)

        # Checkbox 5: Set Random Creation Time Range
        row7 = QHBoxLayout()
        self.chk_random_created_time = CheckBox(i18n.t("随机生成文件创建时间 (指定范围)", "Random Creation Time within Range"), config_card)
        self.dt_start_created_time = DateTimeEdit(config_card)
        self.dt_start_created_time.setDateTime(QDateTime.currentDateTime().addDays(-30))
        self.lbl_created_range_to = CaptionLabel(i18n.t("至", "to"), config_card)
        self.dt_end_created_time = DateTimeEdit(config_card)
        self.dt_end_created_time.setDateTime(QDateTime.currentDateTime())

        row7.addWidget(self.chk_random_created_time)
        row7.addWidget(self.dt_start_created_time)
        row7.addWidget(self.lbl_created_range_to)
        row7.addWidget(self.dt_end_created_time)
        row7.addStretch(1)
        cfg_layout.addLayout(row7)

        # Mutually exclusive checkboxes for fixed created time vs random created time
        self.chk_created_time.stateChanged.connect(self._on_fixed_created_time_checked)
        self.chk_random_created_time.stateChanged.connect(self._on_random_created_time_checked)

        # Checkbox 6: Anonymize
        row4 = QHBoxLayout()
        self.chk_anonymize = CheckBox(i18n.t("一键隐私脱敏 (清空所有作者与公司信息)", "Anonymize All (Clear author & company info)"), config_card)
        row4.addWidget(self.chk_anonymize)
        cfg_layout.addLayout(row4)

        # Switch: OS Sync
        row5 = QHBoxLayout()
        self.lbl_sync_fs = CaptionLabel(i18n.t("同步修改系统文件时间记录", "Sync OS File System Timestamps"), config_card)
        row5.addWidget(self.lbl_sync_fs)
        self.sync_fs_switch = SwitchButton(config_card)
        self.sync_fs_switch.setChecked(True)
        row5.addWidget(self.sync_fs_switch)
        row5.addStretch(1)
        cfg_layout.addLayout(row5)

        self.main_layout.addWidget(config_card)

        # 5. Execution Card
        exec_card = CardWidget(self)
        ex_layout = QVBoxLayout(exec_card)
        ex_layout.setContentsMargins(24, 16, 24, 16)

        self.progress_bar = ProgressBar(exec_card)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()

        ex_row = QHBoxLayout()
        self.btn_run_batch = PrimaryPushButton(FluentIcon.PLAY, i18n.t("开始批量执行修改", "Start Batch Processing"), exec_card)
        self.btn_run_batch.clicked.connect(self._run_batch_process)
        ex_row.addWidget(self.btn_run_batch)
        ex_row.addStretch(1)

        ex_layout.addWidget(self.progress_bar)
        ex_layout.addLayout(ex_row)

        self.main_layout.addWidget(exec_card)

    def _on_fixed_time_checked(self, state):
        if state == Qt.Checked and self.chk_random_time.isChecked():
            self.chk_random_time.setChecked(False)

    def _on_random_time_checked(self, state):
        if state == Qt.Checked and self.chk_total_time.isChecked():
            self.chk_total_time.setChecked(False)

    def _on_fixed_created_time_checked(self, state):
        if state == Qt.Checked and self.chk_random_created_time.isChecked():
            self.chk_random_created_time.setChecked(False)

    def _on_random_created_time_checked(self, state):
        if state == Qt.Checked and self.chk_created_time.isChecked():
            self.chk_created_time.setChecked(False)

    def _set_table_headers(self):
        self.table.setHorizontalHeaderLabels([
            i18n.t("文件名", "File Name"),
            i18n.t("作者", "Author"),
            i18n.t("修改人", "Modifier"),
            i18n.t("总编辑时间(分)", "Total Time (min)"),
            i18n.t("路径", "Path"),
            i18n.t("状态", "Status")
        ])

    def retranslate_ui(self):
        self.title_label.setText(i18n.t("批量文件属性修改", "Batch Document Metadata Editor"))
        self.subtitle_label.setText(
            i18n.t(
                "批量导入 Word 文档，一键统一修改作者、编辑时间、创建时间或进行脱敏处理",
                "Batch import Word documents to unify author, editing time, creation time, or anonymize in bulk"
            )
        )
        self.btn_add_files.setText(i18n.t("添加文件...", "Add Files..."))
        self.btn_add_folder.setText(i18n.t("添加文件夹...", "Add Folder..."))
        self.btn_clear_list.setText(i18n.t("清空列表", "Clear List"))
        self.file_count_label.setText(i18n.t(f"共 {len(self.file_paths)} 个文档", f"{len(self.file_paths)} documents total"))

        self._set_table_headers()
        self.config_card_title.setText(i18n.t("批量修改规则配置", "Batch Modification Rules"))
        self.chk_author.setText(i18n.t("统一设置作者", "Set Unified Author"))
        self.input_author.setPlaceholderText(i18n.t("填入新作者姓名", "Enter new author name"))
        self.chk_modifier.setText(i18n.t("统一设置修改人", "Set Unified Modifier"))
        self.input_modifier.setPlaceholderText(i18n.t("填入新修改人姓名", "Enter new modifier name"))
        self.chk_comments.setText(i18n.t("统一设置备注", "Set Unified Comments"))
        self.input_comments.setPlaceholderText(i18n.t("填入新备注内容", "Enter new comments"))
        self.chk_total_time.setText(i18n.t("统一固定总编辑时间", "Set Fixed Total Editing Time"))
        self.lbl_minutes.setText(i18n.t("分钟", "Minutes"))
        self.chk_random_time.setText(i18n.t("随机生成总编辑时间 (指定范围)", "Random Total Time within Range"))
        self.lbl_range_to.setText(i18n.t("至", "to"))
        self.lbl_random_minutes.setText(i18n.t("分钟", "Minutes"))
        self.chk_created_time.setText(i18n.t("统一固定文件创建时间", "Set Fixed File Creation Time"))
        self.chk_random_created_time.setText(i18n.t("随机生成文件创建时间 (指定范围)", "Random Creation Time within Range"))
        self.lbl_created_range_to.setText(i18n.t("至", "to"))
        self.chk_anonymize.setText(i18n.t("一键隐私脱敏 (清空所有作者与公司信息)", "Anonymize All (Clear author & company info)"))
        self.lbl_sync_fs.setText(i18n.t("同步修改系统文件时间记录", "Sync OS File System Timestamps"))
        self.btn_run_batch.setText(i18n.t("开始批量执行修改", "Start Batch Processing"))

        # Refresh table rows status text
        for i in range(self.table.rowCount()):
            status_item = self.table.item(i, 5)
            if status_item:
                curr_txt = status_item.text()
                if curr_txt in ("待处理", "Pending"):
                    status_item.setText(i18n.t("待处理", "Pending"))
                elif curr_txt in ("成功", "Success"):
                    status_item.setText(i18n.t("成功", "Success"))
                elif curr_txt in ("失败", "Failed"):
                    status_item.setText(i18n.t("失败", "Failed"))

    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            i18n.t("选择 Word 文档", "Select Word Documents"),
            "",
            i18n.t("Word 文档 (*.docx)", "Word Documents (*.docx)")
        )
        if files:
            for f in files:
                if f not in self.file_paths:
                    self.file_paths.append(f)
            self._update_table()

    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, i18n.t("选择文件夹", "Select Folder"))
        if folder:
            added_any = False
            for root, _, files in os.walk(folder):
                for filename in files:
                    if filename.lower().endswith('.docx') and not filename.startswith('~$'):
                        full_path = os.path.join(root, filename)
                        if full_path not in self.file_paths:
                            self.file_paths.append(full_path)
                            added_any = True
            if added_any:
                self._update_table()

    def _clear_list(self):
        self.file_paths.clear()
        self._update_table()

    def _update_table(self):
        self.table.setRowCount(len(self.file_paths))
        self.file_count_label.setText(i18n.t(f"共 {len(self.file_paths)} 个文档", f"{len(self.file_paths)} documents total"))

        for i, file_path in enumerate(self.file_paths):
            filename = os.path.basename(file_path)
            meta = WordMetadataEngine.read_metadata(file_path)

            self.table.setItem(i, 0, QTableWidgetItem(filename))
            self.table.setItem(i, 1, QTableWidgetItem(meta.get('author', '')))
            self.table.setItem(i, 2, QTableWidgetItem(meta.get('last_modified_by', '')))
            self.table.setItem(i, 3, QTableWidgetItem(meta.get('total_editing_time', '0')))
            self.table.setItem(i, 4, QTableWidgetItem(file_path))
            self.table.setItem(i, 5, QTableWidgetItem(i18n.t("待处理", "Pending")))

    def _run_batch_process(self):
        if not self.file_paths:
            InfoBar.warning(
                title=i18n.t("提示", "Notice"),
                content=i18n.t("请先添加需要修改的 Word 文档", "Please add Word documents to modify first"),
                orient=Qt.Horizontal,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
            return

        batch_config = {
            'set_author': self.chk_author.isChecked(),
            'author': self.input_author.text().strip(),
            'set_modifier': self.chk_modifier.isChecked(),
            'modifier': self.input_modifier.text().strip(),
            'set_comments': self.chk_comments.isChecked(),
            'comments': self.input_comments.text().strip(),
            'set_total_time': self.chk_total_time.isChecked(),
            'total_time': self.spin_total_time.value(),
            'set_random_time': self.chk_random_time.isChecked(),
            'min_time': self.spin_min_time.value(),
            'max_time': self.spin_max_time.value(),
            'set_created_time': self.chk_created_time.isChecked(),
            'created_time': self.dt_created_time.dateTime().toPython(),
            'set_random_created_time': self.chk_random_created_time.isChecked(),
            'start_created_time': self.dt_start_created_time.dateTime().toPython(),
            'end_created_time': self.dt_end_created_time.dateTime().toPython(),
            'anonymize': self.chk_anonymize.isChecked()
        }

        if not any([
            batch_config['set_author'],
            batch_config['set_modifier'],
            batch_config['set_comments'],
            batch_config['set_total_time'],
            batch_config['set_random_time'],
            batch_config['set_created_time'],
            batch_config['set_random_created_time'],
            batch_config['anonymize']
        ]):
            InfoBar.warning(
                title=i18n.t("提示", "Notice"),
                content=i18n.t("请至少勾选一种批量修改规则（如：统一修改作者或脱敏）", "Please select at least one modification rule"),
                orient=Qt.Horizontal,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
            return

        self.btn_run_batch.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()

        self.thread = BatchWorkerThread(self.file_paths, batch_config, sync_fs_time=self.sync_fs_switch.isChecked())
        self.thread.progress_changed.connect(self._on_progress)
        self.thread.finished_batch.connect(self._on_finished)
        self.thread.start()

    def _on_progress(self, current, total, filename, success):
        pct = int((current / total) * 100)
        self.progress_bar.setValue(pct)
        item_index = current - 1
        status_text = i18n.t("成功", "Success") if success else i18n.t("失败", "Failed")
        self.table.setItem(item_index, 5, QTableWidgetItem(status_text))

    def _on_finished(self, success_count, fail_count):
        self.btn_run_batch.setEnabled(True)
        self.progress_bar.hide()
        self._update_table()

        InfoBar.success(
            title=i18n.t("批量处理完成", "Batch Processing Complete"),
            content=i18n.t(f"已成功处理 {success_count} 个文件，失败 {fail_count} 个。", f"Successfully processed {success_count} files, failed {fail_count}."),
            orient=Qt.Horizontal,
            position=InfoBarPosition.TOP,
            duration=5000,
            parent=self
        )
