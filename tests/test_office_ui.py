"""Exercise the actual Qt forms and mixed-format background worker offscreen."""
import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pytest
from PySide6.QtCore import QMimeData, QPoint, QPointF, Qt, QUrl, QEventLoop, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QApplication, QFileDialog
from qfluentwidgets import InfoBar

from main import MainWindow
from metadata_engine import OfficeMetadataEngine as Engine, format_ooxml_datetime
from ui.batch_interface import BatchWorkerThread
from ui.i18n import i18n
from ui.file_filters import office_file_filter
from test_office_metadata import make_office, contents  # shared real-document fixture


@pytest.fixture(scope='module')
def app():
    instance = QApplication.instance() or QApplication([])
    yield instance


@pytest.fixture
def window(app, monkeypatch):
    for name in ('success', 'warning', 'error', 'info'):
        monkeypatch.setattr(InfoBar, name, lambda **kwargs: None)
    i18n.set_language('zh')
    win = MainWindow()
    win.show()
    app.processEvents()
    yield win
    i18n.set_language('zh')
    win.close()
    win.deleteLater()
    app.processEvents()


@pytest.mark.parametrize('ext', ['.xlsx', '.pptx', '.xlsm', '.pptm', '.xlsb'])
def test_load_save_and_anonymize_in_single_form(window, make_office, ext):
    path = make_office(ext)
    before = contents(path)
    form = window.single_interface
    form.load_file(str(path))
    assert Path(form.current_file_path) == path
    form.author_input.setText('界面作者')
    form.comments_input.setText('界面备注')
    form.total_time_spin.setValue(88)
    form.created_time_input.setText('2025-04-05 11:12:13')
    form.sync_switch.setChecked(False)
    assert form._save_metadata()
    meta = Engine.read_metadata(path)
    assert meta['author'] == '界面作者'
    assert meta['comments'] == '界面备注'
    assert meta['created_time'] == format_ooxml_datetime('2025-04-05 11:12:13')
    form._anonymize_metadata()
    meta = Engine.read_metadata(path)
    assert all(meta[key] == '' for key in ('author', 'comments', 'company', 'last_modified_by'))
    assert meta['total_editing_time'] == '0'
    after = contents(path)
    for name in before.keys() - {'docProps/core.xml', 'docProps/app.xml'}:
        assert before[name] == after[name]


def test_drag_drop_filters_and_bilingual_ui(window, make_office, app):
    form = window.single_interface
    path = make_office('.pptx')
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(path))])
    enter = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
    form.drop_card.dragEnterEvent(enter)
    assert enter.isAccepted()
    drop = QDropEvent(QPointF(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
    form.drop_card.dropEvent(drop)
    assert Path(form.current_file_path) == path
    for lang in ('en', 'zh'):
        i18n.set_language(lang)
        app.processEvents()
        assert 'Office' in window.windowTitle()
        assert all(pattern in office_file_filter() for pattern in ('*.xlsx', '*.xlsm', '*.xlsb', '*.xls', '*.doc', '*.pptx', '*.pptm', '*.ppt'))
        assert 'PowerPoint' in form.subtitle_label.text()


def test_invalid_input_does_not_overwrite_file_or_current_form(window, make_office):
    path = make_office('.xlsx')
    form = window.single_interface
    form.load_file(str(path))
    original = path.read_bytes()
    form.created_time_input.setText('not a date')
    assert form._save_metadata() is False
    assert path.read_bytes() == original
    bad = path.with_name('broken.pptx')
    bad.write_text('not an Office file')
    form.load_file(str(bad))
    assert form.current_file_path == str(path)


def test_mixed_worker_applies_fixed_date_and_retains_failures(window, make_office):
    files = [str(make_office(ext)) for ext in ('.docx', '.xlsx', '.pptx', '.xlsm', '.pptm')]
    missing = str(Path(files[0]).with_name('missing.xlsx'))
    files.append(missing)
    config = dict(set_author=True, author='统一作者', set_comments=True, comments='统一备注',
                  set_total_time=True, total_time=99, set_created_time=True,
                  created_time=datetime(2025, 1, 2, 11, 4, 5))
    worker = BatchWorkerThread(files, config, sync_fs_time=False)
    events, finished = [], []
    worker.progress_changed.connect(lambda *args: events.append(args))
    worker.finished_batch.connect(lambda *args: finished.append(args))
    worker.run()
    assert finished == [(5, 1)]
    assert len(events) == 6 and events[-1][-1] is False
    for path in files[:-1]:
        meta = Engine.read_metadata(path)
        assert meta['author'] == '统一作者' and meta['comments'] == '统一备注'
        assert meta['total_editing_time'] == '99'
        assert meta['created_time'] == format_ooxml_datetime(config['created_time'])
    batch = window.batch_interface
    batch.file_paths = files
    for args in events:
        batch._on_progress(*args)
    batch._on_finished(5, 1)
    assert all(batch.table.item(index, 5).text() == '成功' for index in range(5))
    assert batch.table.item(5, 5).text() == '失败'
    i18n.set_language('en')
    assert batch.table.item(0, 5).text() == 'Success'
    assert batch.table.item(5, 5).text() == 'Failed'


def test_random_ranges_anonymization_and_worker_snapshot(window, make_office):
    paths = [str(make_office(ext)) for ext in ('.xlsx', '.pptx')]
    config = dict(set_random_time=True, min_time=21, max_time=20,
                  set_random_created_time=True, start_created_time=datetime(2025, 2, 2),
                  end_created_time=datetime(2025, 1, 1))
    worker = BatchWorkerThread(paths, config, False)
    paths.clear()
    assert len(worker.file_list) == 2
    worker.run()
    for path in worker.file_list:
        meta = Engine.read_metadata(path)
        assert 20 <= int(meta['total_editing_time']) <= 21
        timestamp = datetime.fromisoformat(meta['created_time'].replace('Z', '+00:00')).timestamp()
        assert config['end_created_time'].timestamp() <= timestamp <= config['start_created_time'].timestamp()
    BatchWorkerThread(worker.file_list, {'anonymize': True}, False).run()
    for path in worker.file_list:
        meta = Engine.read_metadata(path)
        assert meta['author'] == '' and meta['total_editing_time'] == '0'


def test_recursive_folder_import_and_busy_controls(window, make_office):
    files = [make_office(ext) for ext in ('.docx', '.xlsx', '.pptx', '.xlsm', '.pptm')]
    nested = files[0].parent / 'nested'
    nested.mkdir()
    files[-1].rename(nested / files[-1].name)
    (nested / '~$locked.xlsx').write_bytes(b'lock')
    (nested / 'unsupported.txt').write_bytes(b'unsupported')
    batch = window.batch_interface
    with patch.object(QFileDialog, 'getExistingDirectory', return_value=str(files[0].parent)):
        batch._add_folder()
        batch._add_folder()
    assert len(batch.file_paths) == 5
    batch._set_busy(True)
    assert not batch.btn_add_files.isEnabled() and not batch.btn_clear_list.isEnabled()
    batch._clear_list()
    assert len(batch.file_paths) == 5
    batch._set_busy(False)
    batch.chk_total_time.setChecked(True)
    batch.chk_random_time.setChecked(True)
    assert not batch.chk_total_time.isChecked()
    batch.chk_created_time.setChecked(True)
    batch.chk_random_created_time.setChecked(True)
    assert not batch.chk_created_time.isChecked()


def test_real_background_batch_restores_controls_and_results(window, make_office, app):
    batch = window.batch_interface
    batch._add_paths([str(make_office(ext)) for ext in ('.xlsx', '.pptx')])
    batch.chk_author.setChecked(True)
    batch.input_author.setText('后台线程作者')
    batch.sync_fs_switch.setChecked(False)
    batch._run_batch_process()
    assert batch._batch_running
    assert not batch.btn_add_files.isEnabled()
    loop = QEventLoop()
    batch.thread.finished.connect(loop.quit)
    QTimer.singleShot(5000, loop.quit)
    loop.exec()
    assert batch.thread.wait(2000)
    app.processEvents()
    assert not batch._batch_running
    assert batch.btn_add_files.isEnabled()
    assert all(batch.table.item(index, 5).text() == '成功' for index in range(2))
    assert all(Engine.read_metadata(path)['author'] == '后台线程作者' for path in batch.file_paths)


@pytest.mark.skipif(os.name != 'nt', reason='Windows OLE backend')
@pytest.mark.parametrize('ext', ['.doc', '.xls', '.ppt'])
def test_legacy_form_drag_drop_save_and_anonymize(window, request, ext):
    from test_legacy_metadata import ole_contents
    path = request.getfixturevalue('make_legacy')(ext)
    form = window.single_interface
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(path))])
    enter = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
    form.drop_card.dragEnterEvent(enter)
    assert enter.isAccepted()
    form.drop_card.dropEvent(QDropEvent(QPointF(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier))
    assert Path(form.current_file_path) == path
    before = ole_contents(path)
    form.author_input.setText('旧版界面作者')
    form.comments_input.setText('旧版备注')
    form.sync_switch.setChecked(False)
    assert form._save_metadata()
    assert Engine.read_metadata(path)['author'] == '旧版界面作者'
    if ext == '.doc':
        assert '摘要' in form.total_time_hint.text()
        i18n.set_language('en')
        assert 'summary' in form.total_time_hint.text()
    form._anonymize_metadata()
    meta = Engine.read_metadata(path)
    assert all(meta[key] == '' for key in ('author', 'last_modified_by', 'company', 'comments'))
    assert meta['total_editing_time'] == '0'
    after = ole_contents(path)
    for name in before.keys() - {'\x05SummaryInformation', '\x05DocumentSummaryInformation'}:
        assert before[name] == after[name]


@pytest.fixture
def make_legacy(request, tmp_path):
    # Reuse the independent structured-storage fixture without a runtime dependency.
    from test_legacy_metadata import make_legacy as fixture
    return fixture.__wrapped__(tmp_path)


@pytest.mark.skipif(os.name != 'nt', reason='Windows OLE backend')
def test_encrypted_legacy_load_has_explicit_message(window, make_legacy, monkeypatch):
    valid = make_legacy('.doc')
    window.single_interface.load_file(str(valid))
    encrypted = make_legacy('.xls', encrypted=True)
    messages = []
    monkeypatch.setattr(InfoBar, 'error', lambda **kwargs: messages.append(kwargs['content']))
    window.single_interface.load_file(str(encrypted))
    assert window.single_interface.current_file_path == str(valid)
    assert messages and '加密' in messages[-1]


@pytest.mark.skipif(os.name != 'nt', reason='Windows OLE backend')
def test_mixed_legacy_batch_on_real_qthread(window, make_legacy, make_office, app):
    batch = window.batch_interface
    paths = [str(make_legacy(ext)) for ext in ('.doc', '.xls', '.ppt')]
    paths += [str(make_office(ext)) for ext in ('.xlsx', '.xlsb', '.pptx')]
    batch._add_paths(paths)
    batch.chk_author.setChecked(True)
    batch.input_author.setText('混合后台作者')
    batch.chk_comments.setChecked(True)
    batch.input_comments.setText('混合后台备注')
    batch.sync_fs_switch.setChecked(False)
    batch._run_batch_process()
    loop = QEventLoop()
    batch.thread.finished.connect(loop.quit)
    QTimer.singleShot(5000, loop.quit)
    loop.exec()
    assert batch.thread.wait(2000)
    app.processEvents()
    assert not batch._batch_running
    assert all(batch.table.item(i, 5).text() == '成功' for i in range(len(paths)))
    assert all(Engine.read_metadata(path)['author'] == '混合后台作者' for path in paths)
