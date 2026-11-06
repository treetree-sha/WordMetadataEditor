from metadata_engine import FILE_PATTERNS
from ui.i18n import i18n


def office_file_filter():
    return ';;'.join([
        i18n.t(f'Office 文件 ({FILE_PATTERNS})', f'Office Files ({FILE_PATTERNS})'),
        i18n.t('Word 文档 (*.docx *.docm)', 'Word Documents (*.docx *.docm)'),
        i18n.t('Excel 工作簿 (*.xlsx *.xlsm)', 'Excel Workbooks (*.xlsx *.xlsm)'),
        i18n.t('PowerPoint 演示文稿 (*.pptx *.pptm)', 'PowerPoint Presentations (*.pptx *.pptm)'),
    ])
