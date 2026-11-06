from metadata_engine import FILE_PATTERNS
from ui.i18n import i18n


def office_file_filter():
    return ';;'.join([
        i18n.t(f'Office 文件 ({FILE_PATTERNS})', f'Office Files ({FILE_PATTERNS})'),
        i18n.t('Word 文档 (*.docx *.docm *.doc)', 'Word Documents (*.docx *.docm *.doc)'),
        i18n.t('Excel 工作簿 (*.xlsx *.xlsm *.xlsb *.xls)', 'Excel Workbooks (*.xlsx *.xlsm *.xlsb *.xls)'),
        i18n.t('PowerPoint 演示文稿 (*.pptx *.pptm *.ppt)', 'PowerPoint Presentations (*.pptx *.pptm *.ppt)'),
    ])
