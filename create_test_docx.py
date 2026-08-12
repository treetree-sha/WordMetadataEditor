import docx
from metadata_engine import WordMetadataEngine

doc = docx.Document()
doc.add_heading('测试 Word 文档', 0)
doc.add_paragraph('这是一个用于测试属性编辑器的示例 Word 文档。')
doc.save('test_document.docx')

print("Created test_document.docx")

meta = WordMetadataEngine.read_metadata('test_document.docx')
print("Initial Metadata:", meta)

# Modify metadata
meta['author'] = '张三'
meta['last_modified_by'] = '李四'
meta['total_editing_time'] = '180' # 3 hours
meta['revision'] = '5'
meta['created_time'] = '2025-01-01T09:00:00Z'
meta['modified_time'] = '2025-01-02T14:30:00Z'
meta['title'] = '年度总结报告'
meta['company'] = '示例科技公司'

WordMetadataEngine.write_metadata('test_document.docx', meta, sync_fs_time=True)

updated_meta = WordMetadataEngine.read_metadata('test_document.docx')
print("Updated Metadata:", updated_meta)
