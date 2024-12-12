# File Manipulator

Example **PresentationEditor** usage:
```python
with PresentationEditor("absolute_file_path") as editor:  # absolute_file_path is optional, you can include it or the origin presentation bytes.
  editor.get_slide(slide_index).comment(...)
  new_presentation_bytes = editor.export(absolute_export_path)  # absolute_export_path is optional
```

Example **SelectiveArchiveEditor** usage:
```python
with SelectiveArchiveEditor("absolute_file_path") as editor:  # absolute_file_path is optional, you can include it or the origin file bytes.
  new_file_bytes = editor.upsert(...).upsert_many(...).export(absolute_file_path)  absolute_export_path is optional
```

Example **DocumentEditor** Usage:
```python
with DocumentEditor(FILE_PATH/BYTES) as editor:
    editor.add_comment("TEXT TO HIGHLIGHT", DocCommentData(FILL WITH COMMENT INFO))  # It will automatically look for the text in all pages, as the structure of docx does not specify pages.
    new_document_bytes = editor.export
```

