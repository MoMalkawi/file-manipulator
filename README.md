# File Manipulator

Example **PresentationEditor** usage:
```
with PresentationEditor("absolute_file_path") as editor:  # absolute_file_path is optional, you can include it or the origin presentation bytes.
  editor.get_slide(slide_index).comment(...)
  new_presentation_bytes = editor.export(absolute_export_path)  # absolute_export_path is optional
```

Example **SelectiveArchiveEditor** usage:
```
with SelectiveArchiveEditor("absolute_file_path") as editor:  # absolute_file_path is optional, you can include it or the origin file bytes.
  new_file_bytes = editor.upsert(...).upsert_many(...).export(absolute_file_path)  absolute_export_path is optional
```
