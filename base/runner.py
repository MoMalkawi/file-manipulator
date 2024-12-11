from base.editors.ppt.presentation import PresentationEditor


if __name__ == '__main__':
    with PresentationEditor("test_doc.pptx") as editor:
        (editor
         .get_slide(2)
         .search_and_comment(text="It worked!", text_to_highlight="thank"))
        editor.export("test.pptx")


