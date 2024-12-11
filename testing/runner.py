from base.editors.ppt.presentation import PresentationEditor


if __name__ == '__main__':
    with PresentationEditor("../testing/test_doc.pptx") as editor:
        (editor
         .get_slide(3)
         .search_and_comment(text="It worked!", text_to_highlight="commented 1"))
        editor.export("../testing/test.pptx")


