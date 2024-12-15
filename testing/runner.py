from base.editors.ppt.presentation import PresentationEditor


if __name__ == '__main__':
    with PresentationEditor("../testing/new_original2.pptx") as editor:
        (editor
         .get_slide(2)
         .search_and_comment(text="It worked!", text_to_highlight="enhancing communication between QA"))
        editor.export("../testing/test.pptx")


