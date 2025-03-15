from base.editors.ppt.presentation import PresentationEditor


if __name__ == '__main__':
    with PresentationEditor("../testing/manual_comments.pptx") as editor:
        (editor
         .get_slide(0)
         .search_and_comment(text="It worked!", text_to_highlight="about me"))
        # (editor
        #  .get_slide(1)
        #  .search_and_comment(text="It worked!", text_to_highlight="document provides"))
        editor.export("../testing/test.pptx")


