from base.editors.ppt.presentation import PresentationEditor


if __name__ == '__main__':
    with PresentationEditor("../testing/empty.pptx") as editor:
        (editor
         .get_slide(0)
         .search_and_comment(text="It worked!", text_to_highlight="about me"))
        # (editor
        #  .get_slide(1)
        #  .search_and_comment(text="It worked!", text_to_highlight="document provides"))
        # (editor
        #  .get_slide(2)
        #  .search_and_comment(text="It worked!", text_to_highlight="pwc"))
        # (editor
        #  .get_slide(4)
        #  .search_and_comment(text="It worked!", text_to_highlight="culture of giving"))
        # (editor
        #  .get_slide(5)
        #  .search_and_comment(text="It worked!", text_to_highlight="NPO services"))
        # (editor
        #  .get_slide(5)
        #  .search_and_comment(text="It worked!", text_to_highlight="told us"))
        # (editor
        #  .get_slide(11)
        #  .search_and_comment(text="It worked!", text_to_highlight="monitor water"))
        # (editor
        #  .get_slide(13)
        #  .search_and_comment(text="It worked!", text_to_highlight="accelerators"))
        # (editor
        #  .get_slide(15)
        #  .search_and_comment(text="It worked!", text_to_highlight="1,000"))
        # (editor
        #  .get_slide(15)
        #  .search_and_comment(text="It worked!", text_to_highlight="1,000"))
        # (editor
        #  .get_slide(15)
        #  .search_and_comment(text="It worked!", text_to_highlight="1,000"))
        editor.export("../testing/test.pptx")


