from base.editors.ppt.presentation import PresentationEditor

tests = {
    # "oil.pptx": [
    #     {"index": 0, "text": "It worked1", "target_text": "organic chemistry What fraction"},
    #     {"index": 0, "text": "It worked2", "target_text": "2%"},
    #     {"index": 0, "text": "It worked3", "target_text": "Crude oil"},
    #     {"index": 1, "text": "It worked4", "target_text": "2010"},
    #     {"index": 2, "text": "It worked5", "target_text": "~ 630,000"},
    #     {"index": 3, "text": "It worked6", "target_text": "https://www.atsdr.cdc.gov/toxprofiles/tp123-c2.pdf"}
    # ]
    # "algorithms1.ppt": [
    #     {"index": 0, "text": "It worked1", "target_text": "CS 202"},
    #     {"index": 0, "text": "It worked2", "target_text": "???"},
    #     {"index": 1, "text": "It worked3", "target_text": "a finite set "},
    #     {"index": 1, "text": "It worked4", "target_text": "All programs are algorithms"},
    #     {"index": 2, "text": "It worked5", "target_text": "very hard"},
    #     {"index": 24, "text": "It worked6", "target_text": "j := "}
    # ],
    # "algorithms2.ppt": [
    #     {"index": 5, "text": "It worked1", "target_text": "Read weight-in-pounds"},
    #     {"index": 7, "text": "It worked2", "target_text": "A sequence of instructions"},
    #     {"index": 9, "text": "It worked3", "target_text": "NP-complete"},
    #     {"index": 9, "text": "It worked4", "target_text": "spend our time"},
    #     {"index": 10, "text": "It worked5", "target_text": "sub-routine"},
    #     {"index": 10, "text": "It worked6", "target_text": "function"}
    # ],
    "algorithms3.pptx": [
        {"index": 6, "text": "It worked1", "target_text": "evaluated analytically"},
        # {"index": 6, "text": "It worked2", "target_text": "empirically"},
        # {"index": 6, "text": "It worked3", "target_text": "natural language"},
        # {"index": 8, "text": "It worked4", "target_text": "if/else"},
        # {"index": 12, "text": "It worked5", "target_text": "zero?"},
        # {"index": 13, "text": "It worked6", "target_text": "i.e. else"}
    ],
    # "algorithms4.ppt": [
    #     {"index": 2, "text": "It worked1", "target_text": "registration"},
    #     {"index": 2, "text": "It worked2", "target_text": "grading policy"},
    #     {"index": 6, "text": "It worked3", "target_text": "sorting"},
    #     {"index": 6, "text": "It worked4", "target_text": "4"},
    #     {"index": 13, "text": "It worked5", "target_text": "9"},
    #     {"index": 26, "text": "It worked6", "target_text": "A,B,C"}
    # ]
}

if __name__ == "__main__":
    for file_name, highlights in tests.items():
        with PresentationEditor(file_name) as editor:
            for highlight in highlights:
                editor.get_slide(highlight["index"]).search_and_comment(highlight["text"], highlight["target_text"])
            editor.export("results/" + file_name)