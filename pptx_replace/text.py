from pptx.text.text import TextFrame


def set_frame_text(text_frame: TextFrame, text: str):
    """set text and keep font style"""
    replaced = False
    for paragraph in text_frame.paragraphs:
        # use style of the first run
        if paragraph.runs:
            if not replaced:
                paragraph.runs[0].text = text
                # remove text in other runs
                replaced = True
                for run in paragraph.runs[1:]:
                    run.text = ""
            else:
                paragraph.clear()

    if not replaced:
        text_frame.text = text