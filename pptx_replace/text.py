from pptx.text.text import TextFrame


def delete_paragraph(paragraph):
    p = paragraph._p
    parent_element = p.getparent()
    parent_element.remove(p)


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
                delete_paragraph(paragraph)

    if not replaced:
        text_frame.text = text