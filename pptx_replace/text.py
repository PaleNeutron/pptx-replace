from pptx import TextFrame


def set_frame_text(frame: TextFrame, text: str):
    """set text and keep font style"""
    pgs = frame.paragraphs
    if pgs:
        runs = pgs[0].runs
        if runs:
            font = runs[0].font
    else:
        font = None

    frame.clear()
    frame.text = text
    if font:
        for p in frame.paragraphs:
            for r in p.runs:
                r.font = font
    return frame