from datetime import datetime


def get_highlighted_text_coords(text_to_highlight: str | None, whole_text: str):
    """Get the start index and length of the text to highlight within the shape."""
    if not text_to_highlight:
        return 0, len(whole_text)

    start_index = whole_text.lower().find(text_to_highlight.lower())
    if start_index == -1:
        raise ValueError("Text to highlight not found in the shape.")

    return start_index, len(text_to_highlight)


def generate_string_datetime():
    now = datetime.now()
    formatted_datetime = (
        f"{now.strftime('%Y-%m-%dT%H:%M:%S')}.{now.microsecond // 1000:03d}"
    )
    return formatted_datetime
