import random
import string
import xml.etree.ElementTree as et
from datetime import datetime


def generate_ppt_datetime():
    now = datetime.now()
    formatted_datetime = (
        f"{now.strftime('%Y-%m-%dT%H:%M:%S')}.{now.microsecond // 1000:03d}"
    )
    return formatted_datetime


def ppt_context_hash(input_string):
    """Thankfully a microsoft engineer provided hints on how to hash the context string.
    7r credits: https://learn.microsoft.com/en-gb/answers/questions/1036129/how-to-calculate-context-hash-in-ct-textcharrangec
    """
    n_hash = 0
    for ch in input_string:
        char_val = ord(ch)
        n_hash = (n_hash << 5) + n_hash + char_val
        n_hash &= 0xFFFFFFFF

    return n_hash


def get_highlighted_text_coords(text_to_highlight: str | None, whole_text: str):
    """Get the start index and length of the text to highlight within the shape."""
    if not text_to_highlight:
        return 0, len(whole_text)

    start_index = whole_text.lower().find(text_to_highlight.lower())
    if start_index == -1:
        raise ValueError("Text to highlight not found in the shape.")

    return start_index, len(text_to_highlight)


def validate_element(element: any):
    # The reason for this function is that Element seems to get confused with None, which is weird.
    return isinstance(element, et.Element)


def generate_guid():
    hex_chars = string.digits + "ABCDEF"
    return "".join(random.choice(hex_chars) for _ in range(8))
