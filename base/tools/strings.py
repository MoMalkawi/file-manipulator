from datetime import datetime


def locate_text_in_text(text_to_locate: str, whole_text: str):
    """Get the start index and length of the text to locate within the whole text."""
    start_index = whole_text.lower().find(text_to_locate.lower())
    return start_index, len(text_to_locate)


def locate_text_in_texts(text_to_locate: str,
                         texts: list[str],
                         case_sensitive: bool = False,
                         space_delimit: bool = False) -> list[dict]:
    if not case_sensitive:
        texts = [text.lower() for text in texts]
    all_text = "".join(texts)
    if space_delimit:
        all_text = all_text.replace("\n", " ")
    start_idx = all_text.find(text_to_locate)
    if start_idx == -1:
        return []
    end_idx = start_idx + len(text_to_locate)
    cumulative_length = 0
    remaining_start = start_idx
    remaining_end = end_idx
    results = []
    for i, text in enumerate(texts):
        length = len(text)
        global_start = cumulative_length
        global_end = cumulative_length + length
        if global_end > remaining_start and global_start < remaining_end:
            local_start = max(0, remaining_start - global_start)
            local_end = min(length, remaining_end - global_start)
            highlighted_substring = text[local_start:local_end]
            results.append({
                "index": i,
                "text": text,
                "highlighted_substring": highlighted_substring,
                "local_start": local_start,
                "length": local_end - local_start
            })
        cumulative_length += length
        if cumulative_length >= end_idx:
            break
    return results

def generate_string_datetime():
    now = datetime.now()
    formatted_datetime = (
        f"{now.strftime('%Y-%m-%dT%H:%M:%S')}.{now.microsecond // 1000:03d}"
    )
    return formatted_datetime
