
paragraphs = [
    "This is commented.\r\r",
    "This isn’t commented 1\r",
    "This isn’t commented 2\r"
]


def ppt_context_hash(input_string):

    n_hash = 0
    for ch in input_string:
        char_val = ord(ch)
        n_hash = (n_hash << 5) + n_hash + char_val
        n_hash &= 0xFFFFFFFF

    return n_hash


hash_value = ppt_context_hash("".join(paragraphs))
print("Computed Hash:", hash_value)
# correct val: 133416203

