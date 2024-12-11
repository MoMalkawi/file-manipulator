
paragraphs = [
    "I SSH want do but I no know how\r",
    "\n\n"
    "I name is not bok lao, I lie in slide before.\r",
    "\n\n"
    "I you thank for reading presentation\r"
]


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


hash_value = ppt_context_hash("".join(paragraphs))
print("Computed Hash:", hash_value)
# correct val: 1448756579

