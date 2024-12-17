
def ppt_context_hash(input_string):
    """
    Credits: https://learn.microsoft.com/en-gb/answers/questions/1036129/how-to-calculate-context-hash-in-ct-textcharrangec
    """
    n_hash = 0
    for ch in input_string:
        char_val = ord(ch)
        n_hash = (n_hash << 5) + n_hash + char_val
        n_hash &= 0xFFFFFFFF

    return n_hash
