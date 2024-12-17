import random

import string


def generate_guid():
    hex_chars = string.digits + "ABCDEF"
    return "".join(random.choice(hex_chars) for _ in range(8))
