import re


def generate_sequential_number(base_num, index):
    """093/... -> 093.1/..., 093.2/..., etc."""
    match = re.match(r'^(\d+)(.*)', base_num)
    if match:
        prefix = match.group(1)
        suffix = match.group(2)
        return f"{prefix}.{index + 1}{suffix}"
    return f"{base_num}.{index + 1}"
