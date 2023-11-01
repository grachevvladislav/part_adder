from math import ceil


def zip_function(resource: int, days: int, count: int) -> int:
    factor = count * (1 / resource) * days * 24
    result = ceil(factor + 1.645 * pow(factor, 0.5))
    return result
