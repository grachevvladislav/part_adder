from code.resource_data import RESOURCE
from math import ceil


def zip_function(resource: int, days: int, count: int) -> int:
    factor = count * (1 / resource) * days * 24
    result = ceil(factor + 1.645 * pow(factor, 0.5))
    return result


def zip_calculation(data_counter, data_info):
    resource_dict = {}
    zip_quantity = {}
    for keys, value in RESOURCE.items():
        for key in keys:
            resource_dict[key] = value
    for pn in data_info.keys():
        name = data_info[pn].ru_name
        if name not in resource_dict.keys():
            print(f"Нет в словаре: {name}")
            break
        one_year = zip_function(resource_dict[name], 365, data_counter[pn])
        three_year = zip_function(resource_dict[name], 365 * 3, data_counter[pn])
        zip_quantity[pn] = (one_year, three_year)
    return zip_quantity
