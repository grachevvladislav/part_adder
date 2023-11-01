import re

from .data_structure import Component, ServerSet, Server


def get_json_data(dict_file: dict) -> ServerSet:
    servers = ServerSet()
    for server_name, server_dict in dict_file.items():
        server = Server(name=server_name, config={})
        server.add_component(
            Component(quantity=len(server_dict["fans"]), ru_name="Вентилятор")
        )
        for power_name, power_dict in server_dict["power_supplies"].items():
            if re.match(r"^Power Supply", power_name):
                power_type = "Блок питания"
            elif re.match(r"^Battery", power_name):
                power_type = "Батарея"
            else:
                power_type = None
            server.add_component(
                Component(
                    pn_main=power_dict["model"],
                    pn_opt=power_dict["spare"],
                    ru_name=power_type,
                )
            )

        for proc_dict in server_dict["processors"].values():
            server.add_component(
                Component(ru_name="Процессор", en_name=proc_dict["name"])
            )

        for cpu_mem in server_dict["memory"]["memory_details"].values():
            for mem_dict in cpu_mem.values():
                if not mem_dict["status"] == "Not Present":
                    server.add_component(
                        Component(
                            ru_name="Модуль памяти",
                            pn_main=mem_dict["part"]["number"],
                        )
                    )

        for nic_dict in server_dict["nic_information"].values():
            if not re.match(
                r"^0.v00|Network Controller", nic_dict["port_description"]
            ):
                server.add_component(
                    Component(
                        ru_name="Сетевой адаптер",
                        en_name=nic_dict["port_description"],
                    )
                )

        for array_dict in server_dict["storage"].values():
            server.add_component(
                Component(
                    ru_name="Дисковый контроллер", en_name=array_dict["model"]
                )
            )
            for disk_list in array_dict.get("logical_drives", []):
                for disk in disk_list["physical_drives"]:
                    if disk["media_type"] == "HDD":
                        name = "Жесткий диск"
                    elif disk["media_type"] == "SSD":
                        name = "Твердотельный диск"
                    else:
                        print("Неизвестный тип диска", disk["media_type"])
                        name = disk["media_type"]
                    server.add_component(
                        Component(
                            pn_opt=disk["model"],
                            en_name=(
                                f'{disk["marketing_capacity"]} '
                                f'{disk["media_type"]}'
                            ),
                            ru_name=name,
                        )
                    )
        servers.add(server)
    return servers
