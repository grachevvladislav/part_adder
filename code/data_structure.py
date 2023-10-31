from tabulate import tabulate


class Component:
    def __init__(
        self,
        en_name: str = None,
        ru_name: str = None,
        pn_main: str = None,
        pn_opt: str = None,
        quantity: int = 1,
    ) -> None:
        self.en_name: str = en_name
        self.ru_name: str = ru_name
        self.pn_main: str = pn_main
        self.pn_opt: str = pn_opt
        self.quantity: int = quantity

    def __str__(self) -> str:
        return (
            f"{self.ru_name} | {self.en_name} | {self.pn_main} | "
            f"{self.pn_opt} | {self.quantity}\n"
        )

    def __hash__(self) -> int:
        return hash((self.ru_name, self.en_name, self.pn_main, self.pn_opt))

    def get_list(self) -> list[str]:
        return [
            self.en_name,
            self.ru_name,
            self.pn_main,
            self.pn_opt,
            self.quantity,
        ]


class Server:
    def __init__(
        self, name: str = None, sn: str = None, model: str = None,
        quantity: int = 1, config: dict = None
    ) -> None:
        self.sn: str = sn
        self.name: str = name
        self.model: str = model
        self.quantity: int = quantity
        self.config: dict = config

    def __str__(self) -> str:
        table = tabulate(
            [item.get_list() for item in self.config.values()],
            headers=["EN name", "RU name", "PN main", "PN opt", "PCS"],
        )
        return f"SN: {self.sn}\nDNS: {self.name}\nКонфигурация:\n {table}\n"

    def __hash__(self) -> int:
        return hash(tuple(self.config.keys()))

    def add_component(self, component: Component) -> None:
        key = hash(component)
        if key in self.config.keys():
            self.config[key].quantity += 1
        else:
            self.config[key] = component


class ServerSet:
    def __init__(self) -> None:
        self.collection: dict = {}

    def add_server(self, server: Server) -> None:
        key = hash(server)
        if key in self.collection.keys():
            self.collection[key].quantity += 1
        else:
            self.collection[key] = server
