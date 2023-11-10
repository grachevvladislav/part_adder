from tabulate import tabulate
from .constants import SERVICE_TIME
from .calculator import zip_function
from .data import RESOURCE, COMPONENT_NAME
from .exceptions import UnknownComponent


class Component:
    """Компонент оборудования определенного типа с указанием количества."""
    def __init__(
        self,
        en_name: str = None,
        ru_name: str = None,
        pn_main: str = None,
        pn_opt: str = None,
        comment: str = None,
        quantity: int = None
    ) -> None:
        self.for_repair: dict[str:int] = {}
        self.en_name: str = en_name
        self.ru_name: str = ru_name
        self.pn_main: str = pn_main
        self.pn_opt: str = pn_opt
        self.comment: str = comment
        self.quantity: int = quantity or 1

    def __str__(self) -> str:
        return (
            f"{self.ru_name} | {self.en_name} | {self.pn_main} | "
            f"{self.pn_opt} | {self.quantity}\n"
        )

    def __hash__(self) -> int:
        return hash((self.ru_name, self.en_name, self.pn_main))

    def get_list(self) -> list[str|int]:
        return [
            self.en_name,
            self.ru_name,
            self.pn_main,
            self.pn_opt,
            self.quantity,
        ]


class Server:
    """Уникальный сервер."""
    def __init__(
        self, name: str = None, sn: str = None, model: str = None,
        quantity: int = None, config: dict = None,
        notification: set[str] = None
    ) -> None:
        self.sn: str = sn
        self.name: str = name
        self.model: str = model
        self.quantity: int = quantity or 1
        self.config: dict = config or {}
        self.notification: set[str] = notification or set()

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
            if component.pn_main and not component.en_name:
                try:
                    component.en_name = COMPONENT_NAME[component.pn_main]
                except KeyError:
                    msg = (
                        f'Не найдено название для {component.ru_name} '
                        f'{component.pn_main}'
                    )
                    self.notification.add(msg)
            self.config[key] = component


class ServerSet:
    """Набор сгруппированных серверов."""
    def __init__(self) -> None:
        self.collection: dict = {}
        self.notification: set = set()

    def add(self, item: Server | Component) -> None:
        key = hash(item)
        if key in self.collection.keys():
            self.collection[key].quantity += 1
            self.collection[key].sn += f'\n{item.sn}'
        else:
            self.collection[key] = item
            self.notification.update(item.notification)


class ComponentSet(ServerSet):
    """Набор сгруппированных компонент."""
    def repair_calculation(self) -> None:
        for component in self.collection.values():
            for interval in SERVICE_TIME.items():
                if component.ru_name not in RESOURCE.keys():
                    raise UnknownComponent(f"Нет в словаре: {component.ru_name}")
                component.for_repair[interval[0]] = zip_function(
                    RESOURCE[component.ru_name], interval[1], component.quantity
                )
