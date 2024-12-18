from abc import ABC, abstractmethod

from base.data import BaseModel


class AbstractEditor(ABC):
    """This wraps on all editors, common abstractions can be placed here."""

    def __init__(self):
        self._data: BaseModel | None = None

    @abstractmethod
    def _load_data(self, **kwargs):
        """An optional function to call when the editor deems it necessary."""

    @property
    def data(self) -> BaseModel:
        return self._data
