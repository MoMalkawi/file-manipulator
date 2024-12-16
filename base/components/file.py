from abc import ABC, abstractmethod

from base.data import BaseModel
from base.data.components import XMLFileData


class ArchiveFile:
    def __init__(self, file_name: str, data: str):
        self._file_name: str = file_name
        self._data = data

    @property
    def name(self):
        return self._file_name

    @property
    def data(self):
        return self._data


class ParsedArchiveFile(ABC):
    def __init__(self, file: ArchiveFile):
        self._file = file
        self._file_data_model: XMLFileData | None = None

    @abstractmethod
    def parse(self) -> BaseModel:
        """Locates the required data, parses it and returns it.."""

    @abstractmethod
    def inject(self, archiver, data: XMLFileData):
        """injects required info into the archived existing file."""

    @abstractmethod
    def create(self, archiver, data: XMLFileData):
        """creates the xml file from scratch (This is optional to implement, just keep it empty when implementing)"""

    @property
    def data(self):
        if not self._file_data_model:
            self._file_data_model = self.parse()
        return self._file_data_model
