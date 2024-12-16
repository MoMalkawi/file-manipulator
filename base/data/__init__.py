from enum import Enum
from typing import Any, Generic, List, TypeVar, Union

from humps import camelize
from pydantic import BaseModel as DefaultModel


class BaseModel(DefaultModel):
    class Config:
        alias_generator = camelize
        populate_by_name = True

    def serialize(self):
        return super().model_dump(by_alias=True)