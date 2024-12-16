from enum import ReprEnum


class NameSpace:

    def __init__(self, tag: str, url: str):
        self.tag = tag
        self.url = url


class AbstractNameSpaces(NameSpace, ReprEnum):

    @property
    def url(self):
        return self.value.url

    @property
    def tag(self):
        return self.value.tag

    @staticmethod
    def to_map(namespaces: list["AbstractNameSpaces"]):
        return {
            namespace.tag: namespace.url
            for namespace in namespaces
        }