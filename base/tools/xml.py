from lxml.etree import Element, SubElement

from base.data.namespaces import AbstractNameSpaces


def create_element(
        tag: str,
        root = None,
        text = None,
        namespaces: list[AbstractNameSpaces] = None,
        **attributes) -> None:
    namespaces_map = AbstractNameSpaces.to_map(namespaces) if namespaces else None
    if root is None:
        element = Element(tag, nsmap=namespaces_map)
    else:
        element = SubElement(root, tag)
    if text:
        element.text = text
    element.attrib.update(attributes)
    return element

