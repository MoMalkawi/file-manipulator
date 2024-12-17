from lxml import etree

# @Anzor: let's try to use the new ../docx.py, I kept this so that you transfer what you want
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
}


class DocUtils:
    @staticmethod
    def create_skeleton():
        comments = DocUtils.create_element(
            "comments",
            nsmap=NSMAP,
            **{f'{{{NSMAP["mc"]}}}Ignorable': "w14 w15 w16se wp14"}
        )

        return comments

    @staticmethod
    def create_element(tag, parent=None, nsmap=None, text=None, **attributes):
        ns = '{' + NSMAP['w'] + '}'
        parsed_attributes = {
            etree.QName(NSMAP['w'], key.split('__')[1]) if '__' in key else key: value
            for key, value in attributes.items()
        }

        if parent is None:
            element = etree.Element(ns + tag, nsmap=nsmap or NSMAP)
        else:
            element = etree.SubElement(parent, ns + tag)

        element.attrib.update(parsed_attributes)

        if text:
            element.text = text

        return element
