from base.data.namespaces import AbstractNameSpaces, NameSpace


class DocxNameSpaces(AbstractNameSpaces):
    MAIN = NameSpace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    WORDPROCESSING_CANVAS = NameSpace('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas')
    CHART_EX = NameSpace('cx', 'http://schemas.microsoft.com/office/drawing/2014/chartex')
    CHART_EX_2015 = NameSpace('cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex')
    MARKUP_COMPATIBILITY = NameSpace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    OFFICE = NameSpace('o', 'urn:schemas-microsoft-com:office:office')
    RELATIONSHIP = NameSpace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    MATH = NameSpace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    VML = NameSpace('v', 'urn:schemas-microsoft-com:vml')
    WORDPROCESSING_DRAWING_2010 = NameSpace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')
    WORDPROCESSING_DRAWING = NameSpace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
    WORD_OFFICE = NameSpace('w10', 'urn:schemas-microsoft-com:office:word')
    WORD_2010 = NameSpace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
    WORD_2012 = NameSpace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
    WORD_SYMEX = NameSpace('w16se', 'http://schemas.microsoft.com/office/word/2015/wordml/symex')
    WORDPROCESSING_GROUP = NameSpace('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')
    WORDPROCESSING_INK = NameSpace('wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk')
    WORD_2006 = NameSpace('wne', 'http://schemas.microsoft.com/office/word/2006/wordml')
    WORDPROCESSING_SHAPE = NameSpace('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
