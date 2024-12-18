from lxml import etree

from base.components.doc.comments import DocPageComments
from base.components.doc.page import DocPageHighlighter
from base.data.components.doc import DocCommentData
from base.editors.archive import AbstractArchiveEditor
from base.data.namespaces.docx import NSMAP


class DocumentEditor(AbstractArchiveEditor):

    def __init__(self, doc: str | bytes):
        super().__init__(doc)

    def get_page_count(self):
        # TODO: Estimate page number based on number of paragraphs as well.
        doc = self.archiver.get_file("word/document.xml")
        doc_data = doc.data.encode('utf-8')
        root = etree.fromstring(doc_data)
        namespaces = NSMAP
        page_breaks = root.xpath('.//w:br[@w:type="page"]', namespaces=namespaces)
        section_properties = root.xpath('.//w:sectPr', namespaces=namespaces)
        return max(len(page_breaks) + 1, len(section_properties) + 1)

    def get_comments_text(self):
        comments_file = self.archiver.get_file(f"word/comments.xml")
        tree = etree.fromstring(comments_file.data.encode("utf-8"))
        namespaces = NSMAP
        comments = tree.xpath('//w:p', namespaces=namespaces)
        _comments = []
        for i, comment in enumerate(comments, 1):
            texts = comment.xpath('.//w:t/text()', namespaces=namespaces)
            _comments.append('\n'.join(texts).strip())
        return _comments

    def get_paragraphs(self):
        comments_file = self.archiver.get_file(f"word/document.xml")
        tree = etree.fromstring(comments_file.data.encode("utf-8"))
        namespaces = NSMAP
        paragraphs = tree.xpath('//w:p', namespaces=namespaces)
        _paragraphs = []
        for i, paragraph in enumerate(paragraphs, 1):
            texts = paragraph.xpath('.//w:t/text()', namespaces=namespaces)
            _paragraphs.append('\n'.join(texts).strip())
        return _paragraphs

    def get_comment_id(self):
        comments_file = self.archiver.get_file(f"word/comments.xml")
        tree = etree.fromstring(comments_file.data.encode("utf-8") if isinstance(comments_file.data, str) else comments_file.data)
        namespaces = NSMAP
        comments = tree.xpath('//w:comment', namespaces=namespaces)
        comment_id = max([comment.attrib[f'{{{NSMAP.get("w")}}}id'] for comment in comments])
        return str(int(comment_id) + 1)

    def add_comment(self, text: str, data: DocCommentData):
        parser = DocPageComments(self.archiver.get_file("word/comments.xml"))
        data.comment_id = self.get_comment_id()
        parser.inject(self.archiver, data)
        DocPageHighlighter(self.archiver.get_file("word/document.xml")).highlight(self.archiver, text, data.comment_id)

    @property
    def archiver(self):
        return self._archive

    def _load_data(self, **kwargs):
        """TODO: Store the DocumentEditor Data"""
