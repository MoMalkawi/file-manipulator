from lxml import etree

from base.components.file import ArchiveFile, ParsedArchiveFile
from base.data.components.doc import DocCommentData
from base.data.namespaces.docx import NSMAP


class DocPageHighlighter(ParsedArchiveFile):

    def __init__(self, file: ArchiveFile):
        super().__init__(file)

    def parse(self) -> DocCommentData:
        """No need"""

    def inject(self, archiver, data: DocCommentData):
        """No need"""

    def create(self, archiver, data: DocCommentData):
        """No need"""

    def highlight(self, archiver, text, comment_id):
        doc_file = archiver.get_file("word/document.xml")
        tree = etree.fromstring(doc_file.data.encode('utf-8') if isinstance(doc_file.data, str) else doc_file.data)
        namespaces = NSMAP

        text_parts = text.split("\n")
        current_part_index = 0
        start_run = None

        for paragraph in tree.xpath("//w:p", namespaces=namespaces):
            for run in paragraph.xpath(".//w:r", namespaces=namespaces):
                text_content = "".join(t.text or "" for t in run.xpath(".//w:t", namespaces=namespaces)).strip()
                if not text_content:
                    continue

                if text_parts[current_part_index] in text_content:
                    if current_part_index == 0:
                        comment_range_start = etree.Element("{%s}commentRangeStart" % namespaces["w"], id=comment_id)
                        run.addprevious(comment_range_start)
                        start_run = run

                    if text_content.strip().endswith(text_parts[current_part_index].strip()):
                        current_part_index += 1

                        if current_part_index == len(text_parts):
                            comment_range_end = etree.Element("{%s}commentRangeEnd" % namespaces["w"], id=comment_id)
                            comment_reference = etree.Element("{%s}commentReference" % namespaces["w"], id=comment_id)
                            reference_run = etree.Element("{%s}r" % namespaces["w"])
                            reference_run.append(comment_reference)
                            run.addnext(comment_range_end)
                            run.addnext(reference_run)
                            current_part_index = 0
                            start_run = None
                elif current_part_index > 0:
                    current_part_index = 0
                    if start_run:
                        start_run.getparent().remove(
                            start_run.getprevious())
                        start_run = None

        self._file = ArchiveFile(
            file_name="word/document.xml",
            data=etree.tostring(tree)
        )
        archiver.upsert("word/document.xml", lambda _: self._file)



