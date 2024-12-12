from lxml import etree
from base.editors.archive import SelectiveArchiveEditor
from base.models.parsers import DocCommentData
from base.namespaces.docx import DocUtils
from base.parsers.file import ArchiveFile, ParsedArchiveFile


class DocPageComments(ParsedArchiveFile):


    def __init__(self, file: ArchiveFile):
        super().__init__(file)

    def parse(self) -> DocCommentData:
        """No need to parse this, as we currently don't need to read comments, only insert them ;)"""

    def inject(self, archiver, data: DocCommentData):
        return self.create(archiver, data)

    def create(self, archiver, data: DocCommentData):
        comments_file_name = self._create_comments_file(archiver, data)
        return comments_file_name

    def _create_comments_file(self, archiver, data):
        comments_file = self.__create_comments_file(archiver)
        self._create_comment_xml(data, comments_file)
        self._file = ArchiveFile(
            file_name=comments_file,
            data=etree.tostring(comments_file)
        )

        archiver.upsert("word/comments.xml", lambda _: self._file)
        return comments_file

    def __create_comments_file(self, archiver: SelectiveArchiveEditor):
        comments_file = archiver.get_filenames_in_dir("word")

        if "word/comments.xml" not in comments_file:
            archiver.upsert("word/comments.xml", lambda _: self._file)
        file = archiver.get_file("word/comments.xml")
        if file:
            print(type(file.data))
            tree = etree.fromstring(file.data.encode("utf-8") if isinstance(file.data, str) else file.data)
        else:
            tree = self._create_comment_file_skeleton()
        return tree

    @classmethod
    def _create_comment_xml(cls, data: DocCommentData, skeleton) -> str:
        return data.to_xml(skeleton)

    @classmethod
    def _create_comment_file_skeleton(cls):
        return DocUtils.create_skeleton()
