import xml.etree.ElementTree as et

from base.models.parsers import AuthorData
from base.parsers.file import ParsedArchiveFile, ArchiveFile


class PPTAuthor(ParsedArchiveFile):
    """
    PPTAuthor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Parses/Injects/Creates a PowerPoint Author.
    """

    def __init__(self, author_id: str, file: ArchiveFile | None = None):
        super().__init__(file)
        self._author_id = author_id

    def parse(self):
        if not self._file:
            return None

        authors_root = et.fromstring(self._file.data)

        authors = authors_root.findall(
            ".//{http://schemas.microsoft.com/office/powerpoint/2018/8/main}author"
        )
        if authors and len(authors) > 0:
            editor_author = next(
                (
                    author
                    for author in authors
                    if author.get("id") == f"{{{self._author_id}}}"
                ),
                None,
            )
            if isinstance(editor_author, et.Element):
                return AuthorData(
                    id=editor_author.get("id")[1:-1],
                    name=editor_author.get("name"),
                    initials=editor_author.get("initials"),
                    user_id=editor_author.get("userId"),
                    provider_id=editor_author.get("providerId"),
                )
        return None

    def inject(self, archiver, data: AuthorData):
        if not self._file:
            self.create(archiver, data)
            return

        authors_root = et.fromstring(self._file.data)
        authors_root.append(self._create_author_xml(data))
        self._file = ArchiveFile(
            file_name=self._file.name,
            data=et.tostring(authors_root, encoding="unicode"),
        )
        archiver.upsert(self._file.name, lambda _: self._file)

    def create(self, archiver, data: AuthorData):
        xml_content = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<p188:authorLst "
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            f'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">'
            f'<p188:author id="{{{data.id}}}" name="{data.name}" initials="{data.initials}" '
            f'userId="{data.user_id}" providerId="{data.provider_id}"/>'
            f"</p188:authorLst>"
        )
        self._file = ArchiveFile(file_name="ppt/authors.xml", data=xml_content)
        archiver.upsert(self._file.name, lambda _: self._file)

    @classmethod
    def _create_author_xml(cls, data: AuthorData):
        return et.Element(
            "{http://schemas.microsoft.com/office/powerpoint/2018/8/main}author",
            attrib={
                "id": f"{{{data.id}}}",
                "name": data.name,
                "initials": data.initials,
                "userId": data.user_id,
                "providerId": data.provider_id,
            },
        )
