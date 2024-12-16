import logging
import xml.etree.ElementTree as et

from base.components.file import ParsedArchiveFile, ArchiveFile
from base.data.components.ppt import AuthorData


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
            logging.info("PPTAuthor: Couldn't parse Author file, as it doesn't exist.")
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
        self._add_author_content_type(archiver)
        self._add_presentation_rel(archiver)
        namespace = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"

        root = et.Element(f"{{{namespace}}}authorLst")

        author_attrib = {
            "id": f"{{{data.id}}}",
            "name": data.name,
            "initials": data.initials,
            "userId": data.user_id,
            "providerId": data.provider_id
        }
        et.SubElement(root, f"{{{namespace}}}author", attrib=author_attrib)

        xml_bytes = et.tostring(root, encoding="utf-8", xml_declaration=True)
        xml_content = xml_bytes.decode("utf-8")
        self._file = ArchiveFile(file_name="ppt/authors.xml", data=xml_content)
        archiver.upsert(self._file.name, lambda _: self._file)

    # @classmethod
    # def _add_presentation_rel(cls, archiver):
    #     presentation_rel = archiver.get_file("ppt/_rels/presentation.xml.rels")
    #     xml_content = presentation_rel.data
    #     presentation_rel_xml = et.fromstring(xml_content)
    #     print(presentation_rel_xml.findall("Relationship"))

    @classmethod
    def _add_presentation_rel(cls, archiver):
        presentation_rel = archiver.get_file("ppt/_rels/presentation.xml.rels")
        xml_content = presentation_rel.data
        root = et.fromstring(xml_content)
        ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        rel_tag = f"{{{ns}}}Relationship"

        rids = []
        for rel in root.findall(rel_tag):
            rid = rel.get("Id")
            if rid.startswith("rId"):
                num = int(rid[3:])
                rids.append(num)

        highest_rid = max(rids) if rids else 0

        authors_exists = any(
            rel.get("Type") == "http://schemas.microsoft.com/office/2018/10/relationships/authors"
            and rel.get("Target") == "authors.xml"
            for rel in root.findall(rel_tag)
        )

        if not authors_exists:
            new_rid = f"rId{highest_rid + 1}"
            new_rel = et.SubElement(root, rel_tag)
            new_rel.set("Id", new_rid)
            new_rel.set("Type", "http://schemas.microsoft.com/office/2018/10/relationships/authors")
            new_rel.set("Target", "authors.xml")

        new_rel_file = ArchiveFile(file_name=presentation_rel.name,
                                   data=et.tostring(root, encoding="utf-8").decode("utf-8"))
        archiver.upsert(new_rel_file.name, lambda _: new_rel_file)

    @classmethod
    def _add_author_content_type(cls, archiver):
        content_types_file = archiver.get_file("[Content_Types].xml")
        if not content_types_file or not content_types_file.data:
            # TODO: handle this later on.
            return
        xml_content = content_types_file.data
        root = et.fromstring(xml_content)
        ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        override_tag = f"{{{ns}}}Override"

        authors_exists = any(
            el.get("PartName") == "/ppt/authors.xml"
            and el.get("ContentType") == "application/vnd.ms-powerpoint.authors+xml"
            for el in root.findall(override_tag)
        )

        if not authors_exists:
            new_el = et.SubElement(root, override_tag)
            new_el.set("PartName", "/ppt/authors.xml")
            new_el.set("ContentType", "application/vnd.ms-powerpoint.authors+xml")

        new_content_types_file = ArchiveFile(file_name=content_types_file.name,
                                             data=et.tostring(root, encoding="utf-8").decode("utf-8"))
        archiver.upsert(new_content_types_file.name, lambda _: new_content_types_file)

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
