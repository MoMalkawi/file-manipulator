import xml.etree.ElementTree as et
from uuid import uuid4

from base.models.parsers import PPTCommentData
from base.parsers.file import ArchiveFile, ParsedArchiveFile
from base.utils import generate_guid
from base.models import BaseModel


class PPTSlideComments(ParsedArchiveFile):
    """
    PPTSlideComments

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Parses/Injects/Creates Slide Comments.
    """

    def __init__(self, file: ArchiveFile):
        super().__init__(file)

    def parse(self) -> BaseModel:
        """No need to parse this, as we currently don't need to read comments, only insert them ;)"""

    def inject(self, archiver, data: PPTCommentData):
        if not self._file:
            return self.create(archiver, data)
        comment_xml = self._create_comment_xml(data)
        self._file = ArchiveFile(
            file_name=self._file.name,
            data=self._file.data.replace(
                "</p188:cmLst>", f"{comment_xml}</p188:cmLst>"
            ),
        )
        archiver.upsert(self._file.name, lambda _: self._file)
        return self._file.name

    def create(self, archiver, data: PPTCommentData):
        comments_file_name = self._create_comments_file(archiver, data)
        self._add_comments_relationship(archiver, data)
        # self._link_relation_to_slide(archiver, data, relation_id)
        return comments_file_name

    @classmethod
    def _link_relation_to_slide(cls, archive, data, relation_id: int):
        slide_file_name = f"ppt/slides/slide{data.slide.slide_index + 1}.xml"
        slide = archive.get_file(slide_file_name)
        uri_index = slide.data.index("uri=")
        uri = slide.data[
            uri_index + 5 : slide.data[uri_index:].index("}") + uri_index + 1
        ]
        rel_link_xml = f"""<p:extLst><p:ext uri="{uri}"><p188:commentRel xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" r:id="rId{relation_id}"/></p:ext></p:extLst></p:sld>"""
        new_slide = slide.data.replace("</p:sld>", rel_link_xml)
        archive.upsert(
            slide_file_name,
            lambda _: ArchiveFile(file_name=slide_file_name, data=new_slide),
        )

    def _add_comments_relationship(self, archiver, data) -> int:
        slide_rel_file_name = (
            f"ppt/slides/_rels/slide{data.shape_data.slide_data.slide_index + 1}.xml.rels"
        )
        slide_rel = archiver.get_file(slide_rel_file_name)
        slide_rel_xml = et.fromstring(slide_rel.data)
        relationships = slide_rel_xml.findall(
            ".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        relationship_id = 1
        if relationships:
            relationship_id = max([int(rel.get("Id")[3:]) for rel in relationships]) + 1
        comments_file = self._file.name[self._file.name.rindex("/") + 1 :]
        new_slide_rel = slide_rel.data.replace(
            "</Relationships>",
            f"""<Relationship Id="rId{relationship_id}" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/{comments_file}"/>
        </Relationships>""",
        )  # noqa
        archiver.upsert(
            slide_rel_file_name,
            lambda _: ArchiveFile(file_name=slide_rel_file_name, data=new_slide_rel),
        )
        return relationship_id

    def _create_comments_file(self, archiver, data):
        new_comments_file_index = self._create_comments_file_index(archiver)
        new_comments_file_guid = generate_guid()
        file_name = f"modernComment_{new_comments_file_index}_{new_comments_file_guid}.xml"
        comments_file_name = f"ppt/comments/{file_name}"
        comments_skeleton = self._create_comment_file_skeleton()
        comment_xml = self._create_comment_xml(data)
        self._file = ArchiveFile(
            file_name=comments_file_name,
            data=comments_skeleton.replace(
                "</p188:cmLst>", f"{comment_xml}</p188:cmLst>"
            ),
        )
        archiver.upsert(comments_file_name, lambda _: self._file)
        return comments_file_name

    @classmethod
    def _create_comments_file_index(cls, archiver):
        all_comments_files = archiver.get_filenames_in_dir("ppt/comments")
        if not all_comments_files:
            return 100
        indices = [int(file.split("_")[1]) for file in all_comments_files]
        return max(indices) + 1

    @classmethod
    def _create_comment_xml(cls, data: PPTCommentData) -> str:
        return f"""<p188:cm id="{{{str(uuid4())}}}" authorId="{{{data.author_id}}}" created="{data.creation_date}">
    <ac:txMkLst
        xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command">
        <pc:docMk
            xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"/>
            <pc:sldMk
                xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" cId="{data.shape_data.slide_data.slide_creation_id}" sldId="{data.shape_data.slide_data.slide_id}"/>
                <ac:spMk id="{data.shape_data.id}" creationId="{{{data.shape_data.creation_id}}}"/>
                <ac:txMk cp="{data.highlighted_text_start_index}" len="{data.highlighted_text_length}">
                    <ac:context len="{data.shape_data.text_area_length}" hash="{data.shape_data.text_area_content_hash}"/>
                </ac:txMk>
            </ac:txMkLst>
            <p188:pos x="0" y="0"/>
            <p188:txBody>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                    <a:r>
                        <a:rPr lang="{data.locale}"/>
                        <a:t>{data.text}</a:t>
                    </a:r>
                </a:p>
            </p188:txBody>
        </p188:cm>"""

    @classmethod
    def _create_comment_file_skeleton(cls):
        return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"></p188:cmLst>"""  # noqa
