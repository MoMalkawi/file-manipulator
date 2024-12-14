from abc import ABC
from datetime import datetime, timezone
from uuid import uuid4

from base.models import BaseModel
from base.utils import generate_ppt_datetime


class XMLFileData(BaseModel, ABC):
    ...


class AuthorData(XMLFileData):
    id: str
    name: str
    initials: str
    user_id: str | None
    provider_id: str | None


class PresentationData(XMLFileData):
    slides_count: int
    author: AuthorData


class SlideData(XMLFileData):
    slide_index: int
    slide_id: str
    slide_creation_id: str
    comments_file_path: str | None
    presentation_data: PresentationData


class ShapeData(XMLFileData):
    id: int
    creation_id: str
    name: str
    text: str
    text_area_length: int
    text_area_content_hash: str
    slide_data: SlideData


class PPTCommentData(XMLFileData):
    comment_id: str = str(uuid4())
    author_id: str
    creation_date: str = generate_ppt_datetime()
    shape_data: ShapeData
    highlighted_text_start_index: int
    highlighted_text_length: int
    locale: str = "en-US"
    text: str

class DocCommentData(XMLFileData):
    # w:comment
    initials: str = "DM"  # w:initials
    creation_date: str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")  # w:date
    comment_id: str = "0"  # w:id
    author: str = "Document Manipulator"  # w:author
    text: str  #  w:t

    def to_xml(self, skeleton):
        from base.namespaces.docx import DocUtils

        comment = DocUtils.create_element(
            "comment",
            parent=skeleton,
            w__initials=self.initials,
            w__date=self.creation_date,
            w__id=self.comment_id,
            w__author=self.author
        )

        p = DocUtils.create_element("p", parent=comment)

        r = DocUtils.create_element("r", parent=p)

        DocUtils.create_element("t", parent=r, text=self.text)

        return comment
