from uuid import uuid4

from base.data.components import XMLFileData
from base.utils import generate_ppt_datetime


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