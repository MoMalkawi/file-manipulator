from uuid import uuid4

from base.data.components import XMLFileData
from base.tools.strings import generate_string_datetime


class AuthorData(XMLFileData):
    id: str = "9cef6d51-21ec-46f0-9cd9-55cb7fb4ec41"
    name: str = "Document Manipulator"
    initials: str = "DM"
    user_id: str | None = "S::documentmanipulator@noemail.com::9cef6d51-21ec-46f0-9cd9-55cb7fb4ec41"
    provider_id: str | None = "AD"


class PresentationData(XMLFileData):
    slides_count: int
    author: AuthorData


class SlideData(XMLFileData):
    slide_index: int
    slide_id: str
    slide_creation_id: str | None
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
    creation_date: str = generate_string_datetime()
    shape_data: ShapeData
    highlighted_text_start_index: int
    highlighted_text_length: int
    locale: str = "en-US"
    text: str