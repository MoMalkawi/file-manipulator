from abc import ABC
from uuid import uuid4

from base.models import BaseModel
from base.utils import generate_ppt_datetime


class XMLFileData(BaseModel, ABC):
    ...


class AuthorData(XMLFileData):
    id: str
    name: str
    initials: str
    user_id: str
    provider_id: str


class ShapeData(XMLFileData):
    shape_id: int
    shape_creation_id: str
    text: str
    name: str
    text_area_length: int
    text_area_content_hash: str


class SlideData(XMLFileData):
    slide_index: int
    slide_id: str
    slide_creation_id: str
    comments_file_path: str | None
    shapes: list[ShapeData]

    def get_shape(
        self,
        shape_id: int | None = None,
        search_string: str | None = None,
        search_case_sensitive: bool = True,
        exact_search: bool = False,
    ) -> ShapeData | None:
        if shape_id:
            return next(
                (shape for shape in self.shapes if shape.shape_id == shape_id), None
            )

        if search_string:
            search_string = (
                search_string if search_case_sensitive else search_string.lower()
            )
            for shape in self.shapes:
                shape_text = shape.text if search_case_sensitive else shape.text.lower()
                if (exact_search and search_string == shape_text) or (
                    not exact_search and search_string in shape_text
                ):
                    return shape

        return None


class PPTCommentData(XMLFileData):
    comment_id: str = str(uuid4())
    author_id: str  # (IMPORTANT)
    creation_date: str = generate_ppt_datetime()
    slide: SlideData
    shape: ShapeData
    highlighted_text_start_index: int  # ac:txMk => cp  (IMPORTANT)  obvious
    highlighted_text_length: int  # ac:txMk => len  (IMPORTANT)  obvious
    locale: str = "en-US"
    text: str
    # text_area_length: int  # ac:txMk:ac:context => len  (IMPORTANT) (LENGTH OF TEXT IN SHAPE + 1)
    # text_area_content_hash: str  # ac:txMk:ac:context => hash  (IMPORTANT)  found the algorithm
    # slide_id: str  # sldId  (IMPORTANT)  # <p:sldId id="256" r:id="rId2"/> in presentation.xml (They are ordered by creation date so 1 2 4 3)
    # slide_creation_id: str  # cId  (IMPORTANT)  val="" in slideX.xml (tag=p14:creationId)
    # shape_id: int = 3  # in slideX.xml (it wraps the shape in the text (unique)) <p:cNvPr id="3" name="Subtitle 2">
    # shape_creation_id: str  # in slideX.xml (IT WRAPS THE TEXT BOX (ITS UNIQUE to the text box)) <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{257E15C3-747C-13D4-A229-AF3980E40962}"/>
