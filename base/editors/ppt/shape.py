import logging
import xml.etree.ElementTree as et

from xml.etree.ElementTree import Element

from base.components.ppt.comments import PPTSlideComments
from base.data.components.ppt import SlideData, ShapeData, PPTCommentData
from base.editors.archive import SelectiveArchiveEditor
from base.utils import validate_element, ppt_context_hash, get_highlighted_text_coords


class ShapeEditor:
    """
    ShapeEditor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Edits a specific shape within a slide.
    """

    def __init__(
        self, identifier: str, slide_data: SlideData, archiver: SelectiveArchiveEditor
    ):
        self._id = identifier
        self._slide_data = slide_data
        self._archiver = archiver
        self._data: ShapeData | None = self._load_data()

    def search_and_comment(
        self,
        text: str,
        text_to_highlight: str,
        locale: str = "en-US",
    ):
        """
        Looks for the text to highlight within the shape and adds a comment to it.
        """
        payload = self._construct_comment_payload(
            text=text,
            text_to_highlight=text_to_highlight,
            locale=locale,
        )

        self._add_comment(payload)
        return self

    def comment(
        self,
        text: str,
        start_index: int,  # start index of the text to highlight
        length: int,  # length of the text to highlight
        locale: str = "en-US",
    ):
        """
        Adds a comment to the shape within given location constraints.
        :param start_index: start index of the text to highlight within the shape.
        :param length: length of the text to highlight within the shape.
        """
        payload = self._construct_comment_payload(
            start_index=start_index,
            length=length,
            text=text,
            locale=locale,
        )

        self._add_comment(payload)
        return self

    def _construct_comment_payload(self, **kwargs):
        text_to_highlight = kwargs.get("text_to_highlight")
        start_index, length = get_highlighted_text_coords(text_to_highlight, self._data.text)
        return PPTCommentData(
            author_id=self._slide_data.presentation_data.author.id,
            shape_data=self._data,
            highlighted_text_start_index=kwargs.get("start_index") or start_index,
            highlighted_text_length=kwargs.get("length") or length,
            text=kwargs.get("text"),
            locale=kwargs.get("locale"),
        )

    def _add_comment(self, comment: PPTCommentData):
        comments_file = (
            self._archiver.get_file(
                f"ppt/comments/{self._data.slide_data.comments_file_path}"
            )
            if self._data.slide_data.comments_file_path
            else None
        )
        results_comment_file_name = PPTSlideComments(comments_file).inject(
            self._archiver, comment
        )
        self._data.slide_data.comments_file_path = results_comment_file_name[results_comment_file_name.rindex("/") + 1:]

    def _load_data(self) -> ShapeData | None:
        shape_element = self._extract_element()
        if not validate_element(shape_element):
            logging.error("ShapeEditor: Could not find shape.")
            return None

        shape_metadata = self._extract_metadata(shape_element)
        if not shape_metadata:
            logging.error("ShapeEditor: Could not extract metadata.")
            return None

        shape_content = self._extract_content(shape_element)
        if not shape_content:
            logging.error("ShapeEditor: Could not extract content.")
            return None

        return ShapeData(
            id=shape_metadata.get("id"),
            creation_id=shape_metadata.get("creation_id"),
            name=shape_metadata.get("name"),
            text=shape_content.get("text"),
            text_area_length=shape_content.get("text_area_length"),
            text_area_content_hash=shape_content.get("text_area_content_hash"),
            slide_data=self._slide_data,
        )

    def _extract_content(self, shape_element: Element) -> dict | None:  # noqa
        texts = []
        text_length = 0
        body = shape_element.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}txBody"
        )
        if validate_element(body):
            paragraphs = body.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}p"
            )
            for paragraph in paragraphs:
                if validate_element(paragraph):
                    runs = paragraph.findall(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}r"
                    )
                    if len(runs) > 0:
                        runs_text = ""
                        for run in runs:
                            text_box = run.find(
                                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}t"
                            )
                            if validate_element(text_box):
                                runs_text += text_box.text
                        text_length += len(runs_text) + 1
                        texts.append(runs_text + "\r")
                    elif validate_element(paragraph.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")):
                        texts.append("\r")
                        text_length += 1
            return dict(
                text="".join([text.replace("\r", "\n").replace('\u00A0', ' ') for text in texts]),
                text_area_length=str(text_length),
                text_area_content_hash=str(ppt_context_hash("".join(texts)))
            )
        return None

    def _extract_metadata(self, shape_element: Element) -> dict | None:
        if not shape_element:
            return None
        shape_metadata = shape_element.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr"
        )
        if validate_element(shape_metadata):
            shape_id = shape_metadata.find(
                ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"
            )
            if validate_element(shape_id):
                shape_ext = shape_metadata.find(
                    ".//{http://schemas.openxmlformats.org/drawingml/2006/main}extLst"
                )
                if validate_element(shape_ext):
                    ext_list = shape_ext.findall(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}ext"
                    )
                    for ext in ext_list:
                        shape_creation_id_comp = ext.find(
                            ".//{http://schemas.microsoft.com/office/drawing/2014/main}creationId"
                        )
                        if validate_element(shape_creation_id_comp):
                            return dict(
                                id=self._id,
                                creation_id=shape_creation_id_comp.get("id")[1:-1],
                                name=shape_id.get("name")
                            )
                else:
                    return dict(
                        id=self._id,
                        creation_id="00000000-0000-0000-0000-000000000000",
                        name=shape_id.get("name")
                    )
        return None

    def _extract_element(self) -> Element | None:
        """
        Get the shape element from the slide XML.

        Note: The reason this is semi-duplicated is to force accessing the files.
        """
        slide = self._archiver.get_file(
            f"ppt/slides/slide{self._slide_data.slide_index + 1}.xml"
        )
        slide_xml = et.fromstring(slide.data)
        creation_slide = slide_xml.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
        )
        if validate_element(creation_slide):
            shape_tree = creation_slide.find(
                ".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree"
            )
            if validate_element(shape_tree):
                shapes = shape_tree.findall(
                    ".//{http://schemas.openxmlformats.org/presentationml/2006/main}sp"
                )
                for shape in shapes:
                    shape_metadata = shape.find(
                        ".//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr"
                    )
                    if validate_element(shape_metadata):
                        shape_id = shape_metadata.find(
                            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"
                        )
                        if validate_element(shape_id):
                            if shape_id.get("id") == self._id:
                                return shape
        return None

    @property
    def data(self) -> ShapeData:
        return self._data
