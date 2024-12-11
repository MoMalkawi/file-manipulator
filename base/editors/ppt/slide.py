import logging
import xml.etree.ElementTree as et

from base.editors.ppt.presentation import PresentationEditor
from base.models.parsers import SlideData, PPTCommentData, ShapeData
from base.parsers.ppt.comments import PPTSlideComments
from base.utils import get_highlighted_text_coords, validate_element, ppt_context_hash


class SlideEditor:
    """
    SlideEditor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Edits a specific slide within a presentation.
    """

    def __init__(
        self, presentation_editor: PresentationEditor, slide_index: int, slide_id: str
    ):
        self._presentation_editor = presentation_editor
        self._index = slide_index
        self._id = slide_id
        self._data: SlideData | None = None

    def search_and_comment(
        self,
        text: str,
        text_to_highlight: str,
        shape_id: int | None = None,
        locale: str = "en-US",
    ):
        """
        Looks for the text to highlight within the slide and adds a comment to it.
        :param shape_id: id of the shape to constrict the search in.
        """
        payload = self._construct_comment_payload(
            shape_id=shape_id,
            text=text,
            text_to_highlight=text_to_highlight,
            locale=locale,
        )

        if not payload:
            logging.info("SlideEditor: No shape found for the given text, skipping comment.")
            return self

        self._add_comment(payload)
        return self

    def comment(
        self,
        text: str,
        shape_id: int,  # index of the shape in the slide
        start_index: int,  # start index of the text to highlight
        length: int,  # length of the text to highlight
        locale: str = "en-US",
    ):
        """
        Adds a comment to the slide within given location constraints.
        :param start_index: start index of the text to highlight within the shape.
        :param length: length of the text to highlight within the shape.
        """
        payload = self._construct_comment_payload(
            shape_id=shape_id,
            start_index=start_index,
            length=length,
            text=text,
            locale=locale,
        )

        if not payload:
            logging.info("SlideEditor: Couldn't locate the shape, skipping comment.")
            return self

        self._add_comment(payload)
        return self

    def _add_comment(self, comment: PPTCommentData):
        comments_file = (
            self._presentation_editor.archiver.get_file(
                f"ppt/comments/{self.data.comments_file_path}"
            )
            if self.data.comments_file_path
            else None
        )
        results_comment_file_name = PPTSlideComments(comments_file).inject(
            self._presentation_editor.archiver, comment
        )
        self._data.comments_file_path = results_comment_file_name[results_comment_file_name.rindex("/") + 1:]

    def _construct_comment_payload(self, **kwargs):
        shape = self.data.get_shape(
            shape_id=kwargs.get("shape_id"),
            search_string=kwargs.get("text_to_highlight"),
            search_case_sensitive=False,
        )
        if not shape:
            return None
        text_to_highlight = kwargs.get("text_to_highlight")
        start_index, length = get_highlighted_text_coords(text_to_highlight, shape.text)
        return PPTCommentData(
            author_id=self._presentation_editor.author_id,
            slide=self.data,
            shape=shape,
            highlighted_text_start_index=kwargs.get("start_index") or start_index,
            highlighted_text_length=kwargs.get("length") or length,
            text=kwargs.get("text"),
            locale=kwargs.get("locale"),
        )

    def get_slide_comments(self):
        slide_rel = self._presentation_editor.archiver.get_file(
            f"ppt/slides/_rels/slide{self._index + 1}.xml.rels"
        )
        rel_xml = et.fromstring(slide_rel.data)
        relationships = rel_xml.findall(
            ".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        return next(
            (
                rel.get("Target")[rel.get("Target").rindex("/") + 1 :]
                for rel in relationships
                if rel.get("Type", "").endswith("comments")
            ),
            None,
        )

    def get_slide_creation_date(self):
        slide = self._presentation_editor.archiver.get_file(
            f"ppt/slides/slide{self._index + 1}.xml"
        )
        external_list = et.fromstring(slide.data).findall(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}extLst"
        )
        for ext in external_list:
            nested_ext_list = ext.findall(
                ".//{http://schemas.openxmlformats.org/presentationml/2006/main}ext"
            )
            for nested_ext in nested_ext_list:
                slide_root = nested_ext.findall(
                    ".//{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId"
                )
                if slide_root:
                    return next((root.get("val") for root in slide_root), None)
        return None

    def get_slide_shapes(self) -> list[ShapeData]:
        result = []
        slide = self._presentation_editor.archiver.get_file(
            f"ppt/slides/slide{self._index + 1}.xml"
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
                    shape_data = self._extract_shape_data(shape)
                    if shape_data:
                        result.append(shape_data)
        return result

    @classmethod
    def _extract_shape_data(
        cls, shape: et.Element
    ) -> ShapeData | None:  # TODO: refactor later.
        texts = []
        text_length = 0
        shape_creation_id = ""
        body = shape.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}txBody"
        )
        if validate_element(body):
            paragraphs = body.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}p"
            )
            for paragraph in paragraphs:
                if validate_element(paragraph):
                    run = paragraph.find(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}r"
                    )
                    if validate_element(run):
                        text_box = run.find(
                            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}t"
                        )
                        if text_box:
                            texts.append(text_box.text + "\r")
                            text_length += len(text_box.text) + 1
                    elif validate_element(paragraph.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")):
                        texts.append("\r")
                        text_length += 1

        shape_metadata = shape.find(
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
                            shape_creation_id = shape_creation_id_comp.get("id")[1:-1]
                            break

                return ShapeData(
                    shape_id=int(shape_id.get("id")),
                    shape_creation_id=shape_creation_id,
                    text="\n".join(texts),
                    name=shape_id.get("name"),
                    text_area_length=str(text_length),
                    text_area_content_hash=str(ppt_context_hash("".join(texts)))
                )
        return None

    @property
    def data(self) -> SlideData:
        if not self._data:
            self._data = SlideData(
                slide_index=self._index,
                slide_id=self._id,
                comments_file_path=self.get_slide_comments(),
                slide_creation_id=self.get_slide_creation_date(),
                shapes=self.get_slide_shapes(),
            )
        return self._data
