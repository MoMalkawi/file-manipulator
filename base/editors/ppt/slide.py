import logging
import xml.etree.ElementTree as et

from base.data.components.ppt import PresentationData, SlideData
from base.editors.archive import SelectiveArchiveEditor
from base.editors.ppt.shape import ShapeEditor
from base.tools.xmls import validate_element


class SlideEditor:
    """
    SlideEditor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Edits a specific slide within a presentation.
    """

    def __init__(
        self, index: int, presentation_data: PresentationData, archiver: SelectiveArchiveEditor
    ):
        self._archiver = archiver
        self._index = index
        self._presentation_data: PresentationData = presentation_data
        self._data: SlideData = self._load_data()

    def _load_data(self) -> SlideData:
        return SlideData(
            slide_index=self._index,
            slide_id=self._extract_slide_id(self._index),
            comments_file_path=self._extract_comment_file_name(),
            slide_creation_id=self._extract_slide_creation_id(),
            presentation_data=self._presentation_data,
        )

    def search_and_comment(
        self,
        text: str,
        text_to_highlight: str,
        locale: str = "en-US",
    ) -> "SlideEditor":
        """
        Looks for the text_to_highlight within the shapes' text and prints the details
        of its first occurrence, including which shapes it spans and the indexes within those shapes.
        """
        text_to_highlight = text_to_highlight.lower()
        shapes: list[ShapeEditor] = self.get_shapes()
        all_text = "".join([shape.data.text.lower() for shape in shapes if shape.data])
        all_text = all_text.replace("\n", " ")
        start_idx = all_text.find(text_to_highlight)
        if start_idx == -1:
            logging.error(f"SlideEditor: '{text_to_highlight}' not found in any shape.")
            return self

        end_idx = start_idx + len(text_to_highlight)

        cumulative_length = 0
        remaining_start = start_idx
        remaining_end = end_idx

        results = []

        for i, shape in enumerate(shapes):
            shape_text = shape.data.text
            shape_length = len(shape_text)

            shape_global_start = cumulative_length
            shape_global_end = cumulative_length + shape_length

            if shape_global_end > remaining_start and shape_global_start < remaining_end:
                local_start = max(0, remaining_start - shape_global_start)
                local_end = min(shape_length, remaining_end - shape_global_start)

                highlighted_substring = shape_text[local_start:local_end]

                results.append({
                    "shape_index": i,
                    "shape_text": shape_text,
                    "highlighted_substring": highlighted_substring,
                    "local_start": local_start,
                    "length": local_end - local_start
                })

            cumulative_length += shape_length

            if cumulative_length >= end_idx:
                break

        for r in results:
            curr_shape: ShapeEditor = shapes[r["shape_index"]]
            curr_shape.comment(text, r["local_start"], r["length"], locale)

        return self

    def get_shape(self, shape_id: str) -> ShapeEditor:
        return ShapeEditor(shape_id, self._data, self._archiver)

    def get_shapes(self) -> list[ShapeEditor]:
        shapes_ids = self._extract_shape_ids()
        return [self.get_shape(shape_id.get("id")) for shape_id in shapes_ids]

    def _extract_shape_ids(self) -> list[et.Element]:
        result = []
        slide = self._archiver.get_file(
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
                    shape_metadata = shape.find(
                        ".//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr"
                    )
                    if validate_element(shape_metadata):
                        shape_id = shape_metadata.find(
                            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"
                        )
                        if validate_element(shape_id):
                            result.append(shape_id)
        return result

    def _extract_comment_file_name(self) -> str | None:
        slide_rel = self._archiver.get_file(
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

    def _extract_slide_creation_id(self):
        slide = self._archiver.get_file(
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

    def _extract_slide_id(self, slide_index: int) -> str:
        presentation_file = self._archiver.get_file("ppt/presentation.xml")
        presentation_xml = et.fromstring(presentation_file.data)
        slides_list_root = presentation_xml.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst"
        )
        if validate_element(slides_list_root):
            slides_ids = slides_list_root.findall(
                ".//{http://schemas.openxmlformats.org/presentationml/2006/main}sldId"
            )
            if slides_ids and len(slides_ids) > slide_index:
                return slides_ids[slide_index].get("id")
        raise IndexError(f"Slide index {slide_index} is out of range.")

    @property
    def data(self):
        return self._data
