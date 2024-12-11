import xml.etree.ElementTree as et

from base.editors.archive import SelectiveArchiveEditor
from base.models.parsers import AuthorData
from base.parsers.ppt.author import PPTAuthor
from base.utils import validate_element


class PresentationEditor:
    """
    PresentationEditor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Uses the ArchiveEditor to edit a PowerPoint Presentation file.
    """

    def __init__(self, presentation: str | bytes):
        """
        :param presentation: The presentation file path or bytes to edit.
        """
        self._archive: SelectiveArchiveEditor = SelectiveArchiveEditor(presentation)
        self._author: AuthorData = self._load_author()

    def get_slide(self, slide_index: int):
        from base.editors.ppt.slide import SlideEditor

        return SlideEditor(self, slide_index, self.get_slide_id(slide_index))

    def get_slide_count(self):
        return len(self._archive.get_filenames_in_dir("ppt/slides"))

    def get_slide_id(self, slide_index: int) -> str:
        presentation_file = self._archive.get_file("ppt/presentation.xml")
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

    def _load_author(self) -> AuthorData:
        author = PPTAuthor(
            "A123B456-C789-0D12-E345-F67890123456",
            self._archive.get_file("ppt/authors.xml"),
        )
        if not author.data:
            author.inject(
                self._archive,
                AuthorData(
                    id="A123B456-C789-0D12-E345-F67890123456",
                    name="Document Checker",
                    initials="DC",
                    user_id="S::documentchecker@noemail.com::12345678-90ab-cdef-1234-567890abcdef",
                    provider_id="AD",
                ),
            )
        return author.data

    @property
    def author_id(self):
        return self._author.id

    @property
    def archiver(self):
        return self._archive

    def export(self, path: str | None = None) -> bytes:
        return self._archive.export(path)

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):  # noqa
        self._archive.close()
