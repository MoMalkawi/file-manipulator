from base.editors.archive import SelectiveArchiveEditor
from base.models.parsers import PresentationData
from base.parsers.ppt.author import AuthorData, PPTAuthor
from base.editors.ppt.slide import SlideEditor


class PresentationEditor:
    """
    PresentationEditor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Uses the ArchiveEditor to edit a PowerPoint Presentation file.
    """

    def __init__(self, presentation: str | bytes, custom_author: AuthorData | None = None):
        """
        :param presentation: The presentation file path or bytes to edit.
        """
        self._archive: SelectiveArchiveEditor = SelectiveArchiveEditor(presentation)
        self._data: PresentationData = self._load_data(custom_author=custom_author)

    def _load_data(self, custom_author: AuthorData | None = None) -> PresentationData:
        return PresentationData(
            slides_count=self.load_slide_count(),
            author=self._load_author(custom_author=custom_author),
        )

    def get_slide(self, slide_index: int):
        return SlideEditor(slide_index, self.data, self._archive)

    def load_slide_count(self):
        return len(self._archive.get_filenames_in_dir("ppt/slides"))

    def _load_author(self, custom_author: AuthorData | None = None) -> AuthorData:
        author_data = custom_author or AuthorData(
            id="A123B456-C789-0D12-E345-F67890123456",
            name="Document Checker",
            initials="DC",
            user_id="S::documentchecker@noemail.com::12345678-90ab-cdef-1234-567890abcdef",
            provider_id="AD",
        )

        author = PPTAuthor(
            author_data.id,
            self._archive.get_file("ppt/authors.xml"),
        )
        if not author.data:
            author.inject(
                self._archive,
                author_data,
            )
        return author.data

    def export(self, path: str | None = None) -> bytes:
        return self._archive.export(path)

    @property
    def data(self):
        return self._data

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):  # noqa
        self._archive.close()
