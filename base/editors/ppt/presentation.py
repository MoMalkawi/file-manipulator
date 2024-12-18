from base.components.ppt.author import PPTAuthor
from base.data.components.ppt import AuthorData, PresentationData
from base.data.exceptions.ppt.presentation import SlideIndexOutOfRange
from base.editors.archive import AbstractArchiveEditor
from base.editors.ppt.slide import SlideEditor


class PresentationEditor(AbstractArchiveEditor):
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
        :param custom_author: If a custom author isn't specified a default one will be created.
        """
        super().__init__(presentation)
        self._data: PresentationData | None = self._load_data(custom_author=custom_author)


    def get_slide(self, slide_index: int):
        if self._data.slides_count <= slide_index:
            raise SlideIndexOutOfRange(slide_index)
        return SlideEditor(slide_index, self.data, self._archive)

    def add_author(self, custom_author: AuthorData | None = None) -> AuthorData:
        author_data = custom_author or AuthorData()

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

    def _load_slide_count(self):
        return len(self._archive.get_filenames_in_dir("ppt/slides"))

    def _load_data(self, **kwargs) -> PresentationData:
        return PresentationData(
            slides_count=self._load_slide_count(),
            author=self.add_author(custom_author=kwargs.get("custom_author"))
        )
