from io import BytesIO
from zipfile import ZipFile

from base.parsers.file import ArchiveFile


class SelectiveArchiveEditor:
    """
    SelectiveArchiveEditor

    Author: Mohammad Malkawi
    Email: mohammad.m.malkawi@pwc.com
    --------------------------------------------
    Loads files from an archive, Allows for Modification, then saves the archive as bytes for export.
    Also allows for addition of files.

    Example Usage:
    editor = SelectiveArchiveEditor("path/to/archive-able-file")
    editor
    .upsert("ppts/comments/7r.xml", ArchiveFile(name="...", data="..."))
    .export() -> This closes the editor and returns the bytes of the modified archive for export.
    """

    def __init__(self, file: str | bytes):
        """
        :param file: The path to the archive file or the bytes of the archive file
        """
        self._zip_file: ZipFile = self._load(file)
        self._closed: bool = False
        self._modifications: dict[str, str] = {}  # file_path: new_content

    @classmethod
    def _load(cls, file: str | bytes) -> ZipFile:
        """reads the zip file and keeps it open for appending"""
        if isinstance(file, bytes):
            buffer = BytesIO(file)
        else:
            with open(file, "rb") as file:
                buffer = BytesIO(file.read())
        return ZipFile(buffer, "a")

    def get_filenames_in_dir(self, dir_path: str):
        dir_path = dir_path.rstrip("/") + "/"
        return [
            name
            for name in self._zip_file.namelist()
            if name.startswith(dir_path) and "/" not in name[len(dir_path) :]
        ]

    def get_files(self, *file_names: str) -> list[ArchiveFile]:
        results = []
        for file_name in file_names:
            try:
                if file_name in self._modifications:
                    results.append(
                        ArchiveFile(
                            file_name=file_name, data=self._modifications[file_name]
                        )
                    )
                with self._zip_file.open(file_name) as file:
                    results.append(
                        ArchiveFile(file_name=file_name, data=file.read().decode())
                    )
            except KeyError:
                results.append(None)
        return results

    def get_file(self, file_name: str) -> ArchiveFile | None:
        try:
            if file_name in self._modifications:
                return ArchiveFile(
                    file_name=file_name, data=self._modifications[file_name]
                )
            with self._zip_file.open(file_name) as file:
                return ArchiveFile(file_name=file_name, data=file.read().decode())
        except KeyError:
            return None

    def upsert(self, archive_file_path: str, file_modifier: callable):
        """modify or add a file within the archive."""
        try:
            with self._zip_file.open(archive_file_path) as file_to_modify:
                current_content = file_to_modify.read().decode()
        except KeyError:
            current_content = ""

        new_content = file_modifier(
            ArchiveFile(file_name=archive_file_path, data=current_content)
        )

        # tree = et.ElementTree(et.fromstring(new_content.data))
        # formatted_xml = et.tostring(tree.getroot(), encoding="unicode")
        self._modifications[archive_file_path] = new_content.data
        return self

    def upsert_many(self, archive_files_paths: list[str], files_modifier):
        files_contents = []
        try:
            for file_path in archive_files_paths:
                with self._zip_file.open(file_path) as file_to_modify:
                    files_contents.append(file_to_modify.read().decode())
        except KeyError:
            files_contents.append("")

        new_content = []
        if isinstance(files_modifier, list):
            for i, modifier in enumerate(files_modifier):
                new_content.append(
                    modifier(
                        ArchiveFile(
                            file_name=archive_files_paths[i], data=files_contents[i]
                        )
                    )
                )
        else:
            new_content = files_modifier(
                [
                    ArchiveFile(data=content, file_name=archive_files_paths[i])
                    for i, content in enumerate(files_contents)
                ]
            )

        for content in new_content:
            self._modifications[content.name] = content.data

        return self

    def export(self, export_path: str | None) -> bytes:
        """return the zip file as bytes"""
        buffer = BytesIO()
        with ZipFile(buffer, "w") as export_zip:
            for path, data in self._modifications.items():
                export_zip.writestr(path, data)

            for name in self._zip_file.namelist():
                if name not in self._modifications:
                    with self._zip_file.open(name) as file:
                        export_zip.writestr(name, file.read())

        zip_file_bytes = buffer.getvalue()
        if export_path:
            with open(export_path, "wb") as f:
                f.write(zip_file_bytes)
        return zip_file_bytes

    def close(self):
        if not self._closed:
            self._zip_file.close()
            self._modifications.clear()
            self._closed = True

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):  # noqa
        self.close()
