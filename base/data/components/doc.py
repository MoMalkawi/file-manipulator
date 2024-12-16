from datetime import datetime, timezone

from base.data.components import XMLFileData


class DocCommentData(XMLFileData):
    # w:comment
    initials: str = "DM"  # w:initials
    creation_date: str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")  # w:date
    comment_id: str = "0"  # w:id
    author: str = "Document Manipulator"  # w:author
    text: str  #  w:t

    def to_xml(self, skeleton):
        from base.data.namespaces.docx import DocUtils

        comment = DocUtils.create_element(
            "comment",
            parent=skeleton,
            w__initials=self.initials,
            w__date=self.creation_date,
            w__id=self.comment_id,
            w__author=self.author
        )

        p = DocUtils.create_element("p", parent=comment)

        r = DocUtils.create_element("r", parent=p)

        DocUtils.create_element("t", parent=r, text=self.text)

        return comment