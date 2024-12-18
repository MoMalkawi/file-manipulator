

class CommentTargetTextNotFound(Exception):

    def __init__(self, text_to_locate: str, text_body: str):
        self._text_to_locate = text_to_locate
        self._text_body = text_body
        super().__init__(self.__str__())

    def __str__(self):
        return (f"ShapeEditor: Couldn't add comment due to an issue locating"
                f" [{self._text_to_locate}] in [{self._text_body}]")

