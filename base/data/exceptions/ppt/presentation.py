

class SlideIndexOutOfRange(Exception):

    def __init__(self, requested_index: int):
        self._requested_index = requested_index
        super().__init__(self.__str__())

    def __str__(self):
        return f"Presentation: Slide {self._requested_index} is out of bounds."