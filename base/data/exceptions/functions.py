


class FunctionArgumentMissing(Exception):

    def __init__(self, function_name: str, *missing_fields_names: str):
        self._function_name: str = function_name
        self._missing_fields: str = ", ".join(missing_fields_names)
        super().__init__(self.__str__())

    def __str__(self):
        return f"{self._function_name}: Interrupted due to missing or incorrect argument [{self._missing_fields}]"
