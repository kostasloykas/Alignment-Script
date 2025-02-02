from typing import Dict, List


class Collector:

    # SHEETS THAT WILL BE CREATED
    # CREATE_ON THAT HAS INFORMATION FOR EXCEL1 AND EXCEL2
    # AUTOMATABLE-AUTOMATED THAT HAS INFORMATION FOR EXCEL1 AND EXCEL2
    # AUTOMATED-TOOL THAT HAS INFORMATION FOR EXCEL1 AND EXCEL2
    # NOT ALIGNED RECORDS

    def __init__(self) -> None:
        self.book: Dict[str, List] = {}
        pass

    def Add(self, sheet: str, record: tuple, comment: str = None):
        self.__CreateTheSheetIfHaventAlreadyCreated(sheet)
        self.book[sheet].append(record + (comment,))
        pass

    def Add2(self, sheet: str, record1: List, record2: List, comment=None):
        self.__CreateTheSheetIfHaventAlreadyCreated(sheet)
        self.book[sheet].append((record1, record2) + (comment,))
        pass

    def __CreateTheSheetIfHaventAlreadyCreated(self, sheet: str):
        if sheet not in self.book.keys():
            self.book[sheet] = []
        pass

    pass
