from typing import List
from Excel import Excel
from MyFunctions import debug, error, configurations
from Collector import Collector


class AlignChecker:

    def __init__(self, excel1, excel2) -> None:
        self.__excel1: Excel = None
        self.__excel2: Excel = None
        self.__collector: Collector = None

        self.__excel1 = excel1
        self.__excel2 = excel2
        self.__collector = Collector()

        assert (
            self.__excel1 is not None
            and self.__excel2 is not None
            and self.__collector is not None
        )
        pass

    def CheckAlignmentBetweenTheFiles(self) -> Collector:

        debug("Checking the 'Create on' field...")
        self.__CheckIfCreatedOnFieldIsValid(self.__excel1, "Excel1")
        self.__CheckIfCreatedOnFieldIsValid(self.__excel2, "Excel2")
        debug("'Create on' field checked successfully!")

        debug("Checking the 'Automatable' with 'Automated' field...")
        self.__CheckAutomatableWithAutomatedIfAreAligned(self.__excel1, "Excel1")
        self.__CheckAutomatableWithAutomatedIfAreAligned(self.__excel2, "Excel2")
        debug("'Automatable' and 'Automated' field checked successfully!")

        debug("Checking the 'Automated' with 'Tool' field...")
        self.__CheckAutomatedWithToolIfAreAligned(self.__excel1, "Excel1")
        self.__CheckAutomatedWithToolIfAreAligned(self.__excel2, "Excel2")
        debug("'Automated' and 'Tool' field checked successfully!")

        debug("Comparing all the records seperately...")
        self.__CompareTheFiles(self.__excel1, self.__excel2)
        debug("Comparison finished successfully!")

        return self.__collector

    def __CheckAutomatedWithToolIfAreAligned(self, excel: Excel, excel_name: str):
        for record in excel.records:
            automated, tool = record[4], record[6]
            if automated == "Yes" and tool != configurations["Automation_Tool"]:
                self.__collector.Add(
                    sheet="Automated-Tool",
                    record=record,
                    comment="Since the record is automated the Tool field must be "
                    + configurations["Automation_Tool"]
                    + " in "
                    + excel_name,
                )
                error(
                    "Record",
                    record,
                    "since it is automated the Tool field must be",
                    configurations["Automation_Tool"],
                    "in",
                    excel_name,
                )
        pass

    def __CheckAutomatableWithAutomatedIfAreAligned(
        self, excel: Excel, excel_name: str
    ):
        for record in excel.records:
            automatable, automated = record[3], record[4]

            if automatable == "No" and automated == "Yes":
                self.__collector.Add(
                    sheet="Automatable-Automated",
                    record=record,
                    comment="The record has invalid 'Automatable' and 'Automated' values in "
                    + excel_name,
                )
                error(
                    "Record",
                    record,
                    "has invalid 'Automatable' and 'Automated' values in",
                    excel_name,
                )

        pass

    def __CheckIfCreatedOnFieldIsValid(self, excel: Excel, excel_name: str):
        release: str = excel.records[0][5]

        for record in excel.records:
            release_of_record = record[5]
            if release_of_record > release:
                self.__collector.Add(
                    sheet="Create-on",
                    record=record,
                    comment="The 'Created on' value in "
                    + excel_name
                    + ", it should be "
                    + release,
                )
                error(
                    "Record",
                    record,
                    "has different 'Created on' value in",
                    excel_name,
                    ",it should be",
                    release,
                )
        pass

    def __CompareTheFiles(self, excel1: Excel, excel2: Excel) -> None:

        not_aligned_records = set()

        # Check if all records from excel1 are inside the excel2
        for record in excel1.records:
            id, name, lot2id, automatable, automated, created_on, tool = record
            found = False
            for record2 in excel2.records:

                (
                    id_2,
                    name_2,
                    lot2id_2,
                    automatable_2,
                    automated_2,
                    created_on_2,
                    tool_2,
                ) = record2

                if lot2id == lot2id_2 and name == name_2:
                    found = True
                    if (
                        automated != automated_2
                        or created_on != created_on_2
                        or automatable != automatable_2
                        or tool != tool_2
                    ):
                        not_aligned_records.add((record, record2))
            if not found:
                self.__collector.Add(
                    sheet="Not-Aligned-Records",
                    record=record,
                    comment="Didn't find this record inside Excel2",
                )
                error("Didn't find", record, "inside Excel2")

        # Check if all records from excel2 are inside the excel1
        for record2 in excel2.records:
            id_2, name_2, lot2id_2, automatable_2, automated_2, created_on_2, tool_2 = (
                record2
            )
            found = False
            for record in excel1.records:

                id, name, lot2id, automatable, automated, created_on, tool = record

                if lot2id == lot2id_2 and name == name_2:
                    found = True
                    if (
                        automated != automated_2
                        or created_on != created_on_2
                        or automatable != automatable_2
                        or tool != tool_2
                    ):
                        not_aligned_records.add((record, record2))
            if not found:
                self.__collector.Add(
                    sheet="Not-Aligned-Records",
                    record=record,
                    comment="Didn't find this record inside Excel1",
                )
                error("Didn't find", record, "inside Excel1")

        for record1, record2 in not_aligned_records:
            self.__collector.Add2(
                sheet="Not-Aligned-Records",
                record1=record1,
                record2=record2,
                comment="Records are not aligned",
            )
            error(record1)
            error(record2)
            error()

    pass
