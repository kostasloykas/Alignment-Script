import argparse
from io import BufferedReader
from typing import List
import warnings
import AlignChecker
from Collector import Collector
from Excel import Excel
from AlignChecker import AlignChecker
from Exporter import Exporter


def main():

    [file1, file2] = TakeParameters()
    print("Arguments Parsed")

    file1 = Excel(file1)
    print("Data from Excel1 Loaded")

    file2 = Excel(file2)
    print("Data from Excel2 Loaded")

    checker = AlignChecker(file1, file2)
    collector = checker.CheckAlignmentBetweenTheFiles()
    print("The alignment of the data have been successfully checked")

    exporter = Exporter(collector)
    exporter.ExportTheExcelFile()
    print("The excel file has been successfully exported ")

    print("Process finished!")
    pass


def TakeParameters() -> List[BufferedReader]:
    file1 = None
    file2 = None

    # Define the expected input
    parser = argparse.ArgumentParser(
        description="Compare two Excel files in order to check if 2 Releases are aligned with each other. \
            It is mandatory that the excel files must have these fields -> (Id,Name,Lot2Id,Automatable,Automated,CreatedOn,Tool)"
    )
    parser.add_argument(
        "-f1",
        type=argparse.FileType("rb"),
        help="Path to the first Excel file.",
        required=True,
    )
    parser.add_argument(
        "-f2",
        type=argparse.FileType("rb"),
        help="Path to the second Excel file.",
        required=True,
    )

    # Parse arguments
    arguments = parser.parse_args()

    file1: BufferedReader = arguments.f1
    file2: BufferedReader = arguments.f2

    assert file1 != None and file2 != None
    return [file1, file2]


if __name__ == "__main__":
    main()
