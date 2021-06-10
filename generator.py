from backend import *
import argparse


def argsParser():
    parser = argparse.ArgumentParser()
    parser.add_argument("week_no", help="Provide the week number to process", type=int)
    parser.add_argument(
        "-b",
        "--build_week",
        help="Create a YAML file with all the details of the given week",
        action="store_true",
    )
    parser.add_argument(
        "-p",
        "--get_pdf",
        help="Save the final weekly report as PDF",
        action="store_true",
    )
    parser.add_argument(
        "-d",
        "--get_docx",
        help="Save the final weekly report as PDF",
        action="store_true",
    )
    return parser.parse_args()


args = argsParser()
if not args.build_week and not (args.get_docx or args.get_pdf):
    cprint(
        "Please choose at least one among:\n\
    1. -b      : to build a new YAML with week details\n\
    2. -d | -p : to create docx to pdf of the weekly report",
        "magenta",
    )
if args.build_week:
    Week(args.week_no).getInfo()
if args.get_docx:
    Converter(args.week_no).createDocx(verbose=True)
if args.get_pdf:
    converter = Converter(args.week_no)
    converter.saveAsPDF()
    if not args.get_docx:
        converter.flushDocx()
