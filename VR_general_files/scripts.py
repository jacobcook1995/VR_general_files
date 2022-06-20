import VR_general_files.sharepoint as sp
import argparse
import textwrap


def _VR_file_search_cli():
    """
    This program maps the file structure of the Virtual Rainforest Project
    Sharepoint. It records both the structure and the descriptions attached to
    individual folders and files. EXPLAIN EXCEL READING.

    EXPLAIN THE DOC STRUCTURE.
    """

    desc = textwrap.dedent(_VR_file_search_cli.__doc__)
    fmt = argparse.RawDescriptionHelpFormatter
    parser = argparse.ArgumentParser(description=desc, formatter_class=fmt)

    parser.add_argument(
        "-c",
        "--credentials",
        default="private/secret.ini",
        type=str,
        help="Provide path to client crudentials to access the sharepoint",
        dest="cpath",
    )

    parser.add_argument(
        "-d",
        "--database",
        default="database/file_struct.db",
        type=str,
        help="Provide path to save file structure as a database",
        dest="out",
    )

    args = parser.parse_args()

    sp.scan_files(cpath=args.cpath, out=args.out)
