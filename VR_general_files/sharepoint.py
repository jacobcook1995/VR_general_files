import re
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
import sqlite3
import configparser


def get_sharepoint_folder_contents(ctx, dir):
    """
    Function to provide a dictionary of folders and files within a
    sharepoint folder.
    """

    # Get the sub-directories
    subdirs = dir.folders
    ctx.load(subdirs)
    ctx.execute_query()

    # Get the files
    files = dir.files
    ctx.load(files)
    ctx.execute_query()

    return dict(
        folders=[(sub.properties["ServerRelativeUrl"], sub) for sub in subdirs],
        files=[(f.properties["Name"], f) for f in files],
    )


# Currently this function extracts the description of a file and then returns it
def expand_file_details(ctx, file_url):
    # Find relevant file
    file = (
        ctx.web.get_file_by_server_relative_url(file_url)
        .expand(["versions", "listItemAllFields/properties"])
        .get()
        .execute_query()
    )
    desc = file.listItemAllFields.get_property("Properties").get(
        "OData__x005f_ExtendedDescription"
    )
    # Replace blank strings with None
    if desc == "":
        desc = None

    return desc


# Identical function for the folders
def expand_folder_details(ctx, fold_url):
    # Find relevant folder
    fold = (
        ctx.web.get_folder_by_server_relative_url(fold_url)
        .expand(["versions", "listItemAllFields/properties"])
        .get()
        .execute_query()
    )
    desc = fold.list_item_all_fields.get_property("Properties").get(
        "OData__x005f_ExtendedDescription"
    )
    # Replace blank strings with None
    if desc == "":
        desc = None

    return desc


# Currently this function extracts the description of a file and then returns it
def expand_xlsx_details(ctx, file_url):
    # Find relevant (.xlsx) file
    file = (
        ctx.web.get_file_by_server_relative_url(file_url)
        .expand(["versions", "listItemAllFields/properties"])
        .get()
        .execute_query()
    )
    desc = file.listItemAllFields.get_property("Properties").get(
        "Excelcontents"
    )
    # Replace blank strings with None
    if desc == "":
        desc = None

    return desc


def scan_files(cpath: str, out: str):
    """Recursively scan files

    This function takes a configured document root directory on a sharepoint site
    and scans it recursively for a complete listing of files. The retrieved properties
    include the relative URL, which can be used to link directly to the file. The download
    url, using the config settings and a Row `f` are:

        url =  f"{tenant_name}/:t:/r/{f.relative_url}"

    Files can only be accessed via these URLs after a user has logged in to Sharepoint using
    college credentials and also has access rights to the file, so needs a Sharepoint folder
    with managed access for markers. That is relatively simple to do. There is also an API
    to provide shared links to anyone in the organisation. Users would still need to log in
    but the access management _within the organisation_ can be omitted. The API is needlessly
    obscure though, so not implemented, but the URLs look like.

        url = f"{tenant_name}/:t:/s/{site}/{cryptic_share_code}
    """

    # Generate database to store file structure and comments in
    db = sqlite3.connect(out)

    # Get the configured sharepoint tenant, site and relative url and client_id
    # and client_secret credentials for the application
    conf = configparser.ConfigParser()
    conf.read(cpath)

    tenant_name = conf["sharepoint"]["tenant_name"]
    site = conf["sharepoint"]["site"]
    root_dir_relative_url = conf["sharepoint"]["root_dir_relative_url"]

    # Use these to generate user credential
    client_credentials = ClientCredential(
        conf["sharepoint"]["client_id"],
        conf["sharepoint"]["client_secret"],
    )

    # Connect to sharepoint
    ctx = ClientContext(f"{tenant_name}/sites/{site}").with_credentials(
        client_credentials
    )

    # Get the root directory
    root = ctx.web.get_folder_by_server_relative_url(root_dir_relative_url)

    # Scan the directory for files, until this list is emptied.
    dir_filo = [("root", root)]

    # Now iterate over the directory contents collecting dictionaries of file data
    fold_data = []
    file_data = []
    fold_n = 0
    file_n = 0

    # counters for files, folders, excels sheets with descriptions
    miss_file_desc = 0
    miss_fold_desc = 0
    miss_xlsx_desc = 0

    while dir_filo:

        # Get the first entry from the FILO for directories and scan it
        this_dir = dir_filo.pop(0)
        contents = get_sharepoint_folder_contents(ctx, this_dir[1])

        if this_dir[0] == "root":
            # Generate root folder
            fold_data.append(
                dict(
                    unique_id=0,
                    parent_id=-1,
                    name="ROOT",
                    relative_url=f"/sites/{site}/{root_dir_relative_url}",
                    description=None,
                )
            )
            # Set parent ID as zero
            pID = 0

        else:
            # Find ID of parent folder
            par = next(
                item for item in fold_data if item["relative_url"] == this_dir[0]
            )
            pID = par["unique_id"]

        # Now generate child folders
        for fold in contents["folders"]:
            fold_n += 1
            fold_props = fold[1].properties
            desc = expand_folder_details(ctx, fold_props["ServerRelativeUrl"])
            if desc == None:
                miss_fold_desc += 1
            fold_data.append(
                dict(
                    unique_id=fold_n,
                    parent_id=pID,
                    name=fold_props["Name"],
                    relative_url=fold_props["ServerRelativeUrl"],
                    description=desc,
                )
            )

        # Contents is a dictionary of folders and files, so add the folders
        # onto the front of the directory FILO (depth first search)
        dir_filo = contents["folders"] + dir_filo

        # If there are any files, they are 2-tuples of (name, office365.sharepoint.files.file.File)
        # which can be used to retrieve key information
        for each_file in contents["files"]:

            file_n += 1
            file_props = each_file[1].properties
            desc = expand_file_details(ctx, file_props["ServerRelativeUrl"])
            if desc == None:
                miss_file_desc += 1
            # Check that xlsx files have sheet descriptions
            if file_props["Name"].endswith(".xlsx"):
                xlsx_shts = expand_xlsx_details(ctx, file_props["ServerRelativeUrl"])
                if xlsx_shts == None:
                    miss_xlsx_desc += 1
            else:
                xlsx_shts = None
            file_data.append(
                dict(
                    unique_id=file_n,
                    folder_id=pID,
                    name=file_props["Name"],
                    relative_url=file_props["ServerRelativeUrl"],
                    description=desc,
                    excel_sheets=xlsx_shts,
                )
            )

    # create database table for the folder structure
    cur = db.cursor()
    q = (
        "CREATE TABLE folders"
        "(folder_id int PRIMARY KEY, parent_id int, name text, rel_url text, "
        "description text)"
    )
    cur.execute(q)

    for fold in fold_data:
        q = (
            f"INSERT INTO folders VALUES ({fold['unique_id']}, {fold['parent_id']}, "
            f"'{fold['name']}', '{fold['relative_url']}', '{fold['description']}')"
        )
        cur.execute(q)

    q = (
        "CREATE TABLE files"
        "(file_id int PRIMARY KEY, folder_id int, name text, rel_url text, "
        "description text, excel_sheets text)"
    )
    cur.execute(q)

    for file in file_data:
        q = (
            f"INSERT INTO files VALUES ({file['unique_id']}, {file['folder_id']}, "
            f"'{file['name']}', '{file['relative_url']}', '{file['description']}',"
            f" '{file['excel_sheets']}')"
        )
        cur.execute(q)

    db.commit()

    # Give user a summary of missing file, folder, and Excel sheet descriptions
    print(f"Summary:")
    print(f"The directory has {miss_file_desc} files without a description.")
    print(f"It also has {miss_fold_desc} folders without a description.")
    print(f"Finally, {miss_xlsx_desc} xlsx files are missing sheet descriptions.")

    return
