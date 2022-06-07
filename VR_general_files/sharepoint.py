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

    return dict(folders = [(sub.properties['ServerRelativeUrl'], sub) for sub in subdirs],
                files = [(f.properties['Name'], f) for f in files])

def get_folder_or_file_description(ctx, item):
    """
    Function to find the description (and Excel contents) of a specific sharepoint
    file or folder.
    """

    # Get the full details
    details = item.list_item_all_fields
    ctx.load(details)
    ctx.execute_query()

    return details

def scan_files(cpath: str):
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
    db = sqlite3.connect('database/file_struct.db')

    # Get the configured sharepoint tenant, site and relative url
    # and credentials for a college role user that has been given access
    # to that relative URL.
    conf = configparser.ConfigParser()
    conf.read('private/appconfig_template.ini')

    tenant_name = conf['sharepoint']['tenant_name']
    site = conf['sharepoint']['site']
    root_dir_relative_url = conf['sharepoint']['root_dir_relative_url']

    # Then read in credentials from secret config
    p_conf = configparser.ConfigParser()

    p_conf.read(cpath)

    # Use these to generate user credential
    client_credentials = ClientCredential(p_conf['private-sharepoint']['client_id'],
                                          p_conf['private-sharepoint']['client_secret'])

    # Connect to sharepoint
    ctx = ClientContext(f"{tenant_name}/sites/{site}").with_credentials(client_credentials)

    # Get the root directory
    root = ctx.web.get_folder_by_server_relative_url(root_dir_relative_url)

    # Scan the directory for files, until this list is emptied.
    dir_filo = [('root', root)]

    # Now iterate over the directory contents collecting dictionaries of file data
    fold_data = []
    file_data = []
    fold_n = 0

    while dir_filo:

        # Get the first entry from the FILO for directories and scan it
        this_dir = dir_filo.pop(0)
        contents = get_sharepoint_folder_contents(ctx, this_dir[1])
        details = get_folder_or_file_description(ctx, this_dir[1])
        # FROM DETAILS I CAN FIND AND PRINT THE PROPERTIES OF THE FOLDER/FILE
        print(details.properties)

        # BASICALLY NEED TO GIVE FOLDER A UID, A NAME, THE UID OF IT'S PARENT, + ANY COMMENTS, IGNORE EXCEL DATA
        # EVERYTHING IS EASY BAR THE PARENT ID
        if this_dir[0] == 'root':
            # Generate root folder
            fold_data.append(dict(unique_id = 0,
                                  parent_id = -1,
                                  name = "ROOT",
                                  comments = None))
            # Then generate all child folders
            for fold in contents['folders']:
                fold_n += 1
                fold_props = fold[1].properties
                # WORK OUT HOW TO FIND COMMENTS
                fold_data.append(dict(unique_id = fold_n,
                                      parent_id = 0,
                                      name = fold_props['Name'],
                                      comments = None))
        # else:
        #     # Generate child folders




        # Contents is a dictionary of folders and files, so add the folders
        # onto the front of the directory FILO (depth first search)
        dir_filo = contents['folders'] + dir_filo

        # If there are any files, they are 2-tuples of (name, office365.sharepoint.files.file.File)
        # which can be used to retrieve key information
        for each_file in contents['files']:

            file_props = each_file[1].properties

            # RETURN HERE TO SHORTEN EXECUCTION WHILE DEVELOPING THE ABOVE
            return

            # OKAY SO THE BELOW IS STUFF WRITEN BY DAVID THAT IS POTENTIALLY USEFUL
            # BUT HAS TO BE MODIFIED TO THIS USE CASE

            # Can't see how to filter to only PDFs using the sharepoint API, so
            # do it here.
            if not file_props['Name'].endswith('.pdf'):
                continue

            # Get the student CID
            cid = cid_regex.search(file_props['Name'])
            if cid is not None:
                cid = int(cid.group())

            # The files are expected to be structured within root_dir_relative_url as:
            #   Presentation/Year/Role/File.pdf
            # because the file url is always relative to the account root, need to trim down to the
            # final 3 directories of the path name

            path = file_props['ServerRelativeUrl'].split('/')

            # OKAY SO DAVID IS SAVING DETAILS OF REPORTS INTO FILE DATA HERE
            # IS A LIST THE BEST OPTION IN MY CASE?
            # FOR FILES ALMOST CERTAINLY, CAN SPECIFY THE FOLDER
            # I GUESS I WOULD HAVE TO CONSTRUCT A SEPERATE STRUCTURE FOR FOLDER DATA
            # MAYBE GIVE EACH FOLDER A UID, AS NAMES ARE NOT GUARRENTEED TO BE UNIQUE
            # FOR THE SAME REASON FILES NEED UIDS
            file_data.append(dict(unique_id = file_props['UniqueId'],
                                  filename = file_props['Name'],
                                  filesize = file_props['Length'],
                                  cid = cid,
                                  presentation = path[-4],
                                  academic_year = path[-3],
                                  marker_role = path[-2],
                                  relative_url = file_props['ServerRelativeUrl']))
