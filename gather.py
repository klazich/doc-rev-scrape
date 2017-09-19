import json
import datetime
from os import walk, remove, makedirs
from os.path import *


TARGET_DIR = normpath('H:/Document Registration')
EXT_TARGETS = ['.docx', '.doc', '.xlsx', '.xls']


def gather_filenames(dir=TARGET_DIR, skip=[]):

    files = {}
    file_count =0


    for (dirpath, dirnames, filenames) in walk(dir):

        for filename in filenames:

            ext = splitext(filename)[1]
            filepath = join(dirpath, filename)

            if ext in EXT_TARGETS:
                files[filename] = { 'path': filepath, 'rev': None }
                file_count += 1

    to_json = {
        'data': {
            'count': file_count,
            'timestamp': datetime.datetime.now().isoformat()
        },
        'files': files
    }

    return json.dumps(to_json, indent=4)


