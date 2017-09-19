import re
from shutil import copy2, rmtree
from random import randrange
from os import walk, remove, makedirs, rename, chdir
from os.path import normpath, realpath, join, exists, commonprefix, \
    splitext, basename

from uuid import uuid4
from helper import Com, do_excel, do_docx, do_doc, do_pdf


chdir('C:/Users/uskla/Desktop/doc-rev-check')
# RX_REV = re.compile(r'\b(revision|rev|r)[. ]?.*?(?!(?:release|date:?|'
#                      'review)).*?(\d+)', re.IGNORECASE)
RX_REV = re.compile(r'(rev|rev\.|revision|r)(?:\s+)?(\d+)', re.IGNORECASE)


def get_rev(text):
    found = RX_REV.search(text)
    if found:
        return found.group(0)
    return None


def in_directory(file, directory):
    directory = join(realpath(directory), '')
    file = realpath(file)
    return commonprefix([file, directory]) == directory


def rand20():
    return randrange(1, 20)


PARSEABLE = ['.pdf', '.PDF',
             '.docx', '.DOCX',
             '.doc', '.DOC',
             '.xlsx', '.XLSX',
             '.xls', '.XLS',
             '.xlsm', '.XLSM']
COUNTS = {'doc':  {'success': 0,
                   'no_rev_found': 0,
                   'error': 0},
          'docx': {'success': 0,
                   'no_rev_found': 0,
                   'error': 0},
          'xls':  {'success': 0,
                   'no_rev_found': 0,
                   'error': 0},
          'xlsx': {'success': 0,
                   'no_rev_found': 0,
                   'error': 0},
          'pdf':  {'success': 0,
                   'no_rev_found': 0,
                   'error': 0}}
TARGET_DIRS = [normpath('H:/Document Registration/03 General Procedures'),
               normpath('H:/Document Registration/Specifications/'
                        'Quality (SPQ)/General (SPQ-GEN)'),
               normpath('H:/Document Registration/Specifications/'
                        'Warehouse (SPW)/Logistics (SPW-LOG)'),
               normpath('H:/Document Registration/04 Department Procedures/'
                        'Quality (DPQ)')]
CHECK_DIR = normpath('C:/Users/uskla/Desktop/doc-rev-check/check')

if exists('check'):
    new_name = str(uuid4())
    rename('check', new_name)
    rmtree(new_name, ignore_errors=True)

makedirs(CHECK_DIR)
check = rand20()
to_check = []
i = 1

com = Com()

remove('out.txt')

with open("out.txt", "a") as out:

    for D in TARGET_DIRS:

        for (dirpath, dirnames, filenames) in walk(D):

            if len(filenames) == 0:
                continue

            # does this directory have files we can parse?
            isParseable = False
            for fn in filenames:
                ext = splitext(fn)[1]
                if ext in PARSEABLE:
                    isParseable = True
                    break

            # Begin looping through files in this directory
            for fn in filenames:

                ext = splitext(fn)[1]                   # get file type
                file = normpath(dirpath + '\\' + fn)   # concat base to dir

                if fn.startswith('~') or ext not in PARSEABLE:
                    continue                            # skip open file types

                if i == check:
                    copy2(file, CHECK_DIR)
                    to_check.append(file)
                    mark = '*'
                    i = 0
                    check = rand20()
                else:
                    mark = ' '
                    i += 1

                print('   checking... {} {}'.format(mark, fn), end='\r')

                if ext in ['.pdf', '.PDF']:

                    text = do_pdf(file)

                    rev = get_rev(text)
                    if rev is None:
                        rev = 'NO REV FOUND'
                        COUNTS['pdf']['no_rev_found'] += 1

                    print('{0:>14}'.format(rev))
                    out.write('{0:>25}  PDF  {1} {2}\n'.format(rev, mark, fn))
                    COUNTS['pdf']['success'] += 1

                elif ext in ['.docx', '.DOCX']:

                    text = do_docx(file)

                    if text != 1:
                        rev = get_rev(text)
                        if rev is None:
                            rev = 'NO REV FOUND'
                            COUNTS['docx']['no_rev_found'] += 1
                    else:
                        rev = 'ERROR'
                        COUNTS['docx']['error'] += 1
                    print('{0:>14}'.format(rev))
                    out.write('{0:>25}  DOCX {1} {2}\n'.format(rev, mark, fn))
                    COUNTS['docx']['success'] += 1

                elif ext in ['.doc', '.DOC']:

                    text = do_doc(file, com)

                    if text != 1:
                        rev = get_rev(text)
                        if rev is None:
                            rev = 'NO REV FOUND'
                            COUNTS['doc']['no_rev_found'] += 1
                    else:
                        rev = 'ERROR'
                        COUNTS['doc']['error'] += 1
                    print('{0:>14}'.format(rev))
                    out.write('{0:>25}  DOC  {1} {2}\n'.format(rev, mark, fn))
                    COUNTS['doc']['success'] += 1

                elif ext in ['.xlsx', '.XLSX', '.xlsm', '.XLSM']:

                    text = do_excel(file, com)

                    if text != 1:
                        rev = get_rev(text)
                        if rev is None:
                            rev = 'NO REV FOUND'
                            COUNTS['xlsx']['no_rev_found'] += 1
                    else:
                        rev = 'ERROR'
                        COUNTS['xlsx']['error'] += 1
                    print('{0:>14}'.format(rev))
                    out.write('{0:>25}  XLSX {1} {2}\n'.format(rev, mark, fn))
                    COUNTS['xlsx']['success'] += 1

                elif ext in ['.xls', '.XLS']:

                    text = do_excel(file, com)

                    if text != 1:
                        rev = get_rev(text)
                        if rev is None:
                            rev = 'NO REV FOUND'
                            COUNTS['xls']['no_rev_found'] += 1
                    else:
                        rev = 'ERROR'
                        COUNTS['xls']['error'] += 1
                    print('{0:>14}'.format(rev))
                    out.write('{0:>25}  XLS  {1} {2}\n'.format(rev, mark, fn))
                    COUNTS['xls']['success'] += 1

    com.done()

    print()
    out.write('\n')
    total_parsed = sum(sum(int(x) for x in y.values())
                       for y in COUNTS.values())
    print('total parsed --> ', total_parsed)
    out.write('total parsed --> ' + str(total_parsed) + '\n')
    for k in COUNTS.keys():
        print('               {}'.format(k.upper()))
        print('      SUCCESS: {}'.format(str(COUNTS[k]['success'])))
        print(' NO REV FOUND: {}'.format(str(COUNTS[k]['no_rev_found'])))
        print('        ERROR: {}'.format(str(COUNTS[k]['error'])))
        out.write('               {}\n'.format(k.upper()))
        out.write('      SUCCESS: {}\n'.format(str(COUNTS[k]['success'])))
        out.write(' NO REV FOUND: {}\n'.format(str(COUNTS[k]['no_rev_found'])))
        out.write('        ERROR: {}\n'.format(str(COUNTS[k]['error'])))

    print()
    out.write('\n')
    print('Files copied to "check" directory...')
    out.write('Files copied to "check" directory...\n')
    for filepath in to_check:
        print('    {}'.format(basename(filepath)))
        out.write('    {}\n'.format(basename(filepath)))
