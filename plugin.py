#!/usr/bin/env python

import os
import re
import sys
import uuid
import shutil
import tempfile
import subprocess

# auxiliary KindleUnpack libraries for azw3/mobi splitting
from dualmetafix_mmap import DualMobiMetaFix, pathof
from mobi_split import mobi_split

# for metadata parsing
try:
    from sigil_bs4 import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup


# main plugin routine
def run(bk):
    ''' the main routine '''

    # get OEBPS path
    if bk.launcher_version() >= 20190927:
        OEBPS = bk.get_startingdir(bk.get_opfbookpath())
    else:
        OEBPS = os.path.join('OEBPS')

    # create temp folder
    temp_dir = tempfile.mkdtemp()

    # copy book files
    bk.copy_book_contents_to(temp_dir)

    # get content.opf
    if bk.launcher_version() >= 20190927:
        opf_path = os.path.join(temp_dir, bk.get_opfbookpath())
    else:
        opf_path = os.path.join(temp_dir, OEBPS, 'content.opf')

    # get epub version number
    if bk.launcher_version() >= 20160102:
        epubversion = bk.epub_version()
    else:
        opf_contents = bk.get_opf()
        epubversion = BeautifulSoup(opf_contents, 'lxml').find('package')['version']

    # get metadata soup
    metadata_soup = BeautifulSoup(bk.getmetadataxml(), 'lxml')

    #-----------------------------------------------
    # get required metadata items (title & language
    #-----------------------------------------------

    # get author
    if metadata_soup.find('dc:creator') and metadata_soup.find('dc:creator').string is not None:
        dc_creator = metadata_soup.find('dc:creator').string
        dc_creator = re.sub(r'[/|\?|<|>|\\\\|:|\*|\||"|\^| ]+', '_', dc_creator)
    else:
        dc_creator = ''

    # get title
    if metadata_soup.find('dc:title'):
        dc_title = metadata_soup.find('dc:title').string
    else:
        print('\nError: Missing title metadata!\n\nPlease click OK to close the Plugin Runner window.')
        return -1

    # get language
    if metadata_soup.find('dc:language'):
        pass
    else:
        print('\nError: Missing language metadata!\n\nPlease click OK to close the Plugin Runner window.')
        return -1

    #---------------
    # get/set asin
    #--------------
    asin = str(uuid.uuid4())[24:82]
    if epubversion.startswith("2"):
        dc_identifier = metadata_soup.find('dc:identifier', {'opf:scheme': re.compile('(MOBI-ASIN|AMAZON)', re.IGNORECASE)})
        if dc_identifier is not None:
            asin = dc_identifier.string
    else:
        dc_identifier = metadata_soup.find(string=re.compile("urn:(mobi-asin|amazon)", re.IGNORECASE))
        if dc_identifier is not None:
            asin = dc_identifier.split(':')[2]

    #=================
    # main routine
    #=================

    # get debug preference
    debug = False

    kg_path = shutil.which('kindlegen')

    #------------------------------------------
    # define kindlegen command line parameters
    #------------------------------------------

    # define temporary mobi file name
    mobi_path = os.path.join(temp_dir, OEBPS, 'sigil.mobi')
    args = [kg_path, opf_path]

    args.append('-dont_append_source')
    args.append('-verbose')
    args.append('-o')
    args.append('sigil.mobi')

    # run kindlegen
    print("Running KindleGen ... please wait")
    if debug:
        print('args:', args)
    subprocess.check_call(args)

    #--------------------------------------
    # define output directory and filenames
    #--------------------------------------

    # output directory
    home = os.path.expanduser('~')
    desktop = os.path.join(home, 'Desktop')
    dst_folder = desktop

    # make sure the output file name is safe
    title = re.sub(r'[/|\?|<|>|\\\\|:|\*|\||"|\^| ]+', '_', dc_title)

    # define file paths
    azw_path = os.path.join(dst_folder, title + '_' + asin + '.azw3')

    #=================================================================
    # generate kfx and/or split mobi file into azw3 and mobi7 parts
    #=================================================================

    # if kindlegen didn't fail, there should be a .mobi file
    assert os.path.isfile(mobi_path)

    #------------------------------------------------------------------------
    # add ASIN and set book type to EBOK using KevinH's dualmetafix_mmap.py
    #------------------------------------------------------------------------
    dmf = DualMobiMetaFix(mobi_path, asin)
    open(pathof(mobi_path + '.tmp'), 'wb').write(dmf.getresult())

    # if DualMobiMetaFix didn't fail, there should be a .temp file
    if os.path.isfile(str(mobi_path) + '.tmp'):
        print('\nASIN: ' + asin + ' added')

        # delete original file and rename temp file
        os.remove(str(mobi_path))
        os.rename(str(mobi_path) + '.tmp', str(mobi_path))
    else:
        print('\nASIN couldn\'t be added.')

    #----------------------
    # copy output files
    #----------------------
    mobisplit = mobi_split(pathof(mobi_path))

    if mobisplit.combo:
        outmobi8 = pathof(azw_path)
        open(outmobi8, 'wb').write(mobisplit.getResult8())
        print('AZW3 file copied to ' + azw_path)
    else:
        print('\nPlugin Error: Invalid mobi file format.')

    # delete temp folder
    shutil.rmtree(temp_dir)

    print('\nPlease click OK to close the Plugin Runner window.')

    return 0


def main():
    print('I reached main when I should not have\n')
    return -1


if __name__ == "__main__":
    sys.exit(main())
