#!/usr/bin/env python

import sys
import uuid
import re
import shutil
import tempfile
import os
import os.path
from os.path import expanduser
from subprocess import Popen, PIPE
import logging
import cssutils

# auxiliary KindleUnpack libraries for azw3/mobi splitting
from dualmetafix_mmap import DualMobiMetaFix, pathof
from mobi_split import mobi_split

# for metadata parsing
try:
    from sigil_bs4 import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup

# detect OS
isosx = sys.platform.startswith('darwin')
islinux = sys.platform.startswith('linux')


# get 'C:\Users\<User>\AppData\Local\' folder location
def GetLocalAppData():
    ''' returns the local AppData folder '''
    # check for Windows 7 or higher
    if sys.getwindowsversion().major >= 6:
        return os.path.join(os.getenv('USERPROFILE'), 'AppData', 'Local')
    else:
        return os.path.join(os.getenv('USERPROFILE'), 'Local Settings', 'Application Data')


def GetDesktop():
    ''' returns the desktop/home folder '''
    # output directory
    home = expanduser('~')
    desktop = os.path.join(home, 'Desktop')
    if os.path.isdir(desktop):
        return desktop
    else:
        return home


# find kindlegen binary
def findKindleGen():
    ''' returns the KindleGen path '''
    kg_path = shutil.which('kindlegen')
    return kg_path


# simple kindlegen wrapper
def kgWrapper(*args):
    '''simple KindleGen wrapper '''
    process = Popen(list(args), stdout=PIPE, stderr=PIPE)
    ret = process.communicate()
    return ret


# reverts first last
def LastFirst(author):
    ''' reverses the name of the author '''
    author_parts = author.split(' ')
    if len(author_parts) >= 2:
        LastFirst = author_parts[len(author_parts) - 1] + ', ' + " ".join(author_parts[0:len(author_parts) - 1])
        return LastFirst
    else:
        return author


# main plugin routine
def run(bk):
    ''' the main routine '''
    # the epub temp folder
    ebook_root = bk._w.ebook_root

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

    plugin_warnings = ''
    global toc_def
    toc_def = None
    global srl_def
    srl_def = None
    global cover_def
    cover_def = None
    global asin
    asin = None
    global cff
    cff = False

    # get prefs
    prefs = bk.getPrefs()

    # write new optional settings to prefs file
    if 'check_srl' not in prefs:
        prefs['check_srl'] = True
        bk.savePrefs(prefs)
    if 'save_cleaned_file' not in prefs:
        prefs['save_cleaned_file'] = False
        bk.savePrefs(prefs)
    if 'check_dpi' not in prefs:
        prefs['check_dpi'] = False
        bk.savePrefs(prefs)

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
        dc_creator = LastFirst(metadata_soup.find('dc:creator').string)
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

    #========================================
    # check guides/landmarks cover metadata
    #=======================================

    #----------------------
    # get epub3 metadata
    #----------------------

    if epubversion.startswith("3"):

        # look for nav.xhtml
        opf_soup = BeautifulSoup(bk.get_opf(), 'lxml')
        nav_item = opf_soup.find('item', {'properties': 'nav'})
        if nav_item:
            nav_href = nav_item['href']
            nav_id = bk.href_to_id(nav_href)

            # get landmarks from nav document
            landmarks = {}
            nav_soup = BeautifulSoup(bk.readfile(nav_id), 'html.parser')
            nav_landmarks = nav_soup.find('nav', {'epub:type': 'landmarks'})
            if nav_landmarks is not None:
                for landmark in nav_landmarks.find_all('a', {'epub:type': re.compile('.*?')}):
                    epub_type = landmark['epub:type']
                    if 'href' in landmark.attrs:
                        href = landmark['href']
                        landmarks[epub_type] = href

            # check for required landmarks items
            if 'toc' not in landmarks:
                plugin_warnings += '\nWarning: Missing TOC epub3 landmark: Use Add Semantics > Table of Contents to mark the TOC.'
            else:

                toc_def = os.path.basename(landmarks['toc'])
            if 'bodymatter' not in landmarks:
                if prefs['check_srl']:
                    plugin_warnings += '\nWarning: Missing SRL epub3 landmark: Use Add Semantics > Bodymatter to mark the SRL.'
            else:

                srl_def = os.path.basename(landmarks['bodymatter'])
        else:
            plugin_warnings += '\nError: nav document not found!'

        # look for cover image
        cover_id = None
        cover_item = opf_soup.find('item', {'properties': 'cover-image'})
        if cover_item:
            cover_href = cover_item['href']
            cover_id = bk.href_to_id(cover_href)
            cover_def = os.path.basename(cover_href)
        else:
            plugin_warnings += '\nWarning: Cover not specified (cover-image property missing).'

            # look for cover page image references
            if 'cover' in landmarks:
                base_name = os.path.basename(landmarks['cover'].replace('../', '').split('#')[0])
                cover_page_id = bk.basename_to_id(base_name)
                cover_page_html = bk.readfile(cover_page_id)
                cover_image_href = re.search(r'(href|src)="(\.\.\/Images\/[^"]+)"', cover_page_html)
                if cover_image_href:
                    cover_id = bk.href_to_id(cover_image_href.group(2).replace('../', ''))
                    plugin_warnings += '\n"' + os.path.basename(cover_image_href.group(2)) + '" should have a cover-image property.\n'

    #======================
    # get epub2 metadata
    #======================
    else:
        # get guide items
        opf_guide_items = {}
        for ref_type, title, href in bk.getguide():
            opf_guide_items[ref_type] = href

        #-----------------------------------
        # check for required guide items
        #----------------------------------
        if 'toc' not in opf_guide_items:
            plugin_warnings += '\nWarning: Missing TOC guide item. Use Add Semantics > Table of Contents to mark the TOC.'
        else:

            toc_def = os.path.basename(opf_guide_items['toc'])

        if 'text' not in opf_guide_items:
            if prefs['check_srl']:
                plugin_warnings += '\nWarning: Missing SRL guide item. Use Add Semantics > Text to mark the SRL.'
        else:

            srl_def = os.path.basename(opf_guide_items['text'])

        #----------------------------
        # check for cover image
        #----------------------------
        cover_id = None
        cover_image = metadata_soup.find('meta', {'name': 'cover'})
        if cover_image:
            cover_id = cover_image['content']
            try:
                cover_href = bk.id_to_href(cover_id)
                cover_def = os.path.basename(cover_href)
            except Exception:
                plugin_warnings += '\nWarning: Unmanifested cover id: ' + cover_id
        else:
            plugin_warnings += '\nWarning: Cover not specified (cover metadata missing).'

            # look for cover page image references
            if 'cover' in opf_guide_items:
                base_name = os.path.basename(opf_guide_items['cover'].replace('../', '').split('#')[0])
                cover_page_id = bk.basename_to_id(base_name)
                cover_page_html = bk.readfile(cover_page_id)
                cover_image_href = re.search(r'(href|src)="(\.\.\/Images\/[^"]+)"', cover_page_html)
                if cover_image_href:
                    cover_id = bk.href_to_id(cover_image_href.group(2).replace('../', ''))
                    plugin_warnings += '\n<meta name="cover" content="' + cover_id + '" />'

    #-------------------------------------------------------------------------------------------------------------
    # check for Type 1 CFF fonts; for more information see https://blog.typekit.com/2005/10/06/phasing_out_typ/
    #-------------------------------------------------------------------------------------------------------------

    # get all fonts
    font_manifest_items = []
    font_manifest_items = list(bk.font_iter())

    # fonts aren't found by bk.font_iter() if they have the wrong mime type, e.g. 'application/octet-stream
    if font_manifest_items == []:
        for manifest_id, href, mime in bk.manifest_iter():
            if href.endswith('tf'):
                font_manifest_items.append((manifest_id, href, mime))

    for manifest_id, href, mime in font_manifest_items:
        if bk.launcher_version() >= 20190927:
            font_path = os.path.join(ebook_root, bk.id_to_bookpath(manifest_id))
        else:
            font_path = os.path.join(ebook_root, OEBPS, href)
        with open(font_path, 'rb') as f:
            # CFF/Type 1 fonts start with b'OTTO'
            magic_number = f.read(4)
            if magic_number == b'OTTO':
                plugin_warnings += '\nWarning: CFF/Type1 (Postscript) font found: ' + os.path.basename(href)
                cff = True

    #-------------------------------------
    # check for unsupported CSS selectors
    #-------------------------------------
    for css_id, href in bk.css_iter():
        css = bk.readfile(css_id)
        cssutils.log.setLevel(logging.FATAL)
        sheet = cssutils.parseString(css)
        for rule in sheet:
            if rule.type == rule.STYLE_RULE:
                # find properties that start with max:
                for property in rule.style:
                    if property.name.startswith('max'):
                        #plugin_warnings += '\nWarning: Unsupported CSS property "{}" found.\n'.format(property.name) + rule.cssText + '\n'
                        #print('Warning: Unsupported CSS property "{}" found.'.format(property.name))
                        pass

    #=================
    # main routine
    #=================

    # output directory
    home = expanduser('~')
    desktop = os.path.join(home, 'Desktop')
    if os.path.isdir(prefs['mobi_dir']):
        dst_folder = prefs['mobi_dir']
    else:
        if os.path.isdir(desktop):
            dst_folder = desktop
        else:
            dst_folder = home
        prefs['mobi_dir'] = dst_folder

        bk.savePrefs(prefs)

    # get debug preference
    debug = prefs.get('debug', False)

    if 'kg_path' in prefs:
        kg_path = prefs['kg_path']

        #----------------------------------
        # display confirmation message box
        #----------------------------------
        if plugin_warnings != '':
            plugin_warnings = '\n*************************************************************' + plugin_warnings + '\n*************************************************************\n'
            print(plugin_warnings)
            answer = input('Do you want to ignore warnings? (yes/no)')
            if answer == 'no':
                print('\nPlugin terminated by user.\n\nPlease click OK to close the Plugin Runner window.')
                return -1

        #------------------------------------------
        # define kindlegen command line parameters
        #------------------------------------------

        # define temporary mobi file name
        mobi_path = os.path.join(temp_dir, OEBPS, 'sigil.mobi')
        args = [kg_path, opf_path]

        if prefs['compression'] in ['0', '1', '2'] or [0, 1, 2]:
            args.append('-c' + str(prefs['compression']))

        if prefs['donotaddsource']:
            args.append('-dont_append_source')

        if prefs['verbose']:
            args.append('-verbose')

        if prefs['western']:
            args.append('-western')

        if prefs['gif']:
            args.append('-gif')

        if prefs['locale'] in ['en', 'de', 'fr', 'it', 'es', 'zh', 'ja', 'pt', 'ru']:
            args.append('-locale')
            args.append(prefs['locale'])

        args.append('-o')
        args.append('sigil.mobi')

        # only run kindlegen to generate mobi, mobi7 and mobi8 files
        if prefs['azw3_only'] or prefs['mobi7'] or (not prefs['azw3_only'] and not prefs['mobi7']):

            # run kindlegen
            print("Running KindleGen ... please wait")
            if debug:
                print('args:', args)
            result = kgWrapper(*args)

            # print kindlegen messages
            kg_messages = result[0].decode('utf-8', 'ignore').replace(temp_dir, '').replace('\r\n', '\n')
            print(kg_messages)

            # display kindlegen warning messages
            if re.search(r'W\d+:', kg_messages):
                print('\n*************************************************************\nkindlegen warnings:\n*************************************************************')
                kg_warnings = kg_messages.splitlines()
                for line in kg_warnings:
                    if re.search(r'W\d+:', line):
                        print(line)

        #--------------------------------------
        # define output directory and filenames
        #--------------------------------------

        # output directory
        home = expanduser('~')
        desktop = os.path.join(home, 'Desktop')
        if os.path.isdir(prefs['mobi_dir']):
            dst_folder = prefs['mobi_dir']
        else:
            if os.path.isdir(desktop):
                dst_folder = desktop
            else:
                dst_folder = home
            prefs['mobi_dir'] = dst_folder

            bk.savePrefs(prefs)

        # make sure the output file name is safe
        if dc_creator != '':
            title = dc_creator + '-' + re.sub(r'[/|\?|<|>|\\\\|:|\*|\||"|\^| ]+', '_', dc_title)
        else:
            title = re.sub(r'[/|\?|<|>|\\\\|:|\*|\||"|\^| ]+', '_', dc_title)

        # define file paths
        if prefs['add_asin']:
            dst_path = os.path.join(dst_folder, title + '_' + asin + '.mobi')
            azw_path = os.path.join(dst_folder, title + '_' + asin + '.azw3')
        else:
            dst_path = os.path.join(dst_folder, title + '.mobi')
            azw_path = os.path.join(dst_folder, title + '.azw3')

        #=================================================================
        # generate kfx and/or split mobi file into azw3 and mobi7 parts
        #=================================================================

        # if kindlegen didn't fail, there should be a .mobi file
        if os.path.isfile(mobi_path):

            #------------------------------------------------------------------------
            # add ASIN and set book type to EBOK using KevinH's dualmetafix_mmap.py
            #------------------------------------------------------------------------
            if prefs['azw3_only'] or prefs['mobi7']:
                if prefs['add_asin']:
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
            if prefs['azw3_only'] or prefs['mobi7']:
                mobisplit = mobi_split(pathof(mobi_path))

                if mobisplit.combo:
                    outmobi8 = pathof(azw_path)
                    outmobi7 = outmobi8.replace('.azw3', '.mobi')
                    if prefs['azw3_only']:
                        open(outmobi8, 'wb').write(mobisplit.getResult8())
                        print('AZW3 file copied to ' + azw_path)
                    if prefs['mobi7']:
                        open(outmobi7, 'wb').write(mobisplit.getResult7())
                        print('MOBI7 file copied to ' + azw_path.replace('.azw3', '.mobi'))
                else:
                    print('\nPlugin Error: Invalid mobi file format.')

            else:
                shutil.copyfile(mobi_path, dst_path)
                print('\nMobi file copied to ' + dst_path)

        else:
            if prefs['azw3_only'] or prefs['mobi7'] or (not prefs['azw3_only'] and not prefs['mobi7']):
                #-----------------------------------
                # display KindleGen error messages
                #-----------------------------------
                if re.search(r'E\d+:', kg_messages):
                    print('\n*************************************************************\nkindlegen errors:\n*************************************************************')
                    kg_errors = kg_messages.splitlines()
                    for line in kg_errors:
                        if re.search(r'E\d+:', line):
                            print(line)

        # delete temp folder
        shutil.rmtree(temp_dir)

    else:
        print('\nPlugin Warning: Kindlegen path not selected or invalid.\nPlease re-run the plugin and select the Kindlegen binary.')

    print('\nPlease click OK to close the Plugin Runner window.')

    return 0


def main():
    print('I reached main when I should not have\n')
    return -1


if __name__ == "__main__":
    sys.exit(main())
