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
from io import BytesIO
import logging
from PIL import Image
import cssutils

from tkinter import Tk, BOTH, StringVar, IntVar, BooleanVar, PhotoImage, messagebox, DISABLED
from tkinter.ttk import Frame, Button, Label, Entry, Checkbutton, Combobox
import tkinter.filedialog as tkinter_filedialog

# auxiliary KindleUnpack libraries for azw3/mobi splitting
from dualmetafix_mmap import DualMobiMetaFix, pathof, iswindows
from mobi_split import mobi_split

# for metadata parsing
try:
    from sigil_bs4 import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup

# auxiliary tools
from epub_utils import epub_zip_up_book_contents

# detect OS
isosx = sys.platform.startswith('darwin')
islinux = sys.platform.startswith('linux')


# display kindlegen file selection dialog
def GetFileName(title):
    ''' displays a file selection dialog box '''
    file_path = tkinter_filedialog.askopenfilename(title=title)
    return file_path


# display kindlegen file selection dialog
def GetDir(title):
    ''' displayes a directory selection dialog box '''
    folder = tkinter_filedialog.askdirectory(title=title)
    return folder


# display message box
def AskYesNo(title, message):
    ''' displays a confirmation message box '''
    root = Tk()
    root.withdraw()
    answer = messagebox.askquestion(title, message)
    return answer


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
    kg_path = None

    if iswindows:
        # C:\Users\<User>\AppData\Local\Amazon\Kindle Previewer\lib\kindlegen.exe
        default_windows_path = os.path.join(GetLocalAppData(), 'Amazon', 'Kindle Previewer', 'lib', 'kindlegen.exe')
        if os.path.isfile(default_windows_path):
            kg_path = default_windows_path
        # C:\Users\<User>\AppData\Local\Amazon\Kindle Previewer 3\lib\fc\bin\kindlegen.exe
        default_windows_path2 = os.path.join(GetLocalAppData(), 'Amazon', 'Kindle Previewer 3', 'lib', 'fc', 'bin', 'kindlegen.exe')
        if os.path.isfile(default_windows_path2):
            kg_path = default_windows_path2
        # C:\Users\<User>\AppData\Local\Amazon\Kindle Previewer 4\lib\fc\bin\kindlegen.exe [*** for future versions ***]
        default_windows_path3 = os.path.join(GetLocalAppData(), 'Amazon', 'Kindle Previewer 4', 'lib', 'fc', 'bin', 'kindlegen.exe')
        if os.path.isfile(default_windows_path3):
            kg_path = default_windows_path3

    if islinux:
        # /usr/local/bin/kindlegen
        default_linux_path = os.path.join('/usr', 'local', 'bin', 'kindlegen')
        if os.path.isfile(default_linux_path):
            kg_path = default_linux_path
        # ~/bin/kindlegen
        default_linux_path2 = os.path.join(expanduser('~'), 'bin', 'kindlegen')
        if os.path.isfile(default_linux_path2):
            kg_path = default_linux_path2

    if isosx:
        # /Applications/Kindle Previewer.app/Contents/MacOS/lib/kindlegen
        default_osx_path = os.path.join('/Applications', 'Kindle Previewer.app', 'Contents', 'MacOS', 'lib', 'kindlegen')
        if os.path.isfile(default_osx_path):
            kg_path = default_osx_path
        # /Applications/Kindle Previewer 3.app/Contents/MacOS/lib/kindlegen/fc/bin/kindlegen
        default_osx_path2 = os.path.join('/Applications', 'Kindle Previewer 3.app', 'Contents', 'MacOS', 'lib', 'fc', 'bin', 'kindlegen')
        if os.path.isfile(default_osx_path2):
            kg_path = default_osx_path2
        # /Applications/Kindle Previewer 4.app/Contents/MacOS/lib/kindlegen/fc/bin/kindlegen [*** for future versions ***]
        default_osx_path3 = os.path.join('/Applications', 'Kindle Previewer 4.app', 'Contents', 'MacOS', 'lib', 'fc', 'bin', 'kindlegen')
        if os.path.isfile(default_osx_path3):
            kg_path = default_osx_path3

    # display select file dialog box
    if not kg_path:
        kg_path = GetFileName('Select kindlegen executable')
        if not kg_path or not os.path.basename(kg_path).startswith('kindlegen'):
            kg_path = None
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


class Dialog(Frame):
    ''' the main GUI class '''
    global Cancel
    Cancel = True

    def __init__(self, parent, bk):
        # display the dialog box
        Frame.__init__(self, parent)

        self.parent = parent
        self.bk = bk
        self.initUI()

    def savevalues(self):
        global Cancel
        Cancel = False

        # save dialog box values in dictionary
        prefs = self.bk.getPrefs()
        prefs['donotaddsource'] = self.donotaddsource.get()
        prefs['compression'] = self.compression.get()
        prefs['verbose'] = self.verbose.get()
        prefs['western'] = self.western.get()
        prefs['gif'] = self.gif.get()
        prefs['locale'] = self.locale.get()
        prefs['add_asin'] = self.add_asin.get()
        prefs['azw3_only'] = self.azw3_only.get()
        prefs['mobi7'] = self.mobi7.get()
        prefs['kpf'] = self.kpf.get()
        prefs['kfx'] = self.kfx.get()
        prefs['thumbnail'] = self.thumbnail.get()
        prefs['mobi_dir'] = self.mobi_dir.get()
        prefs['thumbnail_height'] = self.thumbnail_height.get()
        self.bk.savePrefs(prefs)
        self.master.destroy()
        self.master.quit()

    def cancel(self):
        #self.master.destroy()
        self.master.quit()

    def getdir(self):
        mobi_dir = GetDir('Select output folder.')
        if mobi_dir != '':
            self.mobi_dir.set(mobi_dir)
        else:
            self.mobi_dir.set(GetDesktop())

    def initUI(self):
        prefs = self.bk.getPrefs()

        # look for kindlegen binary
        if 'kg_path' not in prefs:
            kg_path = findKindleGen()
            if kg_path and os.path.basename(kg_path).startswith('kindlegen'):
                prefs['kg_path'] = kg_path

        # define dialog box properties
        self.parent.title("KindleGen")
        self.pack(fill=BOTH, expand=1)

        # start reading location
        if srl_def:

            srlLabel = Label(self, text="SRL: " + srl_def)
        else:
            if 'check_srl' in prefs and not prefs['check_srl']:
                srlLabel = Label(self, foreground="blue", text="SRL: IGNORED")
            else:
                srlLabel = Label(self, foreground="red", text="SRL: NOT FOUND")
        srlLabel.place(x=10, y=10)

        # HTML TOC
        if toc_def:
            tocLabel = Label(self, text="TOC: " + toc_def)
        else:
            tocLabel = Label(self, foreground="red", text="TOC: NOT FOUND")
        tocLabel.place(x=10, y=30)

        # Cover
        if cover_def:
            coverLabel = Label(self, text="Cover: " + cover_def)
        else:
            coverLabel = Label(self, foreground="red", text="Cover: NOT FOUND")
        coverLabel.place(x=10, y=50)

        # ASIN
        if asin:
            asinLabel = Label(self, text="ASIN: " + asin)
        else:
            asinLabel = Label(self, foreground="red", text="ASIN: NOT FOUND")
        asinLabel.place(x=10, y=70)

        # CFF/Type 1 (Postscript) font warning
        if cff:
            fontLabel = Label(self, foreground="red", text="CFF/Type 1")
            fontLabel.place(x=140, y=70)

        # Don't add source check button
        self.donotaddsource = BooleanVar(None)
        if 'donotaddsource' in prefs:
            self.donotaddsource.set(prefs['donotaddsource'])
        donotaddsourceCheckbutton = Checkbutton(self, text="Don't add source files", variable=self.donotaddsource)
        donotaddsourceCheckbutton.place(x=10, y=90)

        # compression label
        options = ['0', '1', '2']
        compressionLabel = Label(self, text="Compression: ")
        compressionLabel.place(x=10, y=110)
        # compression list box
        self.compression = StringVar()
        compression = Combobox(self, textvariable=self.compression)
        compression['values'] = options
        if 'compression' in prefs:
            compression.current(int(prefs['compression']))
        else:
            compression.current(0)
        compression.place(x=100, y=110, width=40, height=18)

        # Verbose output check button
        self.verbose = BooleanVar(None)
        if 'verbose' in prefs:
            self.verbose.set(prefs['verbose'])
        verboseCheckbutton = Checkbutton(self, text="Verbose output", variable=self.verbose)
        verboseCheckbutton.place(x=10, y=130)

        # Western check button check button
        self.western = BooleanVar(None)
        if 'western' in prefs:
            self.western.set(prefs['western'])
        westernCheckbutton = Checkbutton(self, text="Western (Windows-1252)", variable=self.western)
        westernCheckbutton.place(x=10, y=150)

        # Gif to jpeg check button
        self.gif = BooleanVar(None)
        if 'gif' in prefs:
            self.gif.set(prefs['gif'])
        gifCheckbutton = Checkbutton(self, text="Convert JPEG to GIF", variable=self.gif)
        gifCheckbutton.place(x=10, y=170)

        # Select locale list box
        locales = ['en', 'de', 'fr', 'it', 'es', 'zh', 'ja', 'pt', 'ru']
        localeLabel = Label(self, text="Language: ")
        localeLabel.place(x=10, y=190)
        # locale list box
        self.locale = StringVar()
        locale = Combobox(self, textvariable=self.locale)
        locale['values'] = locales
        if 'locale' in prefs:
            index = [i for i, x in enumerate(locales) if x == prefs['locale']]
            locale.current(int(index[0]))
        else:
            locale.current(0)
        locale.place(x=100, y=190, width=40, height=18)

        # Add ASIN check button
        self.add_asin = BooleanVar(None)
        if 'add_asin' in prefs:
            self.add_asin.set(prefs['add_asin'])
        add_asinCheckbutton = Checkbutton(self, text="Add fake ASIN", variable=self.add_asin)
        add_asinCheckbutton.place(x=10, y=210)

        # Generate azw3 check button
        self.azw3_only = BooleanVar(None)
        if 'azw3_only' in prefs:
            self.azw3_only.set(prefs['azw3_only'])
        azw3_onlyCheckbutton = Checkbutton(self, text="Generate AZW3", variable=self.azw3_only)
        azw3_onlyCheckbutton.place(x=10, y=230)

        # Generate mobi7 check button
        self.mobi7 = BooleanVar(None)
        if 'mobi7' in prefs:
            self.mobi7.set(prefs['mobi7'])
        mobi7Checkbutton = Checkbutton(self, text="Generate Mobi7", variable=self.mobi7)
        mobi7Checkbutton.place(x=10, y=250)

        # Generate kpf check button
        self.kpf = BooleanVar(None)
        if 'kpf' in prefs:
            self.kpf.set(prefs['kpf'])
        kpfCheckbutton = Checkbutton(self, text="Generate KPF", variable=self.kpf)
        kpfCheckbutton.place(x=10, y=270)
        # there's no Kindle Previewer for Linux
        if islinux:
            self.kpf = BooleanVar(None)
            kpfCheckbutton.config(state=DISABLED)

        # Generate kfx check button
        self.kfx = BooleanVar(None)
        if 'kfx' in prefs:
            self.kfx.set(prefs['kfx'])
        kfxCheckbutton = Checkbutton(self, text="Generate KFX", variable=self.kfx)
        kfxCheckbutton.place(x=10, y=290)
        # there's no Kindle Previewer for Linux
        if islinux:
            self.kfx = BooleanVar(None)
            kfxCheckbutton.config(state=DISABLED)

        # Generate thumbnails check button
        self.thumbnail = BooleanVar(None)
        if 'thumbnail' in prefs and cover_def:
            self.thumbnail.set(prefs['thumbnail'])
        thumbnailCheckbutton = Checkbutton(self, text="Generate thumbnail", variable=self.thumbnail)
        thumbnailCheckbutton.place(x=10, y=310)

        # thumbnail width
        self.thumbnail_height = IntVar(None)
        if 'thumbnail_height' in prefs:
            self.thumbnail_height.set(prefs['thumbnail_height'])
        else:
            self.thumbnail_height.set(330)
        thumbnail_heightEntry = Entry(self, textvariable=self.thumbnail_height)
        thumbnail_heightEntry.place(x=155, y=310, width=30)
        # thumbnail width label
        thumbnail_heightLabel = Label(self, text="px")
        thumbnail_heightLabel.place(x=185, y=310)

        # disable cover thumbnail controls, if cover wasn't found
        if not cover_def:
            thumbnailCheckbutton.config(state=DISABLED)
            thumbnail_heightEntry.config(state=DISABLED)
            thumbnail_heightLabel.config(state=DISABLED)

        # output dir label
        mobi_dirLabel = Label(self, text="Output dir: ")
        mobi_dirLabel.place(x=10, y=330)
        # output dir text box
        self.mobi_dir = StringVar(None)
        if 'mobi_dir' in prefs:
            self.mobi_dir.set(prefs['mobi_dir'])
        else:
            self.mobi_dir.set(GetDesktop())
        mobi_dirEntry = Entry(self, textvariable=self.mobi_dir)
        mobi_dirEntry.place(x=80, y=330, width=90)
        browseButton = Button(self, text="...", command=self.getdir)
        browseButton.place(x=180, y=330, width=30, height=18)

        # OK and Cancel buttons
        cancelButton = Button(self, text="Cancel", command=self.cancel)
        cancelButton.place(x=130, y=360)
        okButton = Button(self, text="OK", command=self.savevalues)
        okButton.place(x=10, y=360)


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

    # macOS calibre-debug location: /Applications/calibre.app/Contents/console.app/Contents/MacOS/calibre-debug
    mac_calibre_debug_path = os.path.join('/Applications', 'calibre.app', 'Contents', 'console.app', 'Contents', 'MacOS', 'calibre-debug')

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

    #------------------------------
    # check minimum cover width
    #------------------------------
    img = None
    if cover_id and os.path.splitext(bk.id_to_href(cover_id))[1][1:].upper() != 'SVG':
        imgdata = bk.readfile(cover_id)
        try:
            img = Image.open(BytesIO(imgdata)).convert('L')
            width, height = img.size
            if width < 500:
                plugin_warnings += '\nWarning: The cover is too small: ' + str(width) + ' x ' + str(height)
            # check recommended dpi
            if prefs['check_dpi']:
                xdpi, ydpi = img.info['dpi']
                if (int(xdpi) or int(ydpi)) < 300:
                    plugin_warnings += '\nInfo: Amazon recommends 300 dpi cover images. Your image has: ' + str(int(xdpi)) + ' x ' + str(int(ydpi)) + ' dpi.'
        except Exception as ex:
            plugin_warnings += '\n*** PYTHON ERROR ***\nAn exception {0} occurred.\nArguments:\n{1!r}'.format(str(ex), ex.args)
            pass

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

    # set Tk parameters for dialog box
    root = Tk()
    root.geometry("240x400+300+300")
    if not isosx:
        icon_img = PhotoImage(file=os.path.join(bk._w.plugin_dir, bk._w.plugin_name, 'sigil.png'))
        root.tk.call('wm', 'iconphoto', root._w, icon_img)
    root.mainloop()

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

    if 'kg_path' in prefs and not Cancel:
        kg_path = prefs['kg_path']

        #----------------------------------
        # display confirmation message box
        #----------------------------------
        if plugin_warnings != '':
            plugin_warnings = '\n*************************************************************' + plugin_warnings + '\n*************************************************************\n'
            print(plugin_warnings)
            answer = AskYesNo('Ignore warnings?', 'Do you want to ignore these warnings?')
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

        if prefs['donotaddsource'] and not prefs['kfx']:
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
        if prefs['azw3_only'] or prefs['mobi7'] or (not prefs['azw3_only'] and not prefs['mobi7'] and not prefs['kfx']):

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
            kfx_path = os.path.join(dst_folder, title + '_' + asin + '.kfx')
        else:
            dst_path = os.path.join(dst_folder, title + '.mobi')
            azw_path = os.path.join(dst_folder, title + '.azw3')
            kfx_path = os.path.join(dst_folder, title + '.kfx')

        cleaned_epub_path = kfx_path.replace('.kfx', '_cleaned.epub')

        #=================================================================
        # generate kfx and/or split mobi file into azw3 and mobi7 parts
        #=================================================================

        # if kindlegen didn't fail, there should be a .mobi file
        if os.path.isfile(mobi_path):

            #--------------------------------------------------------------------
            # generate KFX file from Kindlegen-generated MOBI
            #--------------------------------------------------------------------
            if prefs['kfx']:
                if prefs['add_asin']:
                    if isosx:
                        args = [mac_calibre_debug_path, '-r', 'KFX Output', '--', '-a', asin, '-e', '-p', '0', mobi_path, kfx_path]
                    else:
                        args = ['calibre-debug', '-r', 'KFX Output', '--', '-a', asin, '-e', '-p', '0', mobi_path, kfx_path]
                else:
                    if isosx:
                        args = [mac_calibre_debug_path, '-r', 'KFX Output', '-e', '-p', '0', mobi_path, kfx_path]
                    else:
                        args = ['calibre-debug', '-r', 'KFX Output', '--', '-e', '-p', '0', mobi_path, kfx_path]

                # run Calibre & KFX Output plugin
                print("\nRunning Calibre KFX Output plugin [mobi mode] ... please wait\n")
                try:
                    if debug:
                        print('args:', args)
                    result = kgWrapper(*args)
                except FileNotFoundError:
                    print('calibre-debug not found!\nClick OK to close the window.')
                    return -1

                # print Calibre messages
                calibre_messages = result[0].decode('utf-8', 'ignore')
                print(calibre_messages)

                # get the ASIN number generated by the KFX Plugin
                KFX_ASIN_MATCH = re.search('ASIN=([^,]+),', calibre_messages)
                if KFX_ASIN_MATCH is not None:
                    KFX_ASIN = KFX_ASIN_MATCH.group(1)

                # if the KFX output plugin didn't fail, there should be a KFX file
                if os.path.isfile(kfx_path):
                    if not prefs['add_asin']:
                        kfx_with_asin = kfx_path.replace('.kfx', '_' + KFX_ASIN + '.kfx')
                        os.rename(kfx_path, kfx_with_asin)
                        print('KFX file copied to ' + kfx_with_asin)
                    else:
                        print('KFX file copied to ' + kfx_path)
                else:
                    print('KFX file coudn\'t be generated.')

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
                if not prefs['kfx']:
                    shutil.copyfile(mobi_path, dst_path)
                    print('\nMobi file copied to ' + dst_path)

        else:
            if prefs['azw3_only'] or prefs['mobi7'] or (not prefs['azw3_only'] and not prefs['mobi7'] and not prefs['kfx']):
                #-----------------------------------
                # display KindleGen error messages
                #-----------------------------------
                if re.search(r'E\d+:', kg_messages):
                    print('\n*************************************************************\nkindlegen errors:\n*************************************************************')
                    kg_errors = kg_messages.splitlines()
                    for line in kg_errors:
                        if re.search(r'E\d+:', line):
                            print(line)

            else:
                #===========================
                # generate KFX from epub
                #===========================

                # create mimetype file
                with open(os.path.join(temp_dir, "mimetype"), "w") as mimetype:
                    mimetype.write("application/epub+zip")

                # zip up the epub folder and save the epub in the plugin folder
                epub_path = os.path.join(bk._w.plugin_dir, bk._w.plugin_name, 'temp.epub')
                if os.path.isfile(epub_path):
                    os.remove(str(epub_path))
                epub_zip_up_book_contents(temp_dir, epub_path)

                # assemble command line parameters
                if prefs['add_asin']:
                    if isosx:
                        args = [mac_calibre_debug_path, '-r', 'KFX Output', '--', '-a', asin, '-e', '-p', '0', epub_path, kfx_path]
                    else:
                        args = ['calibre-debug', '-r', 'KFX Output', '--', '-a', asin, '-e', '-p', '0', epub_path, kfx_path]
                else:
                    if isosx:
                        args = [mac_calibre_debug_path, '-r', 'KFX Output', '--', '-e', '-p', '0', epub_path, kfx_path]
                    else:
                        args = ['calibre-debug', '-r', 'KFX Output', '--', '-e', '-p', '0', epub_path, kfx_path]

                # save cleaned file
                if prefs['save_cleaned_file']:
                    args.insert(4, '-c')

                # run Calibre & KFX Output plugin
                print("\nRunning Calibre KFX Output plugin [epub mode]... please wait\n")
                try:
                    if debug:
                        print('args:', args)
                    result = kgWrapper(*args)
                except FileNotFoundError:
                    print('calibre-debug not found!\nClick OK to close the window.')
                    return -1

                # print Calibre messages
                calibre_messages = result[0].decode('utf-8', 'ignore')
                print(calibre_messages)

                # get the ASIN number generated by the KFX Plugin
                KFX_ASIN_MATCH = re.search('ASIN=([^,]+),', calibre_messages)
                if KFX_ASIN_MATCH is not None:
                    KFX_ASIN = KFX_ASIN_MATCH.group(1)

                # if the plugin didn't fail there should be a kfx file
                if os.path.isfile(kfx_path):
                    if not prefs['add_asin']:
                        kfx_with_asin = kfx_path.replace('.kfx', '_' + KFX_ASIN + '.kfx')
                        os.rename(kfx_path, kfx_with_asin)
                        print('KFX file copied to ' + kfx_with_asin)
                    else:
                        print('KFX file copied to ' + kfx_path)
                else:
                    print('KFX file coudn\'t be generated.')

                # if the plugin didn't fail there should be a cleaned epub file
                if prefs['save_cleaned_file']:
                    if os.path.isfile(cleaned_epub_path):
                        print('Cleaned epub file copied to ' + cleaned_epub_path)
                    else:
                        print('Cleaned epub coudn\'t be generated.')

                # delete temp epub from plugin folder
                os.remove(str(epub_path))

        # delete temp folder
        shutil.rmtree(temp_dir)

        #================================================
        # generate mobi thumbnail image for eInk kindles
        #================================================
        if img and prefs['thumbnail']:
            thumbnail_height = prefs['thumbnail_height']
            img.thumbnail((thumbnail_height, thumbnail_height), Image.ANTIALIAS)
            if prefs['add_asin']:
                img_dest_path = os.path.join(dst_folder, 'thumbnail_' + asin + '_EBOK_portrait.jpg')
            else:
                if prefs['kfx'] and KFX_ASIN is not None:
                    img_dest_path = os.path.join(dst_folder, 'thumbnail_' + KFX_ASIN + '_EBOK_portrait.jpg')
                else:
                    img_dest_path = os.path.join(dst_folder, 'thumbnail_EBOK_portrait.jpg')
            img.save(img_dest_path)

            if os.path.isfile(img_dest_path):
                print('\nThumbnail copied to: ' + img_dest_path)
            else:
                print('\nPlugin Error: Thumbnail creation failed.')

    else:
        if Cancel:
            print('\nPlugin terminated by user.')
        else:
            print('\nPlugin Warning: Kindlegen path not selected or invalid.\nPlease re-run the plugin and select the Kindlegen binary.')

    print('\nPlease click OK to close the Plugin Runner window.')

    return 0


def main():
    print('I reached main when I should not have\n')
    return -1


if __name__ == "__main__":
    sys.exit(main())
