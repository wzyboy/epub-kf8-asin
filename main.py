#!/usr/bin/env python

import os
import re
import uuid
import shutil
import zipfile
import logging
import tempfile
import argparse
import subprocess

from bs4 import BeautifulSoup

from dualmetafix_mmap import DualMobiMetaFix
from mobi_split import mobi_split


logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(os.path.basename(__file__))


class EPUB:

    def __init__(self, filename):
        with zipfile.ZipFile(filename, 'r') as zip:
            container_xml = zip.read('META-INF/container.xml').decode()
            container_soup = BeautifulSoup(container_xml, 'xml')
            opf_path = container_soup.find('rootfile').get('full-path')
            opf_xml = zip.read(opf_path).decode()
        self.opf_soup = BeautifulSoup(opf_xml, 'xml')

    @property
    def author(self):
        return self.opf_soup.find('dc:creator').string

    @property
    def title(self):
        return self.opf_soup.find('dc:title').string

    @property
    def language(self):
        return self.opf_soup.find('dc:language').string

    @property
    def identifier(self):
        return self.opf_soup.find('dc:identifier').string

    @property
    def version(self):
        return self.opf_soup.find('package')['version']


def convert(epub_path, kf8_path=None, asin=None, quiet=False):

    logger.info(f'Processing: {epub_path}')
    epub = EPUB(epub_path)

    # ASIN
    if not asin:
        if epub.identifier:
            logger.info(f'Detected book identifier: {epub.identifier}')
            # Identifier looks like a normative UUID.
            if epub.identifier.startswith('urn:uuid:'):
                asin = epub.identifier.split(':')[2]
            # Identifier looks like a bare UUID.
            elif re.match(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', epub.identifier.lower()):
                asin = epub.identifier
            # Identifier looks like a genuine ASIN
            elif re.match(r'^B[0-9A-Z]{9}$', epub.identifier):
                asin = epub.identifier
            else:
                logger.info('NOT using detected book identifier as ASIN')

    if asin:
        logger.info(f'Using UUID/ASIN: {asin}')
    else:
        # Generate fake ASIN
        asin = str(uuid.uuid4())
        logger.info(f'Generated a fake ASIN: {asin}')

    # Make a temp copy of the book
    temp_dir = tempfile.mkdtemp()
    epub_tmp = os.path.join(temp_dir, f'{asin}.epub')
    shutil.copy(epub_path, epub_tmp)

    # Generate temp .mobi file
    mobi_tmp = os.path.join(temp_dir, f'{asin}.mobi')
    kindlegen_cmd = ['kindlegen', epub_tmp, '-dont_append_source']
    logger.info(f'Running: {kindlegen_cmd}')
    if not quiet:
        subprocess.call(kindlegen_cmd)
    else:
        subprocess.call(
            kindlegen_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    assert os.path.isfile(mobi_tmp)

    # Fix metadata of temp .mobi file:
    # - Add ASIN
    # - Set type to EBOK
    logger.info('Fixing metadata ...')
    dmf = DualMobiMetaFix(mobi_tmp, asin)
    with open(mobi_tmp, 'wb') as f:
        f.write(dmf.getresult())

    # KF8 Output
    if not kf8_path:
        epub_dir = os.path.dirname(epub_path)
        clean_title = re.sub(r'[/|\?|<|>|\\\\|:|\*|\||"|\^| ]+', '_', epub.title)
        kf8_path = os.path.join(epub_dir, f'{asin}_{clean_title}.azw3')

    # Extract KF8 from temp .mobi file
    logger.info('Extracting KF8 ...')
    mobisplit = mobi_split(mobi_tmp)
    with open(kf8_path, 'wb') as f:
        f.write(mobisplit.getResult8())

    # Clean up
    logger.info('Cleaning up ...')
    shutil.rmtree(temp_dir)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('epub_path')
    ap.add_argument('-o', '--output', help='KF8 (.azw3) output')
    ap.add_argument('-a', '--asin', help='provide/override ASIN')
    ap.add_argument('-q', '--quiet', action='store_true', help='suppress kindlegen output')
    args = ap.parse_args()

    convert(args.epub_path, args.output, args.asin, args.quiet)


if __name__ == "__main__":
    main()
