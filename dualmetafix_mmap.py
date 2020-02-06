#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fileencoding=UTF-8:ts=4:sw=4:sta:et:sts=4:ai

# Extracted from KindleUnpack code
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, version 3.
# Copyright © 2013 P. Durrant, K. Hendricks, S. Siebert, fandrieu, DiapDealer, nickredding.

from __future__ import unicode_literals, division, absolute_import, print_function

import sys
import os
import getopt
import struct
import locale
import codecs
import traceback

PY2 = sys.version_info[0] == 2
PY3 = sys.version_info[0] == 3

if PY3:
    text_type = str
    binary_type = bytes
else:
    range = xrange
    text_type = unicode
    binary_type = str

iswindows = sys.platform.startswith('win')

# Because Windows (and Mac OS X) allows full unicode filenames and paths
# any paths in pure bytestring python 2.X code must be utf-8 encoded as they will need to
# be converted on the fly to full unicode for Windows platforms.  Any other 8-bit str 
# encoding would lose characters that can not be represented in that encoding

# these are simple support routines to allow use of utf-8 encoded bytestrings as paths in main program
# to be converted on the fly to full unicode as temporary un-named values to prevent
# the potential mixing of unicode and bytestring string values in the main program 

fsencoding = sys.getfilesystemencoding()

def pathof(s, enc=fsencoding):
    if s is None:
        return None
    if isinstance(s, text_type):
        return s
    if isinstance(s, binary_type):
        try:
            return s.decode(enc)
        except:
            pass
    return s

def unicode_argv():
    global iswindows
    global PY3
    if PY3:
        return sys.argv
    if iswindows:
        # Versions 2.x of Python don't support Unicode in sys.argv on
        # Windows, with the underlying Windows API instead replacing multi-byte
        # characters with '?'.  So use shell32.GetCommandLineArgvW to get sys.argv
        # as a list of Unicode strings
        from ctypes import POINTER, byref, cdll, c_int, windll
        from ctypes.wintypes import LPCWSTR, LPWSTR

        GetCommandLineW = cdll.kernel32.GetCommandLineW
        GetCommandLineW.argtypes = []
        GetCommandLineW.restype = LPCWSTR

        CommandLineToArgvW = windll.shell32.CommandLineToArgvW
        CommandLineToArgvW.argtypes = [LPCWSTR, POINTER(c_int)]
        CommandLineToArgvW.restype = POINTER(LPWSTR)

        cmd = GetCommandLineW()
        argc = c_int(0)
        argv = CommandLineToArgvW(cmd, byref(argc))
        if argc.value > 0:
            # Remove Python executable and commands if present
            start = argc.value - len(sys.argv)
            return [argv[i] for i in
                    range(start, argc.value)]
        # this should never happen
        return None
    else:
        argv = []
        argvencoding = sys.stdin.encoding
        if argvencoding is None:
            argvencoding = sys.getfilesystemencoding()
        if argvencoding is None:
            argvencoding = 'utf-8'
        for arg in sys.argv:
            if isinstance(arg, text_type):
                argv.append(arg)
            else:
                argv.append(arg.decode(argvencoding))
        return argv


class DualMetaFixException(Exception):
    pass

# palm database offset constants
number_of_pdb_records = 76
first_pdb_record = 78

# important rec0 offsets
mobi_header_base = 16
mobi_header_length = 20
mobi_version = 36
title_offset = 84

def getint(data,ofs,sz=b'L'):
    i, = struct.unpack_from(b'>'+sz,data,ofs)
    return i

def writeint(data,ofs,n,len=b'L'):
    if len==b'L':
        return data[:ofs]+struct.pack(b'>L',n)+data[ofs+4:]
    else:
        return data[:ofs]+struct.pack(b'>H',n)+data[ofs+2:]

def getsecaddr(datain,secno):
    nsec = getint(datain,number_of_pdb_records,b'H')
    if (secno < 0) | (secno >= nsec):
        emsg = 'requested section number %d out of range (nsec=%d)' % (secno,nsec)
        raise DualMetaFixException(emsg)
    secstart = getint(datain,first_pdb_record+secno*8)
    if secno == nsec-1:
        secend = len(datain)
    else:
        secend = getint(datain,first_pdb_record+(secno+1)*8)
    return secstart,secend

def readsection(datain,secno):
    secstart, secend = getsecaddr(datain,secno)
    return datain[secstart:secend]

# overwrite section - must be exact same length as original
def replacesection(datain,secno,secdata): 
    secstart,secend = getsecaddr(datain,secno)
    seclen = secend - secstart
    if len(secdata) != seclen:
        raise DualMetaFixException('section length change in replacesection')
    datalst = []
    datalst.append(datain[0:secstart])
    datalst.append(secdata)
    datalst.append(datain[secend:])
    dataout = b"".join(datalst)
    return dataout

def get_exth_params(rec0):
    ebase = mobi_header_base + getint(rec0,mobi_header_length)
    if rec0[ebase:ebase+4] != b'EXTH':
        raise DualMetaFixException('EXTH tag not found where expected')
    elen = getint(rec0,ebase+4)
    enum = getint(rec0,ebase+8)
    rlen = len(rec0)
    return ebase,elen,enum,rlen

def add_exth(rec0,exth_num,exth_bytes):
    ebase,elen,enum,rlen = get_exth_params(rec0)
    newrecsize = 8+len(exth_bytes)
    newrec0 = rec0[0:ebase+4]+struct.pack(b'>L',elen+newrecsize)+struct.pack(b'>L',enum+1)+\
              struct.pack(b'>L',exth_num)+struct.pack(b'>L',newrecsize)+exth_bytes+rec0[ebase+12:]
    newrec0 = writeint(newrec0,title_offset,getint(newrec0,title_offset)+newrecsize)
    # keep constant record length by removing newrecsize null bytes from end
    sectail = newrec0[-newrecsize:]
    if sectail != b'\0'*newrecsize:
        raise DualMetaFixException('add_exth: trimmed non-null bytes at end of section')
    newrec0 = newrec0[0:rlen]
    return newrec0

def read_exth(rec0,exth_num):
    exth_values = []
    ebase,elen,enum,rlen = get_exth_params(rec0)
    ebase = ebase+12
    while enum>0:
        exth_id = getint(rec0,ebase)
        if exth_id == exth_num:
            # We might have multiple exths, so build a list.
            exth_values.append(rec0[ebase+8:ebase+getint(rec0,ebase+4)])
        enum = enum-1
        ebase = ebase+getint(rec0,ebase+4)
    return exth_values

def del_exth(rec0,exth_num):
    ebase,elen,enum,rlen = get_exth_params(rec0)
    ebase_idx = ebase+12
    enum_idx = 0
    while enum_idx < enum:
        exth_id = getint(rec0,ebase_idx)
        exth_size = getint(rec0,ebase_idx+4)
        if exth_id == exth_num:
            newrec0 = rec0
            newrec0 = writeint(newrec0,title_offset,getint(newrec0,title_offset)-exth_size)
            newrec0 = newrec0[:ebase_idx]+newrec0[ebase_idx+exth_size:]
            newrec0 = newrec0[0:ebase+4]+struct.pack(b'>L',elen-exth_size)+struct.pack(b'>L',enum-1)+newrec0[ebase+12:]
            newrec0 = newrec0 + b'\0'*(exth_size)
            if rlen != len(newrec0):
                raise DualMetaFixException('del_exth: incorrect section size change')
            return newrec0
        enum_idx += 1
        ebase_idx = ebase_idx+exth_size
    return rec0


class DualMobiMetaFix:

    def __init__(self, infile, asin):
        self.datain = open(pathof(infile), 'rb').read()
        self.datain_rec0 = readsection(self.datain,0)
        self.asin = asin.encode('utf-8')
        # in the first mobi header
        # add 501 to "EBOK", add 113 as asin, add 504 as asin
        rec0 = self.datain_rec0
        rec0 = del_exth(rec0, 501)
        rec0 = del_exth(rec0, 113)
        rec0 = del_exth(rec0, 504)
        rec0 = add_exth(rec0, 113, self.asin)
        rec0 = add_exth(rec0, 504, self.asin)
        rec0 = add_exth(rec0, 501, b'EBOK')
        self.datain = replacesection(self.datain, 0, rec0)

        ver = getint(self.datain_rec0,mobi_version)
        self.combo = (ver!=8)
        if not self.combo:
            return

        exth121 = read_exth(self.datain_rec0,121)
        if len(exth121) == 0:
            self.combo = False
            return
        else:
            # only pay attention to first exth121
            # (there should only be one)
            datain_kf8, = struct.unpack_from(b'>L',exth121[0],0)
            if datain_kf8 == 0xffffffff:
                self.combo = False
                return
        self.datain_kfrec0 =readsection(self.datain,datain_kf8)

        # in the second header
        # add 501 to "EBOK", add 113 as asin, add 504 as asin
        rec0 = self.datain_kfrec0
        rec0 = del_exth(rec0, 501)
        rec0 = del_exth(rec0, 113)
        rec0 = del_exth(rec0, 504)
        rec0 = add_exth(rec0, 504, self.asin)
        rec0 = add_exth(rec0, 113, self.asin)
        rec0 = add_exth(rec0, 501, b'EBOK')
        self.datain = replacesection(self.datain, datain_kf8, rec0)

    def getresult(self):
        return self.datain


def usage(progname):
    print("")
    print("Description:")
    print("   Simple Program to add EBOK and ASIN Info to Meta on Dual Mobis")
    print("  ")
    print("Usage:")
    print("  %s -h asin infile.mobi outfile.mobi" % progname)
    print("  ")
    print("Options:")
    print("    -h           print this help message")
    print("   ")
    print(" Extracted from KindleUnpack code")
    print(" This program is free software: you can redistribute it and/or modify")
    print(" it under the terms of the GNU General Public License as published by")
    print(" the Free Software Foundation, version 3.")
    print(" Copyright © 2013 P. Durrant, K. Hendricks, S. Siebert, fandrieu, DiapDealer, nickredding.")

def main(argv=unicode_argv()):
    print("DualMobiMetaFixer v005")
    progname = os.path.basename(argv[0])
    try:
        opts, args = getopt.getopt(argv[1:], "h")
    except getopt.GetoptError as err:
        print(str(err))
        usage(progname)
        sys.exit(2)

    if len(args) != 3:
        usage(progname)
        sys.exit(2)

    for o, a in opts:
        if o == "-h":
            usage(progname)
            sys.exit(0)

    asin = args[0]
    infile = args[1]
    outfile = args[2]
    # print("ASIN:   ", asin)
    # print("Input:  ", infile)
    # print("Output: ", outfile)

    try:
        # make sure it is really a mobi ebook
        infileext = os.path.splitext(infile)[1].upper()
        if infileext not in ['.MOBI', '.PRC', '.AZW', '.AZW3','.AZW4']:
            raise DualMetaFixException('second parameter must be a Kindle/Mobipocket ebook.')
        mobidata = open(pathof(infile), 'rb').read(100)
        palmheader = mobidata[0:78]
        ident = palmheader[0x3C:0x3C+8]
        if ident != b'BOOKMOBI':
            raise DualMetaFixException('invalid file format not BOOKMOBI')

        dmf = DualMobiMetaFix(infile, asin)
        open(pathof(outfile),'wb').write(dmf.getresult())

    except Exception as e:
        print("Error: %s" % e)
        print(traceback.format_exc())
        return 1

    print("Success")
    return 0


if __name__ == '__main__':
    sys.exit(main())
