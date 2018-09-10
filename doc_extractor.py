#!/usr/bin/env python
# -*- encoding=utf8 -*-
import re
import sys
import struct
import zipfile
import olefile

class OLEFile:
    def __init__(self):
        self.stream_size = 0
        self.flags1 = 0
        self.label = ""
        self.filename = ""
        self.flags2 = 0
        self.unknown1 = b""
        self.unknown2 = b""
        self.native_data_size = 0
        self.native_data = b""

    def parse_data(self, stream):
        self.stream_size, = struct.unpack("<i", stream.read(4))
        self.flags1, = struct.unpack("<H", stream.read(2))
        self.label = self.readstring(stream)
        self.label = self.label.decode("gb18030") # Chinese encoding
        self.filename = self.readstring(stream)
        self.flags2, = struct.unpack("<H", stream.read(2))
        unknown_len, = struct.unpack("<B", stream.read(1))
        self.unknown1, = struct.unpack("<%dp" % (unknown_len), stream.read(unknown_len))
        self.unknown2, = struct.unpack("<2p", stream.read(2))
        command = self.readstring(stream)
        self.native_data_size, = struct.unpack("<i", stream.read(4))
        self.native_data = stream.read(self.native_data_size)

    def readstring(self, stream):
        s = ""
        c = stream.read(1)
        while c != "\0":
            s += c
            c = stream.read(1)
        return s


def extract_embedded_files(docfile):
    zip_ref = zipfile.ZipFile(docfile, "r")
    print "Extract following files to current directory"
    for name in zip_ref.namelist():
        m = re.compile("^word/embeddings/oleObject\d+.bin$")
        if m.match(name):
            ole_data = zip_ref.open(name).read()
            (filename, data) = extract_embedded_file(ole_data)
            if filename:
                print "extract %s" % filename
                with open(filename, "wb") as f:
                    f.write(data)
            else:
                print "extract %s failed" % filename
    zip_ref.close()


def extract_embedded_file(ole_file_data):
    ole = olefile.OleFileIO(ole_file_data)
    if ole.exists('\x01Ole10Native'):
        native_data_stream = ole.openstream('\x01Ole10Native')
        f = OLEFile()
        f.parse_data(native_data_stream)
    return (f.label, f.native_data)


if __name__ == "__main__":
    for docfile in sys.argv[1:]:
        extract_embedded_files(docfile)
