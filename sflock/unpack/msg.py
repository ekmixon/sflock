# Copyright (C) 2016-2018 Jurriaan Bremer.
# This file is part of SFlock - http://www.sflock.org/.
# See the file 'docs/LICENSE.txt' for copying permission.

import six

from sflock.abstracts import Unpacker, File

class MsgFile(Unpacker):
    name = "msgfile"
    exts = b".msg"

    def supported(self):
        return True

    def handles(self):
        if super(MsgFile, self).handles():
            return True

        if self.f.ole:
            for filename in self.f.ole.listdir():
                if filename[0].startswith("__attach"):
                    return True
        return False

    def get_stream(self, *filename):
        if self.f.ole.exists("/".join(filename)):
            return self.f.ole.openstream("/".join(filename)).read()

    def get_string(self, *filename):
        ascii_filename = f'{"/".join(filename)}001E'
        unicode_filename = f'{"/".join(filename)}001F'

        if stream := self.get_stream(unicode_filename):
            return stream.decode("utf16")

        return self.get_stream(ascii_filename)

    def get_attachment(self, dirname):
        filename = (
            self.get_string(dirname, "__substg1.0_3707") or
            self.get_string(dirname, "__substg1.0_3704") or
            "att1"
        )
        contents = self.get_stream(dirname, "__substg1.0_37010102")
        return filename, contents

    def unpack(self, password=None, duplicates=None):
        seen, entries = [], []

        if not self.f.ole:
            self.f.mode = "failed"
            self.f.error = "No OLE structure found"
            return []

        for dirname in self.f.ole.listdir():
            if dirname[0].startswith("__attach") and dirname[0] not in seen:
                filename, contents = self.get_attachment(dirname[0])
                if six.PY3 and isinstance(filename, str):
                    filename = filename.encode()
                entries.append(File(
                    relapath=filename, contents=contents
                ))
                seen.append(dirname[0])

        return self.process(entries, duplicates)
