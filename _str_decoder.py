""" code from `here <https://stackoverflow.com/a/24519338>`__ """

import re
import codecs

ESCAPE_SEQUENCE_RE = re.compile(r'''
    ( \\U........      # 8-digit hex escapes
    | \\u....          # 4-digit hex escapes
    | \\x..            # 2-digit hex escapes
    | \\[0-7]{1,3}     # Octal escapes
    | \\N\{[^}]+\}     # Unicode characters by name
    | \\[\\'"abfnrtv]  # Single-character escapes
    )''', re.UNICODE | re.VERBOSE)

def decode_escapes(s):
    def decode_match(match):
        try:
            return codecs.decode(match.group(0), 'unicode-escape')
        except UnicodeDecodeError:
            # In case we matched the wrong thing after a double-backslash
            return match.group(0)

    return ESCAPE_SEQUENCE_RE.sub(decode_match, s)