import sys

if sys.version_info > (3,):
    xrange = range

try:
    from hexdump import hexdump as hd
    def _hexdump(data):
        return hd(data, result='return')
except ImportError:
    def _hexdump(data):
        def chunked(v, s):
            for i in xrange(0, len(v), s):
                yield v[i:i + s]

        def hexed(it):
            conv = (lambda x: x) if sys.version_info > (3,) else ord
            for chunk in it:
                yield tuple(['%02X' % conv(b) for b in chunk])

        def formatted(d):
            for i, c in enumerate(hexed(chunked(d, 16))):
                yield '%08X: %s  %s' % (i * 16, ' '.join(c[0:8]), ' '.join(c[8:16]))

        return '\n'.join(formatted(data))
