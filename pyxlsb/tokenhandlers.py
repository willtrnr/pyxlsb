from collections import namedtuple

class TokenHandler(object):
    def __init__(self):
        super(TokenHandler, self).__init__()

    def read(self, reader, tokenid):
        return None


class BasicTokenHandler(TokenHandler):
    def __init__(self, name=None):
        super(BasicTokenHandler, self).__init__()
        self.name = name
        self.cls = namedtuple(name, ['tokenid'])

    def read(self, reader, tokenid):
        super(BasicTokenHandler, self).read(reader, tokenid)
        return self.cls(tokenid)


class FuncHandler(TokenHandler):
    cls = namedtuple('tFunc', ['tokenid', 'idx'])

    def __init__(self):
        super(FuncHandler, self).__init__()

    def read(self, reader, tokenid):
        idx = reader.read_short()
        return self.cls(tokenid, idx)


class FuncVarHandler(TokenHandler):
    cls = namedtuple('tFuncVar', ['tokenid', 'args', 'prompt', 'idx', 'is_macro'])

    def __init__(self):
        super(FuncVarHandler, self).__init__()

    def read(self, reader, tokenid):
        args = reader.read_byte()
        idx = reader.read_short()
        return self.cls(tokenid, args & 0x7F, args & 0x80 != 0, idx & 0x7FFFF, idx & 0x8000 != 0)


class FuncCeHandler(TokenHandler):
    cls = namedtuple('tFuncCE', ['tokenid', 'args', 'idx'])

    def __init__(self):
        super(FuncCeHandler, self).__init__()

    def read(self, reader, tokenid):
        args = reader.read_byte()
        idx = reader.read_byte()
        return self.cls(tokenid, args, idx)
