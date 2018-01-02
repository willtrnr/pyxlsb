from .tokenreader import TokenReader

class Formula(object):
    def __init__(self, tokens=None):
        super(Formula, self).__init__()
        self._tokens = list(tokens) if tokens else list()

    def __repr__(self):
        return 'Formula({})'.format(self._tokens)

    def __str__(self):
        return self.stringify()

    def stringify(self):
        tokens = self._tokens[:]
        return tokens.pop().stringify(tokens)

    @classmethod
    def parse(cls, data):
        return cls(TokenReader(data))
