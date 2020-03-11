class Color(object):
    __slots__ = ("r", "g", "b", "a")

    def __init__(self, r=255, g=255, b=255, a=255):
        self.r = r
        self.g = g
        self.b = b
        self.a = a

    @property
    def hex(self):
        return "%0.2X%0.2X%0.2X%0.2X" % (self.a, self.r, self.g, self.b)

    def __hash__(self):
        return (self.a << 24) + (self.r << 16) + (self.g << 8) + (self.b)

    def __eq__(self, other):
        if not other:
            return False
        return (
            self.r == other.r
            and self.g == other.g
            and self.b == other.b
            and self.a == other.a
        )

    def __str__(self):
        return self.hex


Color.WHITE = Color(255, 255, 255, 255)
Color.BLACK = Color(0, 0, 0, 255)
