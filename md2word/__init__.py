__all__ = ["MD2Word"]


def __getattr__(name):
    if name == "MD2Word":
        from md2word.converter import MD2Word
        return MD2Word
    raise AttributeError(f"module 'md2word' has no attribute {name!r}")
