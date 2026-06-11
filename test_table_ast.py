from mistune import Markdown

class ASTRenderer:
    def __call__(self, tokens, state):
        return list(tokens)

md = Markdown(renderer=ASTRenderer())
text = """| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
| Cell **3** | Cell 4   |"""

result = md(text)
import json

def convert(obj):
    if isinstance(obj, dict):
        return {k: convert(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert(v) for v in obj]
    return obj

print(json.dumps(convert(result), indent=2, ensure_ascii=False))
