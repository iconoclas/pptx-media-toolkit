# pptx-media-toolkit

Small stdlib-only Python utilities for working with large PowerPoint files.

- `shrink_pptx.py` - strip/replace heavy media so a PPTX can be sent to translation tools with file-size limits
- `merge_translated_pptx.py` - inject translated slide text (e.g. from DeepL) back into the original PPTX, keeping the original media. Output is plain ZIP32 with XML first and media last, avoiding the PowerPoint "repair" prompt on large decks

## Usage

```bash
python3 shrink_pptx.py --help
python3 merge_translated_pptx.py --orig original.pptx --trans translated.pptx -o merged.pptx
```

Requires Python 3.8+. No third-party dependencies.

## License

MIT
