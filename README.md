# docx-redline

> **Status:** archived on 19 November 2026 and no longer maintained. Kept as a
> working record; fork it if it is useful to you.

Adds word-level tracked changes and comments to an existing Word file under a
named author, without disturbing comments and revisions already in it. It works
on the .docx package directly, so Word attributes every change to that author,
just as if they had typed it. Built for fast, precise manuscript review: an AI
agent (or a person) writes a list of edits and comments, and the tool applies
them with exact character-level placement.

## Usage

Needs [uv](https://docs.astral.sh/uv/); the script declares its own dependency
(`lxml`).

```bash
uv run redline.py IN.docx --dump > visible.txt       # current text, one paragraph per line
uv run redline.py IN.docx ops.json OUT.docx --author "J. Citizen" --initials JC
```

`ops.json` is a list applied in order. Each op finds `b + target + a` in the
current view (insertions shown, deletions hidden). That string must be unique,
and the target must not sit inside someone else's tracked change.

```json
[
  {"op": "rep", "b": "Hello ", "old": "brave", "new": "bold", "a": " new"},
  {"op": "com", "b": "", "t": "Second", "a": " para", "c": "Quick note on this."}
]
```

- `--author` defaults to "J. Citizen" and `--initials` to "JC". To have changes
  show as yours, pass your name exactly as it appears on your existing
  comments in Word.
- An empty `old` gives a pure insertion; an empty `new` gives a pure deletion.
- Put a `com` before a `rep` on the same text so the comment covers both.
- If any op fails, nothing is written. The input file is never overwritten.

To check the result, `pandoc --track-changes=all -t markdown OUT.docx` shows
every change and comment inline.

## How this work is done

Developed by Barry W. Brook with Claude Opus 5.5 (Anthropic) in Claude Code.
Claude wrote the code and documentation while using it to review a manuscript
with tracked changes, and tested it against Word files with existing multi-author
comments and revisions; Barry set the requirements and reviewed the result.

## Citation

See [`CITATION.cff`](CITATION.cff).

## Licence

Code is released under the [MIT Licence](LICENSE). Text is released under
[CC BY-SA 4.0](LICENSE-CONTENT).
