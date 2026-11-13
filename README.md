# docx-redline

Adds word-level tracked changes and comments to an existing Word file under a
named author, without disturbing comments and revisions already in it. It works
on the .docx package directly, so Word attributes every change to that author,
just as if they had typed it.

```bash
uv run redline.py IN.docx --dump > visible.txt       # current text, one paragraph per line
uv run redline.py IN.docx ops.json OUT.docx --author "Barry Brook" --initials BB
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

- An empty `old` gives a pure insertion; an empty `new` gives a pure deletion.
- Put a `com` before a `rep` on the same text so the comment covers both.
- If any op fails, nothing is written. The input file is never overwritten.

To check the result, Anthropic's docx skill validator (`validate.py OUT.docx
--original IN.docx --author "..."`) flags untracked text changes, and
`pandoc --track-changes=all -t markdown OUT.docx` shows every change inline.
