# officecli — Python SDK

A **thin** Python SDK for the [officecli](https://officecli.ai) **resident pipe**. It does one
thing: forward an officecli command to a running resident over its named pipe and
hand back the response — no per-command process spawn, so a loop of edits is
~hundreds of times faster than shelling out to the CLI per command.

"Thin" is the point: there is **no second vocabulary** to learn. A command is the
same dict you'd put in an officecli `batch` list; the SDK just carries it over the
pipe. Anything a `doc.set_cell(...)` / `doc.add_paragraph(...)` method would do is
**fully supported** — you just spell it `doc.send({"command": "set", ...})`, with
the exact same effect. One uniform verb instead of dozens of per-element named
methods: same power, nothing extra to memorize, and new officecli features work
the day they ship without an SDK update.

## The officecli CLI

`pip install officecli-sdk` installs only this SDK. The `officecli` binary must
already be on `PATH`, or be supplied through the `binary` argument to
`open()`/`create()`. If it cannot be found, the SDK raises a clear error.

## Install

```bash
pip install officecli-sdk            # once published — note: import name is `officecli`
# or, from a checkout of this repo:
pip install ./sdk/python
```

The pip/distribution name is `officecli-sdk`, but you `import officecli`
(distribution name ≠ import name, like `pip install pillow` → `import PIL`).

Zero third-party dependencies (standard library only).

## Quickstart

```python
import officecli

# create() makes a new file and returns a live session handle;
# open() does the same for an existing file. Both return a Document.
with officecli.create("report.xlsx", "--force") as doc:
    doc.send({"command": "set", "path": "/Sheet1/A1",
              "props": {"text": "Region", "bold": "true"}})
    doc.send({"command": "set", "path": "/Sheet1/B1", "props": {"formula": "=SUM(B2:B9)"}})

    # read one back (returns the parsed JSON envelope)
    node = doc.send({"command": "get", "path": "/Sheet1/A1"})
    print(node["data"]["results"][0]["text"])     # -> Region

    # many edits in ONE pipe round-trip
    doc.batch([
        {"command": "set", "path": "/Sheet1/A2", "props": {"text": "North"}},
        {"command": "set", "path": "/Sheet1/A3", "props": {"text": "South"}},
    ])

    doc.send({"command": "save"})
# leaving `with` closes the resident (which flushes to disk)

# borrow an already-running resident without owning it: skip `with`/close()
d = officecli.open("report.xlsx")
print(d.send({"command": "view", "mode": "stats"}, as_json=False))
```

See `demo.py` for a fuller example.

## The command dict

`send(item)` and `batch([item, ...])` take the officecli **batch-item** shape:

```jsonc
{ "command": "set",            // or "op"; picks the officecli command
  "path": "/Sheet1/A1",        // every key except command/op/props is forwarded
  "props": { "text": "hi" } }  // verbatim as a command argument
```

Keys are officecli's own batch fields (`command`/`op`, `path`, `parent`, `type`,
`index`, `after`, `before`, `to`, `selector`, `mode`, `depth`, `part`, `xpath`,
`action`, `xml`) plus a nested `props`. The client maintains no field list of its
own — run `officecli help` (or see the batch docs) for the full reference.

`send(..., as_json=False)` requests plain-text output (e.g. `view` / `raw` /
`dump`), mirroring the CLI's `--json` toggle.

## Errors & resilience

- Transport/process failures raise `officecli.OfficeCliError` (`.code` carries the
  exit code). Business outcomes (e.g. `validate` failing, a bad path) are **not**
  exceptions — they live in the returned envelope's `success` field, same as the
  CLI's exit code.
- If the resident has gone (crash, idle-timeout, missing pipe), `send`/`batch`
  transparently restart it and retry once. If it's alive but the pipe is
  unresponsive (busy), they raise rather than risk racing the live resident.

## Versioning

This client derives the resident's pipe address from the document path the same
way officecli does. That derivation is the one piece coupled to officecli
internals, so keep the client version compatible with your installed officecli.
