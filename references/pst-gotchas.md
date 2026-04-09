# PST Gotchas

Primary recommendation order:

1. `libratom`
2. `Aspose.Email for Python via .NET` fallback when `libratom` is not viable in the active Python runtime

Tradeoff:

- strongest OSS path for `.pst`
- heavier Windows setup than ideal for lightweight deployment
- current `Python 3.13` environment may block `libratom` dependency resolution

Commercial low-friction option:

- `Aspose.Email for Python via .NET`

Tradeoff:

- lower engineering risk
- commercial licensing and vendor dependency
- may be easier to install than the OSS path in some Windows/Python environments

Implementation rule:

- keep `.pst` support behind an isolated adapter boundary
- treat `.pst` as the most fragile ingestion path in the system
- export `.pst` imports through the same `canonical records -> JSONL -> SQLite` path as Outlook live export
