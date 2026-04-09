# MSG Gotchas

Recommended default parser:

- `extract-msg`

Why:

- established OSS option
- broad `.msg` coverage
- practical first implementation choice

Main risk:

- license profile is stricter than permissive MIT-style packages, so confirm compatibility before distribution decisions.

Fallback:

- `python-oxmsg` if a permissive license is mandatory, but it is materially less mature.
