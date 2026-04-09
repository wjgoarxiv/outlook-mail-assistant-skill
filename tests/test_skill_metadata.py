from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_skill_frontmatter_has_required_fields():
    text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    lines = text.splitlines()

    assert lines[0] == "---"
    end_index = lines.index("---", 1)
    frontmatter = {}
    for line in lines[1:end_index]:
        key, _, value = line.partition(":")
        frontmatter[key.strip()] = value.strip().strip('"')

    assert frontmatter["name"] == "outlook-mail-assistant"
    assert frontmatter["description"]


def test_openai_yaml_mentions_explicit_skill_name():
    openai_yaml = (ROOT / "agents" / "openai.yaml").read_text(encoding="utf-8")

    assert "$outlook-mail-assistant" in openai_yaml
