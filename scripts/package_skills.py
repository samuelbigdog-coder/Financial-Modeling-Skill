"""Build and validate one deterministic Cowork ZIP per skill."""
from hashlib import sha256
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parents[1]
SKILLS = ROOT / "skills"
DIST = ROOT / "dist"
EXPECTED = {"financial-modeling", "valuation-discipline", "equity-research-analyst"}


def package(folder):
    name = folder.name
    files = sorted(folder.rglob("*.md"))
    entry = folder / "SKILL.md"
    content = entry.read_text(encoding="utf-8-sig")
    frontmatter = re.match(
        r"\A---\nname: ([a-z][a-z-]+)\ndescription: (.{1,1024})\n---\n",
        content,
    )
    if not frontmatter or frontmatter[1] != name:
        raise ValueError(f"Invalid skill frontmatter: {entry}")

    expected_members = {}
    for file in files:
        if file.is_symlink() or not file.resolve().is_relative_to(folder.resolve()):
            raise ValueError(f"Unexpected path: {file}")
        text = file.read_text(encoding="utf-8-sig")
        for target in re.findall(r"\]\(([^)]+)\)", text):
            if target.startswith(("https://", "http://", "#")):
                continue
            local_file = (file.parent / target.split("#", 1)[0]).resolve()
            if local_file not in files:
                raise ValueError(f"Missing packaged link {target} in {file}")
        expected_members[file.relative_to(SKILLS).as_posix()] = text.encode("utf-8")

    archive = DIST / f"{name}.zip"
    with ZipFile(archive, "w") as output:
        for member, payload in expected_members.items():
            info = ZipInfo(member)
            info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            output.writestr(info, payload)

    with ZipFile(archive) as output:
        if output.testzip() or output.namelist() != list(expected_members):
            raise ValueError(f"Invalid ZIP inventory: {archive}")
        for member, payload in expected_members.items():
            if output.read(member) != payload:
                raise ValueError(f"ZIP payload mismatch: {member}")
    return f"{sha256(archive.read_bytes()).hexdigest()}  {archive.name}"


if __name__ == "__main__":
    folders = sorted(folder for folder in SKILLS.iterdir() if folder.is_dir())
    if {folder.name for folder in folders} != EXPECTED:
        raise SystemExit("Expected exactly three skill directories")
    DIST.mkdir(exist_ok=True)
    manifest = "\n".join(package(folder) for folder in folders) + "\n"
    (DIST / "SHA256SUMS").write_text(manifest, encoding="ascii", newline="\n")
    print(manifest, end="")
