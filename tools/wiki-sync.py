# -*- coding: utf-8 -*-
"""Builds the wiki copy from docs/.

The documentation lives in docs/, where links point at .md files so that they work
while browsing the repository. A GitHub wiki wants links by page name, without the
extension, and wants the index page called Home.md. This script does that translation.

Usage:

    python tools/wiki-sync.py ../psu3d-wiki

Then, inside the output folder:

    git add -A && git commit -m "update wiki" && git push

The wiki repository only comes into existence after the first page is created through
the GitHub interface, under Wiki and then Create the first page. Once, by hand.
"""
import io
import os
import re
import sys
import glob

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(os.path.dirname(HERE), "docs")

SIDEBAR = """**Psu3D**

[Home](Home)

**Getting started**

[Installation](Installation)
[Getting started](Getting-Started)
[Concepts](Concepts)
[How the 3D works](How-The-3D-Works)

**Guides**

[Canvas](Canvas)
[Camera](Camera)
[Renderer and primitives](Renderer-And-Primitives)
[Scene](Scene)
[Materials](Materials)
[Lighting and fog](Lighting-And-Fog)
[Body and physics](Body-And-Physics)
[Levels in JSON](Levels-In-JSON)

**Internals**

[Depth sorting](Depth-Sorting)
[Spatial index](Spatial-Index)
[Performance](Performance)
[Known limits](Known-Limits)

**Practice**

[Recipes](Recipes)
[Examples](Examples)
[Self test](Self-Test)
[Troubleshooting](Troubleshooting)
[FAQ](FAQ)

**Reference**

[PCore](API-PCore)
[PLighting](API-PLighting)
[PMaterial](API-PMaterial)
[PMaterials](API-PMaterials)
[PCanvas](API-PCanvas)
[PCamera](API-PCamera)
[PRenderer](API-PRenderer)
[PScene](API-PScene)
[PBody](API-PBody)
[PLevel](API-PLevel)
[Psu3D](API-Psu3D)
[PSelfTest](API-PSelfTest)
"""


def main(out_dir):
    if not os.path.isdir(DOCS):
        print("docs/ not found at", DOCS)
        return 1

    os.makedirs(out_dir, exist_ok=True)
    written = 0

    for path in sorted(glob.glob(os.path.join(DOCS, "*.md"))):
        name = os.path.basename(path)
        text = io.open(path, encoding="utf-8", newline="").read().replace("\r\n", "\n")

        # a file link becomes a page link
        text = re.sub(r"\]\((?!http)([A-Za-z0-9._-]+)\.md(#[^)]*)?\)",
                      lambda m: "](%s%s)" % (m.group(1), m.group(2) or ""), text)
        # the repository index is the wiki Home
        text = text.replace("](README)", "](Home)")

        target = "Home.md" if name == "README.md" else name
        io.open(os.path.join(out_dir, target), "w", encoding="utf-8", newline="\n").write(text)
        written += 1

    io.open(os.path.join(out_dir, "_Sidebar.md"), "w", encoding="utf-8", newline="\n").write(SIDEBAR)
    print("%d pages written to %s" % (written + 1, out_dir))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
