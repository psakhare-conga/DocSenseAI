#!/usr/bin/env python3
"""
embed-addin.py — Embeds the DocSense AI Word Add-in reference into a .docx file.

When the modified .docx is opened in Word Desktop or Word Online, Word reads
the embedded webextension entry and automatically opens the DocSense AI task pane.

────────────────────────────────────────────────────────────────────────────────
SHARING REQUIREMENTS
────────────────────────────────────────────────────────────────────────────────
Scenario A – Same machine / dev demo
  storeType = "developer"  ← default below
  Works only on machines where manifest.xml has already been sideloaded.

Scenario B – Share with anyone in your org via Word Online (RECOMMENDED)
  1. Deploy the add-in frontend to a public HTTPS URL
     (free options: Vercel, Azure Static Web Apps, GitHub Pages + Actions)
  2. Update ADDIN_URL below to that public URL
  3. An M365 admin deploys the manifest via:
     admin.microsoft.com → Settings → Integrated Apps → Upload custom app
  4. Change STORE_TYPE = "Centralized" and STORE = ""
  After admin deployment, the add-in auto-loads for everyone in the org
  when they open any document that contains the embedded reference.

Scenario C – Public / AppSource
  Publish to Microsoft AppSource, then set STORE_TYPE = "OMEX"

Usage:
  python embed-addin.py input.docx
  python embed-addin.py input.docx output.docx
────────────────────────────────────────────────────────────────────────────────
"""

import zipfile
import shutil
import os
import sys
import re

# ── Configuration ─────────────────────────────────────────────────────────────
ADDIN_ID    = "3b9d2a8c-4f1e-4d6b-a5c3-9e2f1b4d7a8e"   # must match manifest.xml
ADDIN_TITLE = "DocSense AI"
ADDIN_URL   = "https://localhost:3000/taskpane.html"     # change to public URL for sharing

# developer  → sideloaded (localhost / local network)
# Centralized → M365 Admin deployed (set STORE = "")
# OMEX        → AppSource published
STORE_TYPE  = "developer"
STORE       = "en-US"   # set to "" for Centralized deployment
# ──────────────────────────────────────────────────────────────────────────────

WEBEXT_CONTENT_TYPE   = "application/vnd.ms-office.webextension+xml"
TASKPANE_CONTENT_TYPE = "application/vnd.ms-office.webextensiontaskpanes+xml"

REL_TYPE_TASKPANES = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
REL_TYPE_WEBEXT    = "http://schemas.microsoft.com/office/2011/relationships/webextension"

WEBEXTENSION_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension
  xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
  id="{{{ADDIN_ID}}}">
  <we:reference
    id="{ADDIN_ID}"
    version="1.0.0.0"
    store="{STORE}"
    storeType="{STORE_TYPE}"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="addinTitle" value="{ADDIN_TITLE}"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>"""

TASKPANE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes
  xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="4">
    <wetp:webextensionref
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      r:id="rId_we1"/>
  </wetp:taskpane>
</wetp:taskpanes>"""

TASKPANE_RELS_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rId_we1"
    Type="{REL_TYPE_WEBEXT}"
    Target="../webextensions/webextension1.xml"/>
</Relationships>"""


def embed_addin(input_path: str, output_path: str) -> None:
    if not os.path.exists(input_path):
        print(f"Error: file not found → {input_path}")
        sys.exit(1)

    # Read all existing zip entries
    files: dict[str, bytes] = {}
    with zipfile.ZipFile(input_path, "r") as z:
        for name in z.namelist():
            files[name] = z.read(name)

    # ── 1. [Content_Types].xml ────────────────────────────────────────────────
    ct = files["[Content_Types].xml"].decode("utf-8")
    if WEBEXT_CONTENT_TYPE not in ct:
        ct = ct.replace(
            "</Types>",
            f'  <Override PartName="/word/webextensions/webextension1.xml" ContentType="{WEBEXT_CONTENT_TYPE}"/>\n'
            f'  <Override PartName="/word/taskpanes/taskpane1.xml" ContentType="{TASKPANE_CONTENT_TYPE}"/>\n'
            "</Types>",
        )
        files["[Content_Types].xml"] = ct.encode("utf-8")
        print("  [Content_Types].xml  → updated")
    else:
        print("  [Content_Types].xml  → already contains webextension entry, skipped")

    # ── 2. word/_rels/document.xml.rels ──────────────────────────────────────
    rels_key = "word/_rels/document.xml.rels"
    rels = files[rels_key].decode("utf-8")
    if REL_TYPE_TASKPANES not in rels:
        existing_nums = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels)]
        next_id = max(existing_nums, default=0) + 1
        rels = rels.replace(
            "</Relationships>",
            f'  <Relationship Id="rId{next_id}" '
            f'Type="{REL_TYPE_TASKPANES}" '
            f'Target="../taskpanes/taskpane1.xml"/>\n'
            "</Relationships>",
        )
        files[rels_key] = rels.encode("utf-8")
        print("  document.xml.rels    → taskpanes relationship added")
    else:
        print("  document.xml.rels    → taskpanes relationship already present, skipped")

    # ── 3. New parts ──────────────────────────────────────────────────────────
    files["word/webextensions/webextension1.xml"]       = WEBEXTENSION_XML.encode("utf-8")
    files["word/taskpanes/taskpane1.xml"]               = TASKPANE_XML.encode("utf-8")
    files["word/taskpanes/_rels/taskpane1.xml.rels"]    = TASKPANE_RELS_XML.encode("utf-8")
    print("  webextension1.xml    → written")
    print("  taskpane1.xml        → written")
    print("  taskpane1.xml.rels   → written")

    # ── 4. Write output ───────────────────────────────────────────────────────
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)

    print()
    print(f"✓  Output → {output_path}")
    print(f"   Add-in ID  : {ADDIN_ID}")
    print(f"   Store type : {STORE_TYPE}")
    print(f"   Task pane  : {ADDIN_URL}")
    print()

    if STORE_TYPE == "developer":
        print("⚠  DEVELOPER MODE — add-in auto-loads only on machines where")
        print("   manifest.xml has already been sideloaded via Word > Insert > Add-ins.")
        print()
        print("   To share with any user via Word Online:")
        print("   1. Deploy the add-in to a public HTTPS URL")
        print("      (free: https://vercel.com  or  Azure Static Web Apps)")
        print("   2. Update ADDIN_URL and set STORE_TYPE = 'Centralized' in this script")
        print("   3. M365 admin deploys manifest via admin.microsoft.com")
        print("      Settings → Integrated Apps → Upload custom app → paste manifest URL")
        print("   4. Re-run this script — add-in will auto-load for the whole org")
    elif STORE_TYPE == "Centralized":
        print("✓  CENTRALIZED DEPLOYMENT — add-in will auto-load for all users in")
        print("   the org once the admin has deployed the manifest.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python embed-addin.py input.docx [output.docx]")
        sys.exit(1)

    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else inp.replace(".docx", "-with-addin.docx")
    if out == inp:
        out = inp.replace(".docx", "-with-addin.docx")

    print(f"\nEmbedding DocSense AI add-in into: {inp}")
    print("─" * 50)
    embed_addin(inp, out)
