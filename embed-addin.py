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
ADDIN_TITLE = "Conga AI Assistance"
ADDIN_URL   = "https://blue-mushroom-008de5a0f.7.azurestaticapps.net/taskpane.html"

# developer  → sideloaded (localhost / local network)
# EXCatalog   → M365 Admin deployed (Centralized Deployment)
# OMEX        → AppSource published
STORE_TYPE  = "EXCatalog"
STORE       = "EXCatalog"   # must match storeType for Centralized deployment
# ──────────────────────────────────────────────────────────────────────────────

WEBEXT_CONTENT_TYPE   = "application/vnd.ms-office.webextension+xml"
TASKPANE_CONTENT_TYPE = "application/vnd.ms-office.webextensiontaskpanes+xml"

REL_TYPE_TASKPANES = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
REL_TYPE_WEBEXT    = "http://schemas.microsoft.com/office/2011/relationships/webextension"

def make_webextension_xml(addin_id, store, store_type, addin_title):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension
  xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
  id="{{{addin_id}}}">
  <we:reference
    id="{addin_id}"
    version="1.0.0.0"
    store="{store}"
    storeType="{store_type}"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="addinTitle" value="{addin_title}"/>
    <we:property name="TaskpaneId" value="DocSenseTaskPane"/>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>"""


def embed_addin(input_path: str, output_path: str) -> None:
    if not os.path.exists(input_path):
        print(f"Error: file not found → {input_path}")
        sys.exit(1)

    # Read all existing zip entries
    files: dict[str, bytes] = {}
    with zipfile.ZipFile(input_path, "r") as z:
        for name in z.namelist():
            files[name] = z.read(name)

    # ── Discover existing webextension slots ──────────────────────────────────
    existing_webexts = sorted(
        n for n in files if re.match(r"word/webextensions/webextension\d+\.xml", n)
    )
    print(f"  Found {len(existing_webexts)} existing webextension(s): {existing_webexts}")

    # Check if our add-in is already the one in slot 1 (nothing to do)
    if existing_webexts:
        first = files[existing_webexts[0]].decode("utf-8", errors="replace")
        if ADDIN_ID in first:
            print(f"  Our add-in is already embedded as {existing_webexts[0]}, nothing to do.")
            import shutil
            shutil.copy2(input_path, output_path)
            print(f"\n✓  Copied unchanged → {output_path}")
            return

    # ── Strategy: replace the first (or only) webextension slot with ours ────
    # This removes the old add-in and puts ours in its place, reusing the
    # existing taskpanes.xml / taskpanes.xml.rels references unchanged.
    target_slot = existing_webexts[0] if existing_webexts else "word/webextensions/webextension1.xml"
    print(f"  Replacing {target_slot} with our add-in (removing old add-in)")

    files[target_slot] = make_webextension_xml(ADDIN_ID, STORE, STORE_TYPE, ADDIN_TITLE).encode("utf-8")

    # Remove any extra webextension slots (2, 3, ...) that may have been added
    # by a previous run of this script
    for slot in existing_webexts[1:]:
        del files[slot]
        print(f"  Removed extra slot: {slot}")

    # ── Ensure Content_Types has the webextension entry ───────────────────────
    ct = files["[Content_Types].xml"].decode("utf-8")
    slot_name = target_slot.split("/")[-1]   # e.g. webextension1.xml
    if slot_name not in ct:
        ct = ct.replace(
            "</Types>",
            f'  <Override PartName="/word/webextensions/{slot_name}" ContentType="{WEBEXT_CONTENT_TYPE}"/>\n</Types>',
        )
        files["[Content_Types].xml"] = ct.encode("utf-8")
        print(f"  [Content_Types].xml  → {slot_name} entry added")

    # ── Ensure taskpanes.xml exists and shows the pane (visibility=1) ─────────
    tp_key = "word/webextensions/taskpanes.xml"
    rels_key = "word/webextensions/_rels/taskpanes.xml.rels"

    if tp_key not in files:
        # Build from scratch pointing to rId1 → webextension1.xml
        files[rels_key] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL_TYPE_WEBEXT}" Target="{slot_name}"/>'
            "</Relationships>"
        ).encode("utf-8")
        files[tp_key] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">'
            '<wetp:taskpane dockstate="right" visibility="1" width="350" row="4">'
            '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
            "</wetp:taskpane></wetp:taskpanes>"
        ).encode("utf-8")
        # Content_Types for taskpanes
        ct = files["[Content_Types].xml"].decode("utf-8")
        if TASKPANE_CONTENT_TYPE not in ct:
            ct = ct.replace("</Types>", f'  <Override PartName="/word/webextensions/taskpanes.xml" ContentType="{TASKPANE_CONTENT_TYPE}"/>\n</Types>')
            files["[Content_Types].xml"] = ct.encode("utf-8")
        # _rels/.rels
        root_rels = files["_rels/.rels"].decode("utf-8")
        if REL_TYPE_TASKPANES not in root_rels:
            root_rels = root_rels.replace("</Relationships>", f'  <Relationship Id="rId_tp" Type="{REL_TYPE_TASKPANES}" Target="word/webextensions/taskpanes.xml"/>\n</Relationships>')
            files["_rels/.rels"] = root_rels.encode("utf-8")
        print(f"  taskpanes.xml        → created")
    else:
        # Ensure visibility="1" on the first taskpane
        tp = files[tp_key].decode("utf-8")
        tp_fixed = re.sub(r'visibility="0"', 'visibility="1"', tp)
        if tp_fixed != tp:
            files[tp_key] = tp_fixed.encode("utf-8")
            print(f"  taskpanes.xml        → visibility set to 1")
        else:
            print(f"  taskpanes.xml        → unchanged (visibility already 1)")

    # ── Write output ──────────────────────────────────────────────────────────
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)

    print()
    print(f"✓  Output → {output_path}")
    print(f"   Add-in ID  : {ADDIN_ID}")
    print(f"   Store type : {STORE_TYPE}")
    print(f"   Task pane  : {ADDIN_URL}")
    print()

    if STORE_TYPE == "EXCatalog":
        print("✓  CENTRALIZED DEPLOYMENT — add-in will auto-load for all users in")
        print("   the org once the admin has deployed the manifest.")
    elif STORE_TYPE == "developer":
        print("⚠  DEVELOPER MODE — add-in auto-loads only on machines where")
        print("   manifest.xml has already been sideloaded via Word > Insert > Add-ins.")


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
