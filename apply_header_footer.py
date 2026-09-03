#!/usr/bin/env python3
"""
Batch-apply a header image and footer image to a folder of .docx files.

Usage:
    python3 apply_header_footer.py <input_dir> <header_img> <footer_img> <output_dir>
"""
import os
import re
import shutil
import subprocess
import sys
import zipfile

# Fixed EMU extents that match the reference document (Lab_Header_UpScaled / FOOTER_new)
HEADER_CX, HEADER_CY = 5943600, 1363345
FOOTER_CX, FOOTER_CY = 6267450, 591820

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def build_part_xml(tag, style, cx, cy, doc_pr_id, pic_name, r_id):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:{tag} xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        f'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        f'xmlns:o="urn:schemas-microsoft-com:office:office" '
        f'xmlns:r="{NS_R}" '
        f'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        f'xmlns:v="urn:schemas-microsoft-com:vml" '
        f'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        f'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        f'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        f'mc:Ignorable="w14 wp14">'
        f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
        f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_pr_id}" name="{pic_name}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="1" name="{pic_name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{r_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:{tag}>'
    )


def next_free_id(existing_ids, prefix="rId"):
    n = 1
    while f"{prefix}{n}" in existing_ids:
        n += 1
    return f"{prefix}{n}"


def process_docx(src_path, header_img, footer_img, out_path, work_root):
    base = os.path.splitext(os.path.basename(src_path))[0]
    unpack_dir = os.path.join(work_root, "_unpack_" + base)
    if os.path.exists(unpack_dir):
        shutil.rmtree(unpack_dir)
    os.makedirs(unpack_dir)

    with zipfile.ZipFile(src_path, "r") as z:
        z.extractall(unpack_dir)

    word_dir = os.path.join(unpack_dir, "word")
    media_dir = os.path.join(word_dir, "media")
    os.makedirs(media_dir, exist_ok=True)

    header_ext = os.path.splitext(header_img)[1].lstrip(".").lower()
    footer_ext = os.path.splitext(footer_img)[1].lstrip(".").lower()
    if header_ext == "jpeg":
        header_ext = "jpeg"
    if footer_ext == "jpeg":
        footer_ext = "jpeg"

    # unique media filenames
    def unique_media_name(ext):
        i = 1
        while True:
            name = f"hf_inject_{i}.{ext}"
            if not os.path.exists(os.path.join(media_dir, name)):
                return name
            i += 1

    header_media_name = unique_media_name(header_ext)
    footer_media_name = unique_media_name(footer_ext)
    shutil.copy(header_img, os.path.join(media_dir, header_media_name))
    shutil.copy(footer_img, os.path.join(media_dir, footer_media_name))

    # --- header1.xml / footer1.xml ---
    header_rel_id_local = "rId1"
    footer_rel_id_local = "rId1"
    header_xml = build_part_xml("hdr", "Header", HEADER_CX, HEADER_CY, 9001, header_media_name, header_rel_id_local)
    footer_xml = build_part_xml("ftr", "Footer", FOOTER_CX, FOOTER_CY, 9002, footer_media_name, footer_rel_id_local)

    with open(os.path.join(word_dir, "header1.xml"), "w", encoding="utf-8") as f:
        f.write(header_xml)
    with open(os.path.join(word_dir, "footer1.xml"), "w", encoding="utf-8") as f:
        f.write(footer_xml)

    rels_dir = os.path.join(word_dir, "_rels")
    os.makedirs(rels_dir, exist_ok=True)

    def write_part_rels(path, target):
        content = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{target}"/>'
            '</Relationships>'
        )
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)

    write_part_rels(os.path.join(rels_dir, "header1.xml.rels"), header_media_name)
    write_part_rels(os.path.join(rels_dir, "footer1.xml.rels"), footer_media_name)

    # --- document.xml.rels: add header/footer relationships ---
    doc_rels_path = os.path.join(rels_dir, "document.xml.rels")
    with open(doc_rels_path, "r", encoding="utf-8") as f:
        doc_rels = f.read()

    existing_ids = set(re.findall(r'Id="(rId\d+)"', doc_rels))

    # Remove any existing header/footer relationships (we're replacing them)
    doc_rels = re.sub(
        r'<Relationship [^>]*Type="[^"]*/relationships/(header|footer)"[^>]*/>',
        "", doc_rels
    )

    header_rid = next_free_id(existing_ids)
    existing_ids.add(header_rid)
    footer_rid = next_free_id(existing_ids)
    existing_ids.add(footer_rid)

    new_rels = (
        f'<Relationship Id="{header_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        f'<Relationship Id="{footer_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    )
    doc_rels = doc_rels.replace("</Relationships>", new_rels + "</Relationships>")
    with open(doc_rels_path, "w", encoding="utf-8") as f:
        f.write(doc_rels)

    # --- [Content_Types].xml ---
    ct_path = os.path.join(unpack_dir, "[Content_Types].xml")
    with open(ct_path, "r", encoding="utf-8") as f:
        ct = f.read()

    def ensure_default_ext(ct, ext, content_type):
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace(
                "<Default ",
                f'<Default Extension="{ext}" ContentType="{content_type}"/><Default ',
                1
            )
        return ct

    img_ct = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "gif": "image/gif"}
    ct = ensure_default_ext(ct, header_ext, img_ct.get(header_ext, "image/jpeg"))
    if footer_ext != header_ext:
        ct = ensure_default_ext(ct, footer_ext, img_ct.get(footer_ext, "image/jpeg"))

    if "/word/header1.xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
        )
    if "/word/footer1.xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
        )
    with open(ct_path, "w", encoding="utf-8") as f:
        f.write(ct)

    # --- document.xml: wire headerReference/footerReference into every sectPr ---
    doc_xml_path = os.path.join(word_dir, "document.xml")
    with open(doc_xml_path, "r", encoding="utf-8") as f:
        doc_xml = f.read()

    def fix_sectpr(match):
        full = match.group(0)
        # strip any existing header/footer references
        inner = re.sub(r'<w:(headerReference|footerReference)[^/]*/>', "", full)
        refs = (
            f'<w:headerReference w:type="default" r:id="{header_rid}"/>'
            f'<w:footerReference w:type="default" r:id="{footer_rid}"/>'
        )
        # insert refs right after the opening <w:sectPr ...> tag
        inner = re.sub(r'(<w:sectPr[^>]*>)', r'\1' + refs, inner, count=1)
        return inner

    doc_xml, n = re.subn(r'<w:sectPr[^>]*>.*?</w:sectPr>', fix_sectpr, doc_xml, flags=re.DOTALL)
    if n == 0:
        print(f"  WARNING: no <w:sectPr> found in {base} — header/footer not wired!")

    with open(doc_xml_path, "w", encoding="utf-8") as f:
        f.write(doc_xml)

    # --- rezip ---
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(unpack_dir):
            for file in files:
                full = os.path.join(root, file)
                arcname = os.path.relpath(full, unpack_dir)
                zf.write(full, arcname)

    shutil.rmtree(unpack_dir)


def main():
    input_dir, header_img, footer_img, output_dir = sys.argv[1:5]
    os.makedirs(output_dir, exist_ok=True)
    work_root = "/home/claude/work/_batch_tmp"
    os.makedirs(work_root, exist_ok=True)

    files = sorted(
        f for f in os.listdir(input_dir)
        if f.lower().endswith(".docx") and not f.startswith("~$")
    )
    print(f"Found {len(files)} .docx files")
    for fname in files:
        src = os.path.join(input_dir, fname)
        out = os.path.join(output_dir, fname)
        print(f"Processing {fname} ...")
        process_docx(src, header_img, footer_img, out, work_root)
    print("Done.")


if __name__ == "__main__":
    main()
