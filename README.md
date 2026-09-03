# 🖨️ Batch Word Header & Footer Injector

> Stamp a header image and a footer image onto **hundreds of `.docx` files** in one command — no VBA macros, no Word automation, no manual copy-paste.

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![DOCX](https://img.shields.io/badge/File%20Format-.docx-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-brightgreen?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Production%20Ready-success?style=for-the-badge)

---

## 🚀 Why this exists

Manually opening dozens (or hundreds) of Word documents to insert the same header/footer image is a soul-crushing, error-prone chore. This tool treats a `.docx` for what it really is — a ZIP archive of XML — and surgically injects a correctly-sized, correctly-linked header and footer into **every document in a folder**, in seconds.

Built for a real-world use case: applying a clinical laboratory's letterhead (logo header + signatory footer) across an entire catalog of report templates.

---

## ✨ Features

| | |
|---|---|
| 🔁 **Fully batch** | Point it at a folder, get a folder back — 1 file or 1,000 |
| 🖼️ **Any image format** | JPG, JPEG, PNG — header and footer can even differ |
| 🧠 **Smart & non-destructive** | Overwrites existing headers/footers cleanly, or adds new ones if none exist |
| 🔒 **Safe reruns** | Rerun anytime — always produces a consistent result (idempotent) |
| 🧹 **Auto-skips junk** | Ignores Word lock files (`~$file.docx`) automatically |
| 🧾 **Content untouched** | Only headers/footers are modified — body text, tables, formatting stay exactly as-is |
| ⚡ **Zero dependencies** | Pure Python standard library — no `python-docx`, no MS Word/COM required |
| ✅ **Multi-section aware** | Wires header/footer references into *every* section of a document, not just the first |

---

## 🛠️ Requirements

- Python 3.8 or later
- That's it. No pip installs, no Word installation, no macros.

---

## 📦 Usage

```bash
python3 apply_header_footer.py <input_folder> <header_image> <footer_image> <output_folder>
```

### Example

```bash
python3 apply_header_footer.py \
    "Tests Catalog" \
    "Lab_Header_UpScaled.jpeg" \
    "FOOTER_new-Picsart-AiImageEnhancer.jpg" \
    "output"
```

This will:
1. Scan `Tests Catalog/` for every `.docx` file
2. Inject `Lab_Header_UpScaled.jpeg` as the header and `FOOTER_new-Picsart-AiImageEnhancer.jpg` as the footer into each one
3. Write the finished documents into `output/`, preserving original filenames

---

## ⚙️ How it works (under the hood)

A `.docx` file is just a ZIP archive of XML parts. For each document, the script:

1. **Unzips** the file into a temporary working directory
2. **Copies** the header/footer images into `word/media/`
3. **Generates** `header1.xml` / `footer1.xml` with a correctly-scaled inline picture
4. **Wires up relationships** — `document.xml.rels`, `header1.xml.rels`, `footer1.xml.rels`
5. **Registers content types** in `[Content_Types].xml`
6. **Patches every `<w:sectPr>`** in `document.xml` so the header/footer applies to all sections
7. **Re-zips** everything into a valid `.docx`

Every output file is verified for ZIP integrity before being handed back.

---

## 📁 Project Structure

```
.
├── apply_header_footer.py   # The batch script
└── README.md                 # You are here
```

---

## 🧪 Tested On

- Multi-page clinical/lab report templates
- Documents with existing headers/footers (cleanly overwritten)
- Documents with no headers/footers at all
- Batches of 80+ files in a single run

---

## 🗺️ Roadmap Ideas

- [ ] CLI flags for custom header/footer sizing
- [ ] Support for first-page-different / even-page headers
- [ ] Drag-and-drop GUI wrapper
- [ ] `.doc` (legacy) auto-conversion support

Contributions and PRs welcome!

---

## 👨‍💻 Developer

**Shah Mubarak Zaib**
Final-year BS Computer Science student at Islamia College Peshawar, focused on AI/ML engineering.

Builds practical tools that sit at the intersection of automation, clinical software, and machine learning — this script grew out of real production needs while developing lab report software for a clinical laboratory.

- 💼 Open to AI/ML Engineer opportunities
- 🧬 Interests: Machine Learning, Clinical Software Development, Technical Writing
- 🔗 GitHub:*https://github.com/zaibutman*
- 🔗 LinkedIn:*www.linkedin.com/in/shahmubarakzaib*

If this tool saved you time, a ⭐ on the repo goes a long way.

---

## 📄 License

MIT — free to use, modify, and distribute.
