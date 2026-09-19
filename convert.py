#!/usr/bin/env python3
"""
Converter Core - Main Entry Point
Lightweight CLI runner coordinating multi-format document conversion.
"""

from core.bootstrap import FILES_DIR, DOCS_DIR, IMGS_DIR
from core.i18n import t, toggle_language, get_language
from core.engines import get_converter_options

CATEGORIES = [
    ("cat_documents", DOCS_DIR, {".pdf", ".docx", ".doc"}),
    ("cat_images", IMGS_DIR, {".png", ".jpg", ".jpeg"}),
]

def select_category(categories):
    while True:
        print(f"\n{t('select_category_title')}")
        for idx, (cat_key, cat_dir, valid_exts) in enumerate(categories, 1):
            count = len([
                f for f in cat_dir.iterdir()
                if f.is_file() and not f.name.startswith("~$") and f.name != ".gitkeep" and f.suffix.lower() in valid_exts
            ]) if cat_dir.exists() else 0
            print(f"[{idx}] {t(cat_key)} ({count} files)")

        prompt = t("select_category", count=len(categories))
        choice = input(prompt).strip().lower()

        if choice == "q":
            return None
        if choice == "l":
            new_lang = toggle_language()
            print(t("language_switched", lang=new_lang.upper()))
            continue
        if choice == "":
            return categories[0]
        if choice.isdigit() and 1 <= int(choice) <= len(categories):
            return categories[int(choice) - 1]

        print(t("invalid_choice", count=len(categories)))

def select_file(files):
    while True:
        prompt = t("select_file", count=len(files))
        choice = input(prompt).strip().lower()

        if choice == "q":
            return None
        if choice == "l":
            new_lang = toggle_language()
            print(t("language_switched", lang=new_lang.upper()))
            continue
        if choice == "":
            return files[0]
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice) - 1]

        print(t("invalid_choice", count=len(files)))

def select_mode(options):
    print(t("select_mode_title"))
    for idx, (label_key, _, _) in enumerate(options, 1):
        print(f"[{idx}] {t(label_key)}")

    while True:
        choice = input(t("select_mode", count=len(options))).strip().lower()
        if choice == "q":
            return None
        if choice == "":
            return options[0]
        if choice.isdigit() and 1 <= int(choice) <= len(options):
            return options[int(choice) - 1]

        print(t("invalid_choice", count=len(options)))

def main():
    selected_cat = select_category(CATEGORIES)
    if not selected_cat:
        return

    cat_key, cat_dir, valid_exts = selected_cat
    files = [
        f for f in sorted(cat_dir.iterdir())
        if f.is_file() and not f.name.startswith("~$") and f.name != ".gitkeep" and f.suffix.lower() in valid_exts
    ] if cat_dir.exists() else []

    if not files:
        print(f"\n{t('no_files_in_category', dir=cat_dir.name)}")
        return

    print(f"\n{t('files_in_category', dir=cat_dir.name)}")
    for idx, f in enumerate(files, 1):
        print(f"[{idx}] {f.name}")

    src_file = select_file(files)
    if not src_file:
        return

    options = get_converter_options(src_file.suffix)
    if not options:
        print(t("unsupported", ext=src_file.suffix))
        return

    selected_opt = select_mode(options)
    if not selected_opt:
        return

    _, target_ext, convert_fn = selected_opt

    # Clean double extensions (e.g., file.docx.docx)
    base_name = src_file.stem
    if base_name.endswith(".docx") and target_ext == "docx":
        base_name = base_name[:-5]

    dst_file = src_file.parent / f"{base_name}.{target_ext}"
    if target_ext == "txt":
        print(t("extracting_text", src=src_file.name, dst=dst_file.name))
    else:
        print(t("converting", src=src_file.name, dst=dst_file.name))
    convert_fn(src_file, dst_file)
    print(t("done", path=dst_file))

if __name__ == "__main__":
    main()
