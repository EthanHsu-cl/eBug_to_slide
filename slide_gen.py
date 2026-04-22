import copy
from io import BytesIO
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Emu, Pt

from parser import BugData, ImageStep

TEMPLATE_PATH = Path(__file__).parent / "Template" / "New Layout.pptx"

# Image area position and size (inches) — from New Layout.pptx inspection
IMG_LEFT, IMG_TOP = 1.2467, 1.6255
IMG_W,    IMG_H   = 10.8399, 5.7217

_TYPE_COLORS = {
    "Current":   RGBColor(0xC0, 0x00, 0x00),
    "Reference": RGBColor(0xFF, 0xC0, 0x00),
    "Proposal":  RGBColor(0xFF, 0x66, 0x00),
}


def _in(value: float) -> Emu:
    return Inches(value)


def _find_layout(prs: Presentation, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return prs.slide_layouts[0]


def _clone_slide_shapes(src_slide, dst_slide) -> None:
    """Replace dst_slide's shapes with deep copies of src_slide's shapes."""
    sp_tree = dst_slide.shapes._spTree
    # spTree layout: [0]=nvGrpSpPr, [1]=grpSpPr, [2+]=shapes from layout.
    # Remove all shapes add_slide() already added, keep only the two metadata nodes.
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)
    # Append shapes from template slide (skip its own metadata nodes).
    for el in list(src_slide.shapes._spTree)[2:]:
        sp_tree.append(copy.deepcopy(el))


def _remove_template_slides(prs: Presentation, count: int) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for _ in range(count):
        sld_id_lst.remove(sld_id_lst[0])


def _safe_set_text(tf, text: str) -> None:
    """Set the text of the first paragraph, preserving existing run formatting."""
    p = tf.paragraphs[0]
    if p.runs:
        for run in p.runs[1:]:
            run.text = ""
        p.runs[0].text = text
    else:
        p.text = text


def _set_type_label(tf, text: str, color: RGBColor) -> None:
    """Set type label text and apply the correct type color to the run."""
    p = tf.paragraphs[0]
    if p.runs:
        for run in p.runs[1:]:
            run.text = ""
        p.runs[0].text = text
        p.runs[0].font.color.rgb = color
    else:
        p.text = text
        if p.runs:
            p.runs[0].font.color.rgb = color


def _insert_image(slide, image_data: bytes) -> None:
    """Scale image to fit the main image area and add it to the slide."""
    if not image_data:
        return
    img = Image.open(BytesIO(image_data))
    img_w_px, img_h_px = img.size

    area_w = _in(IMG_W)
    area_h = _in(IMG_H)
    scale = min(area_w / img_w_px, area_h / img_h_px)
    final_w = int(img_w_px * scale)
    final_h = int(img_h_px * scale)

    left = _in(IMG_LEFT) + (area_w - final_w) // 2
    top  = _in(IMG_TOP)  + (area_h - final_h) // 2

    slide.shapes.add_picture(BytesIO(image_data), left, top, final_w, final_h)


def _update_slide(
    slide,
    step: ImageStep,
    title: str,
    version_info: str,
    section_texts: dict[str, str],
) -> None:
    """Populate all text fields and insert the image on a cloned slide."""
    # Remove any existing picture shapes from the template (replaced by actual image)
    for shape in list(slide.shapes):
        if int(shape.shape_type) == 13:  # MSO_SHAPE_TYPE.PICTURE
            shape.element.getparent().remove(shape.element)

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        name = shape.name

        if name == "標題 2":
            _safe_set_text(shape.text_frame, title)

        elif name == "Title 1" and shape.top < _in(1.3):
            # Type label: section description with correct color
            section_text = section_texts.get(step.img_type, "")
            label = f"{step.img_type}: {section_text}" if section_text else f"{step.img_type}:"
            color = _TYPE_COLORS.get(step.img_type, RGBColor(0, 0, 0))
            _set_type_label(shape.text_frame, label, color)

        elif name == "文字方塊 4":
            shape.top = Emu(0)
            _safe_set_text(shape.text_frame, "PX, fix in X/X")
            p = shape.text_frame.paragraphs[0]
            for run in p.runs:
                run.font.size = Pt(16)

        elif name == "文字方塊 9":
            # Annotation box (intentionally off right edge of slide)
            _safe_set_text(shape.text_frame, step.text)

    _insert_image(slide, step.image_data)


def generate(bug_data: BugData, output_path: str) -> None:
    """Generate a PPTX at output_path — one slide per ImageStep."""
    prs = Presentation(str(TEMPLATE_PATH))
    n_template_slides = len(prs.slides)

    layout = _find_layout(prs, "1_只有標題 (no BK)")
    template_slide = prs.slides[0]

    for step in bug_data.image_steps:
        new_slide = prs.slides.add_slide(layout)
        _clone_slide_shapes(template_slide, new_slide)
        _update_slide(new_slide, step, bug_data.title, bug_data.version_info, bug_data.section_texts)

    _remove_template_slides(prs, n_template_slides)
    prs.save(output_path)
    print(f"Saved: {output_path}  ({len(bug_data.image_steps)} slide(s))")
