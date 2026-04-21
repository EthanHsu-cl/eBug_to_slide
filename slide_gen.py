import copy
from io import BytesIO
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

from parser import BugData, ImageStep

TEMPLATE_PATH = Path(__file__).parent / "UX ppt template.pptx"

# Exact positions cloned from the template (in inches)
TITLE_LEFT, TITLE_TOP = 0.548, 0.333
TITLE_W, TITLE_H = 12.644, 0.687

TYPE_LEFT, TYPE_TOP = 0.548, 0.869
TYPE_W, TYPE_H = 9.654, 0.645

VERSION_LEFT, VERSION_TOP = 11.610, 0.131
VERSION_W, VERSION_H = 1.606, 0.404

IMG_LEFT, IMG_TOP = 0.548, 1.470
IMG_W, IMG_H = 12.237, 5.722

ANNOT_LEFT, ANNOT_TOP = 13.458, 3.097
ANNOT_W, ANNOT_H = 2.542, 0.505


def _in(value: float) -> Emu:
    return Inches(value)


def _find_layout(prs: Presentation, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return prs.slide_layouts[0]


def _clone_slide_shapes(src_slide, dst_slide) -> None:
    """Deep-copy all shapes from src_slide into dst_slide's shape tree."""
    sp_tree = dst_slide.shapes._spTree
    for el in src_slide.shapes._spTree:
        sp_tree.append(copy.deepcopy(el))


def _remove_template_slides(prs: Presentation, count: int) -> None:
    """Remove the first *count* slides (the original template slides)."""
    sld_id_lst = prs.slides._sldIdLst
    for _ in range(count):
        sld_id_lst.remove(sld_id_lst[0])


def _set_shape_text(slide, shape_name: str, text: str) -> None:
    """Set text on the first shape whose name matches shape_name."""
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            # Preserve paragraph formatting; just replace run text
            for para in tf.paragraphs:
                for run in para.runs:
                    run.text = ""
            if tf.paragraphs:
                tf.paragraphs[0].runs[0].text = text if tf.paragraphs[0].runs else None
                if not tf.paragraphs[0].runs:
                    tf.paragraphs[0].text = text
            return


def _replace_image_placeholder(slide, image_data: bytes) -> None:
    """
    Remove the large main-area text box (the one whose text starts with '<image')
    and insert the actual image, scaled to fit the same area with aspect ratio preserved.
    """
    # Find and remove the placeholder text box
    placeholder_shape = None
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip().startswith("<image"):
            placeholder_shape = shape
            break

    if placeholder_shape is not None:
        sp = placeholder_shape.element
        sp.getparent().remove(sp)

    if not image_data:
        return

    # Calculate scaled dimensions preserving aspect ratio
    img = Image.open(BytesIO(image_data))
    img_w_px, img_h_px = img.size

    area_w = _in(IMG_W)
    area_h = _in(IMG_H)

    scale = min(area_w / img_w_px, area_h / img_h_px)
    final_w = int(img_w_px * scale)
    final_h = int(img_h_px * scale)

    # Center within the image area
    left = _in(IMG_LEFT) + (area_w - final_w) // 2
    top = _in(IMG_TOP) + (area_h - final_h) // 2

    slide.shapes.add_picture(BytesIO(image_data), left, top, final_w, final_h)


def _update_slide(slide, step: ImageStep, title: str, version_info: str) -> None:
    """Populate a cloned slide with data from one ImageStep."""
    # Title placeholder (shape named '標題 2')
    for shape in slide.shapes:
        if shape.shape_type == 1 and hasattr(shape, "placeholder_format"):  # PLACEHOLDER
            if shape.placeholder_format and shape.placeholder_format.idx == 0:
                shape.text_frame.paragraphs[0].runs[0].text = title
                break
    # Fallback: find any shape with "Short Description" text
    for shape in slide.shapes:
        if shape.has_text_frame and "Short Description" in shape.text_frame.text:
            tf = shape.text_frame
            if tf.paragraphs and tf.paragraphs[0].runs:
                tf.paragraphs[0].runs[0].text = title
            break

    # Type label ("Current: <version>")
    type_label = f"{step.img_type}: {version_info}"
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip().startswith(
            ("Current:", "Reference:", "Proposal:")
        ):
            tf = shape.text_frame
            if tf.paragraphs and tf.paragraphs[0].runs:
                tf.paragraphs[0].runs[0].text = type_label
            break

    # Version info box ("PX, fix in X/X")
    for shape in slide.shapes:
        if shape.has_text_frame and "fix in" in shape.text_frame.text.lower():
            tf = shape.text_frame
            if tf.paragraphs and tf.paragraphs[0].runs:
                tf.paragraphs[0].runs[0].text = version_info
            break

    # Annotation / step description box (the off-slide-right one)
    for shape in slide.shapes:
        if shape.has_text_frame and shape.left > _in(13.0):
            tf = shape.text_frame
            if tf.paragraphs and tf.paragraphs[0].runs:
                tf.paragraphs[0].runs[0].text = step.text
            elif tf.paragraphs:
                tf.paragraphs[0].text = step.text
            break

    # Replace image placeholder with the actual image
    _replace_image_placeholder(slide, step.image_data)


def generate(bug_data: BugData, output_path: str) -> None:
    """
    Generate a PPTX at *output_path* based on the template and *bug_data*.
    One slide is created per ImageStep.
    """
    prs = Presentation(str(TEMPLATE_PATH))
    n_template_slides = len(prs.slides)

    layout = _find_layout(prs, "1_只有標題 (no BK)")
    template_slide = prs.slides[0]  # slide 1 is the "Current with annotation" style

    for step in bug_data.image_steps:
        new_slide = prs.slides.add_slide(layout)
        _clone_slide_shapes(template_slide, new_slide)
        _update_slide(new_slide, step, bug_data.title, bug_data.version_info)

    _remove_template_slides(prs, n_template_slides)
    prs.save(output_path)
    print(f"Saved: {output_path}  ({len(bug_data.image_steps)} slide(s))")
