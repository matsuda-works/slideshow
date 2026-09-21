import argparse
import os
import zipfile
import tempfile
import datetime
import struct
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from lxml import etree


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
VIDEO_EXTS = {".mp4", ".mov", ".avi", ".wmv", ".m4v", ".mpg", ".mpeg"}


def is_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTS


def is_video(path: Path) -> bool:
    return path.suffix.lower() in VIDEO_EXTS


def get_video_duration(path: Path) -> float:
    """
    Parses an MP4 or MOV file to extract its duration in seconds.
    Uses pure Python to parse the ISO Base Media File Format.
    """
    try:
        with open(path, "rb") as f:
            f.seek(0, 2)
            file_size = f.tell()
            f.seek(0)
            
            def parse_boxes(start: int, end: int):
                f.seek(start)
                offset = start
                while offset < end:
                    header = f.read(8)
                    if len(header) < 8:
                        break
                    box_size, box_type = struct.unpack(">I4s", header)
                    if box_size == 1:
                        large_size = f.read(8)
                        if len(large_size) < 8:
                            break
                        box_size = struct.unpack(">Q", large_size)[0]
                        box_data_offset = 16
                    elif box_size == 0:
                        box_size = file_size - offset
                        box_data_offset = 8
                    else:
                        box_data_offset = 8
                    
                    if box_type == b"moov":
                        dur = parse_boxes(offset + box_data_offset, offset + box_size)
                        if dur is not None:
                            return dur
                    elif box_type == b"mvhd":
                        f.seek(offset + box_data_offset)
                        version = struct.unpack("B", f.read(1))[0]
                        f.seek(3, 1)  # Skip flags (3 bytes)
                        if version == 1:
                            f.seek(16, 1)  # Skip creation/modification times (16 bytes)
                            timescale = struct.unpack(">I", f.read(4))[0]
                            duration = struct.unpack(">Q", f.read(8))[0]
                        else:
                            f.seek(8, 1)  # Skip creation/modification times (8 bytes)
                            timescale = struct.unpack(">I", f.read(4))[0]
                            duration = struct.unpack(">I", f.read(4))[0]
                        if timescale > 0:
                            return duration / timescale
                    
                    offset += box_size
                    f.seek(offset)
                return None
            
            duration = parse_boxes(0, file_size)
            if duration is not None:
                return duration
    except Exception:
        pass
    return 0.0


def get_shooting_datetime(path: Path) -> datetime.datetime:
    if is_image(path):
        try:
            with Image.open(path) as img:
                exif = img.getexif()
                if exif:
                    # 34665 is Exif Offset for Exif IFD
                    exif_ifd = exif.get_ifd(34665)
                    if exif_ifd:
                        dt_orig = exif_ifd.get(36867)  # DateTimeOriginal
                        if dt_orig:
                            try:
                                return datetime.datetime.strptime(dt_orig.strip(), "%Y:%m:%d %H:%M:%S")
                            except ValueError:
                                pass
                    dt = exif.get(306)  # DateTime
                    if dt:
                        try:
                            return datetime.datetime.strptime(dt.strip(), "%Y:%m:%d %H:%M:%S")
                        except ValueError:
                            pass
        except Exception:
            pass

    try:
        return datetime.datetime.fromtimestamp(os.path.getmtime(path))
    except Exception:
        return datetime.datetime.min


def set_slide_auto_advance(slide, duration_sec: float):
    """
    Sets a slide to advance automatically after a specified number of seconds.
    Inserts <p:transition advTm="..."/> in the correct schema order.
    """
    duration_ms = int(duration_sec * 1000)
    transition_xml = (
        f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'advTm="{duration_ms}"/>'
    )
    transition_element = parse_xml(transition_xml)
    
    children = list(slide.element)
    clr_map_ovr_idx = -1
    c_sld_idx = -1
    
    for i, child in enumerate(children):
        tag = child.tag
        if tag.endswith('clrMapOvr'):
            clr_map_ovr_idx = i
        elif tag.endswith('cSld'):
            c_sld_idx = i
            
    if clr_map_ovr_idx != -1:
        slide.element.insert(clr_map_ovr_idx + 1, transition_element)
    elif c_sld_idx != -1:
        slide.element.insert(c_sld_idx + 1, transition_element)
    else:
        slide.element.append(transition_element)


def enable_presentation_loop(prs):
    """
    Modifies ppt/presProps.xml to set loop="true" under p:showPr,
    enabling "Loop continuously until 'Esc'" programmatically.
    """
    pres_props_rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
    try:
        part = prs.part.part_related_by(pres_props_rel)
    except KeyError:
        return
        
    root = etree.fromstring(part.blob)
    namespaces = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    
    show_pr_elements = root.xpath('.//p:showPr', namespaces=namespaces)
    if show_pr_elements:
        show_pr = show_pr_elements[0]
        show_pr.set('loop', 'true')
    else:
        show_pr = etree.Element('{http://schemas.openxmlformats.org/presentationml/2006/main}showPr', loop='true')
        ext_lst_elements = root.xpath('./p:extLst', namespaces=namespaces)
        if ext_lst_elements:
            ext_lst = ext_lst_elements[0]
            ext_lst.addprevious(show_pr)
        else:
            root.append(show_pr)
            
    part.blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def collect_media_files(folder: Path):
    files = []
    for p in folder.rglob("*"):
        if p.is_file() and (is_image(p) or is_video(p)):
            files.append(p)
    return sorted(files, key=lambda x: (get_shooting_datetime(x), x.name.lower()))


def set_white_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)


def add_centered_text(slide, text, prs, font_size=36):
    left = Inches(0.8)
    top = Inches(2.8)
    width = prs.slide_width - Inches(1.6)
    height = Inches(1.2)

    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = True
    run.font.color.rgb = RGBColor(30, 30, 30)

    p.alignment = 1  # center


def add_caption(slide, text, prs):
    left = Inches(0.2)
    top = prs.slide_height - Inches(0.45)
    width = prs.slide_width - Inches(0.4)
    height = Inches(0.3)

    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(100, 100, 100)

    p.alignment = 1  # center


def prepare_image_for_ppt(image_path: Path, temp_dir: Path):
    """
    EXIF回転を反映し、PowerPointで扱いやすいJPEGに変換する。
    """
    img = Image.open(image_path)
    img = ImageOps.exif_transpose(img)

    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    elif img.mode == "L":
        img = img.convert("RGB")

    out_path = temp_dir / f"{image_path.stem}_converted.jpg"
    img.save(out_path, "JPEG", quality=92)
    return out_path


def add_image_slide(prs, image_path: Path, temp_dir: Path, show_filename=True, duration=3.0):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_white_background(slide)
    set_slide_auto_advance(slide, duration)

    converted = prepare_image_for_ppt(image_path, temp_dir)

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    with Image.open(converted) as img:
        img_w, img_h = img.size

    img_ratio = img_w / img_h
    slide_ratio = slide_w / slide_h

    if img_ratio > slide_ratio:
        # 横長：幅いっぱい
        width = slide_w
        height = int(slide_w / img_ratio)
        left = 0
        top = int((slide_h - height) / 2)
    else:
        # 縦長：高さいっぱい
        height = slide_h
        width = int(slide_h * img_ratio)
        left = int((slide_w - width) / 2)
        top = 0

    slide.shapes.add_picture(str(converted), left, top, width, height)

    if show_filename:
        add_caption(slide, image_path.name, prs)


def add_video_slide(prs, video_path: Path, show_filename=True, duration=0.0):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_white_background(slide)
    set_slide_auto_advance(slide, duration)

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    margin_x = Inches(0.7)
    margin_y = Inches(0.8)

    left = margin_x
    top = margin_y
    width = slide_w - margin_x * 2
    height = slide_h - margin_y * 2

    # 動画を埋め込み
    # PowerPoint側の環境によって再生可否は動画形式に依存します。
    try:
        slide.shapes.add_movie(
            str(video_path),
            left,
            top,
            width,
            height,
            mime_type="video/mp4"
        )
        
        # XMLを書き換えて動画を自動再生（autoplay）に設定する
        cond_elements = slide.element.xpath('.//p:video//p:cond')
        for cond in cond_elements:
            if cond.get('delay') == 'indefinite':
                cond.set('delay', '0')
    except Exception:
        # 埋め込みに失敗した場合は、代わりにファイル名表示とリンク風スライドにする
        add_centered_text(slide, f"動画ファイル\n{video_path.name}", prs, font_size=28)

    if show_filename:
        add_caption(slide, video_path.name, prs)


def extract_zip_if_needed(input_path: Path):
    """
    ZIPなら一時フォルダに展開して、そのフォルダを返す。
    フォルダならそのまま返す。
    """
    if input_path.is_dir():
        return input_path, None

    if input_path.is_file() and input_path.suffix.lower() == ".zip":
        temp_dir = tempfile.TemporaryDirectory()
        extract_dir = Path(temp_dir.name) / "extracted"
        extract_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as z:
            z.extractall(extract_dir)

        return extract_dir, temp_dir

    raise ValueError("入力には、ZIPファイルまたはフォルダを指定してください。")


def create_slideshow(input_path: Path, output_path: Path, title: str, show_filename=True, duration=3.0):
    media_folder, temp_zip_dir = extract_zip_if_needed(input_path)

    with tempfile.TemporaryDirectory() as work_dir:
        work_dir = Path(work_dir)

        media_files = collect_media_files(media_folder)

        if not media_files:
            raise RuntimeError("写真または動画ファイルが見つかりませんでした。")

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        # 表紙
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_white_background(slide)
        add_centered_text(slide, title, prs, font_size=40)
        set_slide_auto_advance(slide, duration)

        if show_filename:
            subtitle_box = slide.shapes.add_textbox(
                Inches(1),
                Inches(4.0),
                prs.slide_width - Inches(2),
                Inches(0.6)
            )
            tf = subtitle_box.text_frame
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = f"{len(media_files)} files"
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(120, 120, 120)
            p.alignment = 1

        # メディアスライド
        for file_path in media_files:
            if is_image(file_path):
                add_image_slide(prs, file_path, work_dir, show_filename=show_filename, duration=duration)
            elif is_video(file_path):
                video_duration = get_video_duration(file_path)
                if video_duration <= 0.0:
                    video_duration = 10.0  # Fallback duration
                add_video_slide(prs, file_path, show_filename=show_filename, duration=video_duration + 0.5)

        # 終了スライド
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_white_background(slide)
        add_centered_text(slide, "End", prs, font_size=44)
        set_slide_auto_advance(slide, duration)

        enable_presentation_loop(prs)
        prs.save(output_path)

    if temp_zip_dir is not None:
        temp_zip_dir.cleanup()


def main():
    parser = argparse.ArgumentParser(
        description="写真と動画からPowerPointスライドショーを作成します。"
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(Path(__file__).parent / "photo"),
        help="写真・動画フォルダ、またはZIPファイルのパス。既定値: main.pyと同じフォルダのphoto"
    )
    parser.add_argument(
        "-o",
        "--output",
        default="slideshow.pptx",
        help="出力するPowerPointファイル名。既定値: slideshow.pptx"
    )
    parser.add_argument(
        "-t",
        "--title",
        default="フォトスライドショー",
        help="表紙タイトル"
    )
    parser.add_argument(
        "--no-filename",
        action="store_true",
        help="スライド下部にファイル名を表示しない"
    )
    parser.add_argument(
        "-d",
        "--duration",
        type=float,
        default=3.0,
        help="スライドごとの表示時間（秒）。既定値: 3.0"
    )

    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    try:
        create_slideshow(
            input_path=input_path,
            output_path=output_path,
            title=args.title,
            show_filename=not args.no_filename,
            duration=args.duration
        )
        print("PowerPointを作成しました:")
        print(output_path)
    except PermissionError:
        print(f"\n【エラー】出力ファイル '{output_path.name}' が他のプログラム（PowerPointなど）で開かれているため、上書き保存できません。")
        print("PowerPointでこのファイルを閉じてから、再度実行してください。\n")


if __name__ == "__main__":
    main()