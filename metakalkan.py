#!/usr/bin/env python3
"""Pardus MetaKalkan - yerel belge gizlilik denetçisi."""

import os
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set
from urllib.parse import unquote, urlparse

try:
    import gi
    gi.require_version("Gtk", "3.0")
    gi.require_version("Gdk", "3.0")
    from gi.repository import Gdk, GLib, Gtk
except (ImportError, ValueError) as exc:
    print("GTK 3 / PyGObject bulunamadı. README'deki sistem bağımlılıklarını kurun.", file=sys.stderr)
    raise SystemExit(1) from exc

from lxml import etree
from PIL import ExifTags, Image
import pikepdf


APP_NAME = "Pardus MetaKalkan"
SUPPORTED = {".pdf", ".docx", ".xlsx", ".jpg", ".jpeg", ".tif", ".tiff", ".webp", ".png"}
OOXML_NS = {"cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass
class Finding:
    key: str
    category: str
    title: str
    value: str
    selected: bool = True


def compact_value(value: object, limit: int = 110) -> str:
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


def output_path(source: Path) -> Path:
    candidate = source.with_name(f"{source.stem}_temizlenmis{source.suffix}")
    number = 2
    while candidate.exists():
        candidate = source.with_name(f"{source.stem}_temizlenmis_{number}{source.suffix}")
        number += 1
    return candidate


class Inspector:
    """Desteklenen biçimleri analiz eder ve seçili bulguları güvenli kopyada siler."""

    def inspect(self, path: Path) -> List[Finding]:
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            return self._inspect_pdf(path)
        if suffix in {".docx", ".xlsx"}:
            return self._inspect_ooxml(path)
        if suffix in {".jpg", ".jpeg", ".tif", ".tiff", ".webp", ".png"}:
            return self._inspect_image(path)
        raise ValueError("Desteklenmeyen dosya türü")

    def clean(self, source: Path, findings: Iterable[Finding], destination: Path) -> None:
        selected = {item.key for item in findings if item.selected}
        if source.suffix.lower() == ".pdf":
            self._clean_pdf(source, destination, selected)
        elif source.suffix.lower() in {".docx", ".xlsx"}:
            self._clean_ooxml(source, destination, selected)
        else:
            self._clean_image(source, destination, selected)

    def _inspect_pdf(self, path: Path) -> List[Finding]:
        result: List[Finding] = []
        with pikepdf.open(path) as pdf:
            info = pdf.docinfo
            for name, value in info.items():
                key = f"pdf:info:{str(name)}"
                result.append(Finding(key, "PDF meta verisi", str(name).lstrip("/"), compact_value(value)))
            if "/Metadata" in pdf.Root:
                result.append(Finding("pdf:xmp", "PDF meta verisi", "XMP meta veri paketi", "Belge içinde gömülü"))
            if "/PieceInfo" in pdf.Root:
                result.append(Finding("pdf:pieceinfo", "PDF meta verisi", "PieceInfo", "Uygulama verisi içeriyor"))
            if "/MarkInfo" in pdf.Root:
                result.append(Finding("pdf:markinfo", "PDF meta verisi", "İşaretleme bilgisi", "Belge işaretleme ayarları"))
        return result

    def _clean_pdf(self, source: Path, destination: Path, selected: Set[str]) -> None:
        with pikepdf.open(source) as pdf:
            for key in list(selected):
                if key.startswith("pdf:info:"):
                    name = pikepdf.Name(key.removeprefix("pdf:info:"))
                    if name in pdf.docinfo:
                        del pdf.docinfo[name]
            if "pdf:xmp" in selected and "/Metadata" in pdf.Root:
                del pdf.Root["/Metadata"]
            if "pdf:pieceinfo" in selected and "/PieceInfo" in pdf.Root:
                del pdf.Root["/PieceInfo"]
            if "pdf:markinfo" in selected and "/MarkInfo" in pdf.Root:
                del pdf.Root["/MarkInfo"]
            pdf.save(destination)

    def _inspect_image(self, path: Path) -> List[Finding]:
        result: List[Finding] = []
        with Image.open(path) as image:
            exif = image.getexif()
            for tag_id, value in exif.items():
                name = ExifTags.TAGS.get(tag_id, f"EXIF {tag_id}")
                if tag_id == 34853:
                    gps = exif.get_ifd(tag_id)
                    if gps:
                        result.append(Finding("image:gps", "Konum", "GPS konumu", self._format_gps(gps)))
                    continue
                result.append(Finding(f"image:exif:{tag_id}", "Fotoğraf meta verisi", name, compact_value(value)))
            for field in ("icc_profile", "xmp", "XML:com.adobe.xmp"):
                if field in image.info:
                    result.append(Finding(f"image:info:{field}", "Fotoğraf meta verisi", field, "Gömülü profil/veri"))
        return result

    def _format_gps(self, gps: Dict[int, object]) -> str:
        def decimal(values: object, direction: object) -> Optional[float]:
            try:
                deg, minutes, seconds = values  # type: ignore[misc]
                answer = float(deg) + float(minutes) / 60 + float(seconds) / 3600
                return -answer if str(direction) in {"S", "W"} else answer
            except (TypeError, ValueError, ZeroDivisionError):
                return None
        latitude = decimal(gps.get(2), gps.get(1))
        longitude = decimal(gps.get(4), gps.get(3))
        if latitude is not None and longitude is not None:
            return f"{latitude:.6f}, {longitude:.6f}"
        return "GPS EXIF etiketi mevcut"

    def _clean_image(self, source: Path, destination: Path, selected: Set[str]) -> None:
        # EXIF'i seçili etiketsiz yeniden kurar; piksel verisi değişmeden korunur.
        with Image.open(source) as image:
            image.load()
            original = image.getexif()
            new_exif = Image.Exif()
            for tag_id, value in original.items():
                key = "image:gps" if tag_id == 34853 else f"image:exif:{tag_id}"
                if key not in selected:
                    new_exif[tag_id] = value
            save_args: Dict[str, object] = {}
            if source.suffix.lower() not in {".png"}:
                save_args["exif"] = new_exif.tobytes()
            # Renk profili teknik bir meta veridir; kullanıcı onu seçmedikçe koru.
            if "icc_profile" in image.info and "image:info:icc_profile" not in selected:
                save_args["icc_profile"] = image.info["icc_profile"]
            # WebP, XMP'yi Pillow üzerinden kayıpsız taşıyabilen biçimlerden biridir.
            if image.format == "WEBP" and "xmp" in image.info and "image:info:xmp" not in selected:
                save_args["xmp"] = image.info["xmp"]
            if image.format == "JPEG":
                save_args["quality"] = "keep"
                save_args["subsampling"] = "keep"
            image.save(destination, **save_args)

    def _inspect_ooxml(self, path: Path) -> List[Finding]:
        result: List[Finding] = []
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "docProps/core.xml" in names:
                root = etree.fromstring(archive.read("docProps/core.xml"))
                for elem in root:
                    value = (elem.text or "").strip()
                    if value:
                        local = etree.QName(elem).localname
                        result.append(Finding(f"ooxml:core:{local}", "Belge meta verisi", local, compact_value(value)))
            if "docProps/app.xml" in names:
                root = etree.fromstring(archive.read("docProps/app.xml"))
                for elem in root:
                    value = (elem.text or "").strip()
                    if value and etree.QName(elem).localname in {"Company", "Manager", "LastAuthor", "Application"}:
                        local = etree.QName(elem).localname
                        result.append(Finding(f"ooxml:app:{local}", "Kurum/uygulama bilgisi", local, compact_value(value)))
            if "docProps/custom.xml" in names:
                root = etree.fromstring(archive.read("docProps/custom.xml"))
                for prop in root:
                    name = prop.get("name", "Özel özellik")
                    result.append(Finding(f"ooxml:custom:{name}", "Özel özellik", name, compact_value(" ".join(prop.itertext()))))
            if path.suffix.lower() == ".docx":
                comment_parts = [name for name in names if name.startswith("word/comments") and name.endswith(".xml")]
                if comment_parts:
                    count = sum(len(etree.fromstring(archive.read(name))) for name in comment_parts if name == "word/comments.xml")
                    result.append(Finding("ooxml:comments", "Word gizlilik verisi", "Yorumlar", f"{count} yorum ve ilişkili veri"))
                revisions = self._count_revisions(archive, names)
                if revisions:
                    result.append(Finding("ooxml:revisions", "Word gizlilik verisi", "Revizyon izleri", f"{revisions} izlenen değişiklik"))
            if path.suffix.lower() == ".xlsx":
                comments = [name for name in names if name.startswith("xl/comments") or name.startswith("xl/threadedComments")]
                if comments:
                    result.append(Finding("ooxml:comments", "Excel gizlilik verisi", "Hücre yorumları", f"{len(comments)} yorum bölümü"))
        return result

    def _count_revisions(self, archive: zipfile.ZipFile, names: Set[str]) -> int:
        count = 0
        for name in names:
            if name.startswith("word/") and name.endswith(".xml"):
                try:
                    root = etree.fromstring(archive.read(name))
                    count += len(root.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo", namespaces=OOXML_NS))
                except etree.XMLSyntaxError:
                    pass
        return count

    def _clean_ooxml(self, source: Path, destination: Path, selected: Set[str]) -> None:
        remove_comments = "ooxml:comments" in selected
        remove_revisions = "ooxml:revisions" in selected
        with zipfile.ZipFile(source) as read_archive, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as write_archive:
            names = set(read_archive.namelist())
            skipped = self._ooxml_skipped_parts(names, remove_comments)
            for info in read_archive.infolist():
                name = info.filename
                if name in skipped:
                    continue
                data = read_archive.read(name)
                if name == "docProps/core.xml":
                    data = self._clean_core(data, selected)
                elif name == "docProps/app.xml":
                    data = self._clean_app(data, selected)
                elif name == "docProps/custom.xml":
                    data = self._clean_custom(data, selected)
                elif remove_revisions and name.startswith("word/") and name.endswith(".xml"):
                    data = self._accept_revisions(data)
                elif remove_comments and name.endswith(".rels"):
                    data = self._strip_comment_relationships(data)
                elif remove_comments and name == "[Content_Types].xml":
                    data = self._strip_comment_content_types(data)
                write_archive.writestr(info, data)

    def _ooxml_skipped_parts(self, names: Set[str], remove_comments: bool) -> Set[str]:
        if not remove_comments:
            return set()
        return {name for name in names if (name.startswith("word/comments") or name.startswith("xl/comments")
                or name.startswith("xl/threadedComments") or name.startswith("xl/persons"))}

    def _clean_core(self, data: bytes, selected: Set[str]) -> bytes:
        root = etree.fromstring(data)
        for elem in list(root):
            if f"ooxml:core:{etree.QName(elem).localname}" in selected:
                root.remove(elem)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _clean_app(self, data: bytes, selected: Set[str]) -> bytes:
        root = etree.fromstring(data)
        for elem in list(root):
            if f"ooxml:app:{etree.QName(elem).localname}" in selected:
                root.remove(elem)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _clean_custom(self, data: bytes, selected: Set[str]) -> bytes:
        root = etree.fromstring(data)
        for elem in list(root):
            if f"ooxml:custom:{elem.get('name', 'Özel özellik')}" in selected:
                root.remove(elem)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _strip_comment_relationships(self, data: bytes) -> bytes:
        root = etree.fromstring(data)
        for relation in list(root):
            kind = relation.get("Type", "").lower()
            if "comment" in kind or "person" in kind:
                root.remove(relation)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _strip_comment_content_types(self, data: bytes) -> bytes:
        root = etree.fromstring(data)
        for override in list(root):
            content_type = override.get("ContentType", "").lower()
            if "comment" in content_type or "threadedcomment" in content_type or "person" in content_type:
                root.remove(override)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _accept_revisions(self, data: bytes) -> bytes:
        try:
            root = etree.fromstring(data)
        except etree.XMLSyntaxError:
            return data
        # Ekleme ve taşıma-hedeflerindeki içerik korunur; silinen içerik çıkarılır.
        for elem in root.xpath(".//w:del | .//w:moveFrom", namespaces=OOXML_NS):
            parent = elem.getparent()
            if parent is not None:
                parent.remove(elem)
        for elem in root.xpath(".//w:ins | .//w:moveTo", namespaces=OOXML_NS):
            parent = elem.getparent()
            if parent is not None:
                index = parent.index(elem)
                for child in list(elem):
                    elem.remove(child)
                    parent.insert(index, child)
                    index += 1
                parent.remove(elem)
        settings = root.xpath(".//w:trackRevisions", namespaces=OOXML_NS)
        for elem in settings:
            elem.getparent().remove(elem)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class MetaKalkanWindow(Gtk.ApplicationWindow):
    def __init__(self, app: Gtk.Application):
        super().__init__(application=app, title=APP_NAME)
        self.set_default_size(940, 610)
        self.set_border_width(18)
        self.inspector = Inspector()
        self.source: Optional[Path] = None
        self.findings: List[Finding] = []
        self.store = Gtk.ListStore(bool, str, str, str, str)
        self._build_ui()

    def _build_ui(self) -> None:
        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        self.add(root)
        title = Gtk.Label()
        title.set_markup("<span size='xx-large' weight='bold'>Pardus MetaKalkan</span>\nDosyayı göndermeden önce gizlilik verilerini denetleyin.")
        title.set_xalign(0)
        root.pack_start(title, False, False, 0)
        self.drop_area = Gtk.EventBox()
        self.drop_area.set_visible_window(True)
        self.drop_area.set_size_request(-1, 105)
        self.drop_area.get_style_context().add_class("drop-area")
        drop_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=5, margin_top=18, margin_bottom=18)
        self.drop_label = Gtk.Label(label="Dosyayı buraya sürükleyip bırakın")
        self.drop_label.set_markup("<b>Dosyayı buraya sürükleyip bırakın</b>")
        self.file_label = Gtk.Label(label="veya aşağıdan bir dosya seçin (.pdf, .docx, .xlsx, JPEG/TIFF/WebP/PNG)")
        drop_box.pack_start(self.drop_label, False, False, 0)
        drop_box.pack_start(self.file_label, False, False, 0)
        self.drop_area.add(drop_box)
        self.drop_area.drag_dest_set(Gtk.DestDefaults.ALL, [Gtk.TargetEntry.new("text/uri-list", 0, 0)], Gdk.DragAction.COPY)
        self.drop_area.connect("drag-data-received", self.on_drag_data_received)
        root.pack_start(self.drop_area, False, False, 0)
        actions = Gtk.Box(spacing=8)
        browse = Gtk.Button.new_with_label("Gözat…")
        browse.connect("clicked", self.on_browse)
        self.select_all = Gtk.Button.new_with_label("Tümünü seç")
        self.select_all.connect("clicked", self.on_select_all, True)
        self.select_none = Gtk.Button.new_with_label("Seçimi kaldır")
        self.select_none.connect("clicked", self.on_select_all, False)
        self.clean_button = Gtk.Button.new_with_label("Seçilenleri sil ve kopya oluştur")
        self.clean_button.get_style_context().add_class("suggested-action")
        self.clean_button.connect("clicked", self.on_clean)
        for button in (browse, self.select_all, self.select_none, self.clean_button):
            actions.pack_start(button, False, False, 0)
        root.pack_start(actions, False, False, 0)
        scroller = Gtk.ScrolledWindow()
        scroller.set_vexpand(True)
        tree = Gtk.TreeView(model=self.store)
        tree.set_headers_visible(True)
        toggle = Gtk.CellRendererToggle()
        toggle.connect("toggled", self.on_toggled)
        col = Gtk.TreeViewColumn("Sil", toggle, active=0)
        tree.append_column(col)
        for number, name in ((1, "Tür"), (2, "Bulunan veri"), (3, "Değer")):
            renderer = Gtk.CellRendererText()
            renderer.set_property("ellipsize", 3)
            column = Gtk.TreeViewColumn(name, renderer, text=number)
            column.set_resizable(True)
            column.set_expand(number == 3)
            tree.append_column(column)
        scroller.add(tree)
        root.pack_start(scroller, True, True, 0)
        self.status = Gtk.Label(label="Bir dosya seçildiğinde bulunan gizlilik verileri burada listelenir.")
        self.status.set_xalign(0)
        root.pack_start(self.status, False, False, 0)

    def on_browse(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileChooserDialog("Denetlenecek dosyayı seç", self, Gtk.FileChooserAction.OPEN,
                                       ("İptal", Gtk.ResponseType.CANCEL, "Aç", Gtk.ResponseType.ACCEPT))
        file_filter = Gtk.FileFilter()
        file_filter.set_name("Desteklenen dosyalar")
        for pattern in ("*.pdf", "*.docx", "*.xlsx", "*.jpg", "*.jpeg", "*.tif", "*.tiff", "*.webp", "*.png"):
            file_filter.add_pattern(pattern)
            file_filter.add_pattern(pattern.upper())
        dialog.add_filter(file_filter)
        if dialog.run() == Gtk.ResponseType.ACCEPT:
            filename = dialog.get_filename()
            if filename:
                self.load_file(Path(filename))
        dialog.destroy()

    def on_drag_data_received(self, _widget, context, _x, _y, data, _info, time) -> None:
        uris = data.get_uris()
        if uris:
            parsed = urlparse(uris[0])
            self.load_file(Path(unquote(parsed.path)))
            context.finish(True, False, time)
        else:
            context.finish(False, False, time)

    def load_file(self, path: Path) -> None:
        if path.suffix.lower() not in SUPPORTED:
            self.show_error("Desteklenmeyen dosya", "PDF, DOCX, XLSX veya desteklenen bir görsel seçin.")
            return
        try:
            findings = self.inspector.inspect(path)
        except Exception as exc:
            self.show_error("Dosya okunamadı", str(exc))
            return
        self.source, self.findings = path, findings
        self.store.clear()
        for finding in findings:
            self.store.append([finding.selected, finding.category, finding.title, finding.value, finding.key])
        self.file_label.set_text(path.name)
        if findings:
            self.status.set_text(f"Bu dosyayı göndermeden önce {len(findings)} gizlilik verisi gözden geçirilmeli.")
        else:
            self.status.set_text("Bu dosyada desteklenen türlerde görünür bir meta veri bulunmadı.")

    def on_toggled(self, _renderer: Gtk.CellRendererToggle, path: str) -> None:
        row = self.store[path]
        row[0] = not row[0]
        key = row[4]
        for finding in self.findings:
            if finding.key == key:
                finding.selected = row[0]
                break

    def on_select_all(self, _button: Gtk.Button, active: bool) -> None:
        for row in self.store:
            row[0] = active
        for finding in self.findings:
            finding.selected = active

    def on_clean(self, _button: Gtk.Button) -> None:
        if not self.source:
            self.show_error("Dosya seçilmedi", "Önce Gözat düğmesiyle bir dosya seçin veya dosyayı sürükleyin.")
            return
        selected_count = sum(item.selected for item in self.findings)
        if not selected_count:
            self.show_error("Seçim yok", "Silmek için en az bir gizlilik verisi seçin.")
            return
        target = output_path(self.source)
        try:
            self.inspector.clean(self.source, self.findings, target)
        except Exception as exc:
            self.show_error("Temizleme tamamlanamadı", str(exc))
            return
        self.status.set_text(f"{selected_count} veri temizlendi. Yeni dosya: {target.name}")
        dialog = Gtk.MessageDialog(self, 0, Gtk.MessageType.INFO, Gtk.ButtonsType.OK, "Temiz kopya oluşturuldu")
        dialog.format_secondary_text(f"{selected_count} seçili veri kaldırıldı.\n{target}")
        dialog.run()
        dialog.destroy()

    def show_error(self, title: str, detail: str) -> None:
        dialog = Gtk.MessageDialog(self, 0, Gtk.MessageType.ERROR, Gtk.ButtonsType.CLOSE, title)
        dialog.format_secondary_text(detail)
        dialog.run()
        dialog.destroy()


class MetaKalkanApplication(Gtk.Application):
    def __init__(self) -> None:
        super().__init__(application_id="tr.gov.pardus.MetaKalkan")

    def do_activate(self) -> None:
        window = MetaKalkanWindow(self)
        window.show_all()


if __name__ == "__main__":
    application = MetaKalkanApplication()
    raise SystemExit(application.run(sys.argv))
