#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import sys
import html
import time
import os
import io
import argparse
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta
from typing import List, Optional, Dict, Tuple
from urllib.parse import urljoin

import feedparser
import requests
from bs4 import BeautifulSoup
from PIL import Image  # python3 -m pip install pillow

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# -----------------------------
# Veri Modeli
# -----------------------------
@dataclass
class Haber:
    kaynak: str
    baslik: str
    link: str
    tarih: datetime
    ozet: str
    entry: object  # RSS entry (resim bulmak için)


# -----------------------------
# RSS Kaynakları (5adet)
# -----------------------------
RSS_KAYNAKLAR: List[Tuple[str, str]] = [
    ("The Hacker News", "https://thehackernews.com/feeds/posts/default?alt=rss"),
    ("BleepingComputer", "https://www.bleepingcomputer.com/feed/"),
    ("Krebs on Security", "https://krebsonsecurity.com/feed/"),
    ("Cisco Talos Intelligence", "https://blog.talosintelligence.com/rss/"),
    ("PortSwigger Research", "https://portswigger.net/research/rss"),
]


# -----------------------------
# Yardımcılar
# -----------------------------
def simdi_utc() -> datetime:
    return datetime.now(timezone.utc)


def temiz_metin(s: str) -> str:
    if not s:
        return ""
    s = html.unescape(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def html_to_text(html_str: str) -> str:
    if not html_str:
        return ""
    soup = BeautifulSoup(html_str, "lxml")
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    return temiz_metin(soup.get_text(separator=" ", strip=True))


def entry_tarihi(entry) -> Optional[datetime]:
    for key in ("published_parsed", "updated_parsed"):
        if hasattr(entry, key) and getattr(entry, key):
            st = getattr(entry, key)
            try:
                return datetime.fromtimestamp(time.mktime(st), tz=timezone.utc)
            except Exception:
                pass
    return None


def meta_description_getir(url: str, timeout: int = 12) -> str:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; CyberNewsDocxBot/1.0)"}
        r = requests.get(url, headers=headers, timeout=timeout)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")

        for sel in [
            ("meta", {"name": "description"}),
            ("meta", {"property": "og:description"}),
            ("meta", {"name": "twitter:description"}),
        ]:
            tag = soup.find(sel[0], sel[1])
            if tag and tag.get("content"):
                d = temiz_metin(tag["content"])
                if d:
                    return d

        ps = soup.find_all("p")
        chunks = []
        for p in ps:
            t = temiz_metin(p.get_text(" ", strip=True))
            if t and len(t) > 60:
                chunks.append(t)
            if len(chunks) >= 2:
                break
        return " ".join(chunks)[:600].strip()
    except Exception:
        return ""


def ozet_uret(entry, link: str, max_len: int = 420, sayfadan_getir: bool = True) -> str:
    raw = ""
    if hasattr(entry, "summary") and entry.summary:
        raw = entry.summary
    elif hasattr(entry, "description") and entry.description:
        raw = entry.description

    text = html_to_text(raw)

    if sayfadan_getir and (not text or len(text) < 80):
        fetched = meta_description_getir(link)
        if fetched and len(fetched) > len(text):
            text = fetched

    text = temiz_metin(text)
    if len(text) > max_len:
        text = text[: max_len - 1].rstrip() + "…"
    return text


def son_gunleri_filtrele(items: List[Haber], gun: int) -> List[Haber]:
    cutoff = simdi_utc() - timedelta(days=gun)
    return [x for x in items if x.tarih and x.tarih >= cutoff]


def tekrarlari_temizle(items: List[Haber]) -> List[Haber]:
    seen = set()
    out = []
    for it in items:
        key = (it.baslik.lower().strip(), it.link.strip())
        if key in seen:
            continue
        seen.add(key)
        out.append(it)
    return out


def fmt_tarih(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


# -----------------------------
# Görsel: URL bul / indir / DOCX'e ekle
# -----------------------------
def gorsel_url_bul(entry, link: str) -> str:
    # 1) media:content / media:thumbnail
    for attr in ("media_content", "media_thumbnail"):
        if hasattr(entry, attr):
            arr = getattr(entry, attr) or []
            if isinstance(arr, list) and arr:
                u = arr[0].get("url")
                if u:
                    return temiz_metin(u)

    # 2) enclosures
    if hasattr(entry, "links"):
        for l in entry.links:
            if l.get("rel") == "enclosure":
                href = l.get("href")
                typ = (l.get("type") or "").lower()
                if href and ("image" in typ or href.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))):
                    return temiz_metin(href)

    # 3) sayfadan og:image / twitter:image
    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; CyberNewsDocxBot/1.0)"}
        r = requests.get(link, headers=headers, timeout=12)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")

        for sel in [
            ("meta", {"property": "og:image"}),
            ("meta", {"name": "twitter:image"}),
            ("meta", {"property": "og:image:secure_url"}),
        ]:
            tag = soup.find(sel[0], sel[1])
            if tag and tag.get("content"):
                u = temiz_metin(tag["content"])
                if u:
                    return urljoin(link, u)
    except Exception:
        pass

    return ""


def gorsel_indir_tmp(image_url: str, timeout: int = 15) -> str:
    if not image_url:
        return ""

    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; CyberNewsDocxBot/1.0)"}
        r = requests.get(image_url, headers=headers, timeout=timeout, stream=True)
        r.raise_for_status()

        ctype = (r.headers.get("Content-Type") or "").lower()
        if "image" not in ctype:
            return ""

        content = r.content
        if not content:
            return ""

        # Çok büyükleri alma (6MB)
        if len(content) > 6 * 1024 * 1024:
            return ""

        # Pillow ile doğrula
        with Image.open(io.BytesIO(content)) as im:
            im.verify()

        ext = ".jpg"
        if "png" in ctype:
            ext = ".png"
        elif "webp" in ctype:
            ext = ".webp"

        fd, path = tempfile.mkstemp(suffix=ext)
        with os.fdopen(fd, "wb") as f:
            f.write(content)
        return path

    except Exception:
        return ""


def karta_gorsel_ekle(right_cell, entry, link: str, max_width_inch: float = 4.9) -> bool:
    """
    Resmi başlığın altına, biraz küçültülmüş ve ince çerçeveli şekilde ekler.
    Tasarımı bozmaz.
    """
    img_url = gorsel_url_bul(entry, link)
    if not img_url:
        return False

    img_path = gorsel_indir_tmp(img_url)
    if not img_path:
        return False

    try:
        # Resim için tek hücreli mini tablo (çerçeve efekti)
        img_table = right_cell.add_table(rows=1, cols=1)
        img_table.autofit = False
        img_table.columns[0].width = Inches(max_width_inch)

        cell = img_table.cell(0, 0)

        # 🔲 İnce, modern çerçeve
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            border = OxmlElement(f"w:{side}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "6")           # ince çizgi
            border.set(qn("w:color"), "D1D5DB")   # açık gri
            tcBorders.append(border)
        tcPr.append(tcBorders)

        # İç padding (nefes aldırır)
        set_cell_margins(cell, top=80, start=80, bottom=80, end=80)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run()
        run.add_picture(img_path, width=Inches(max_width_inch))

        return True

    except Exception:
        return False

    finally:
        try:
            os.remove(img_path)
        except Exception:
            pass


# -----------------------------
# DOCX stil yardımcıları
# -----------------------------
def set_doc_defaults(doc: Document):
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)


def add_hyperlink(paragraph, url: str, text: str):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    c = OxmlElement("w:color")
    c.set(qn("w:val"), "0563C1")
    rPr.append(c)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def shade_cell(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color.replace("#", ""))
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{k}")
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


# -----------------------------
# Rapor oluşturma
# -----------------------------
def docx_olustur(items: List[Haber], out_path: str, gun: int):
    doc = Document()
    set_doc_defaults(doc)

    title = doc.add_paragraph("Haftalık Siber Güvenlik Haber Özeti")
    title.style = doc.styles["Title"]
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    subtitle = doc.add_paragraph(
        f"Kapsam: Son {gun} gün • Oluşturulma: {fmt_tarih(simdi_utc())}"
    )
    subtitle.runs[0].font.size = Pt(10)
    subtitle.runs[0].font.color.rgb = RGBColor(90, 90, 90)

    doc.add_paragraph("")

    by_source: Dict[str, int] = {}
    for it in items:
        by_source[it.kaynak] = by_source.get(it.kaynak, 0) + 1

    stats = doc.add_paragraph()
    stats_run = stats.add_run(f"Toplam haber: {len(items)}  •  Kaynak sayısı: {len(by_source)}")
    stats_run.bold = True
    stats_run.font.size = Pt(10)

    if by_source:
        detail = ", ".join(
            [f"{k}: {v}" for k, v in sorted(by_source.items(), key=lambda x: (-x[1], x[0]))]
        )
        p = doc.add_paragraph(detail)
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.color.rgb = RGBColor(90, 90, 90)

    doc.add_paragraph("")

    for idx, it in enumerate(items, start=1):
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        table.columns[0].width = Inches(1.4)
        table.columns[1].width = Inches(5.8)

        left = table.cell(0, 0)
        right = table.cell(0, 1)

        # ✅ Tasarım: koyu sol panel + açık gri içerik
        shade_cell(left, "#111827")
        shade_cell(right, "#F3F4F6")
        set_cell_margins(left)
        set_cell_margins(right)

        # Sol: Kaynak + Tarih
        lp = left.paragraphs[0]
        lp.alignment = WD_ALIGN_PARAGRAPH.LEFT

        r1 = lp.add_run(it.kaynak.upper())
        r1.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(255, 255, 255)

        lp2 = left.add_paragraph(fmt_tarih(it.tarih))
        lp2.runs[0].font.size = Pt(9)
        lp2.runs[0].font.color.rgb = RGBColor(220, 220, 220)

        # Sağ: Başlık
        rp = right.paragraphs[0]
        rp.paragraph_format.space_after = Pt(4)

        tr = rp.add_run(f"{idx}. {it.baslik}")
        tr.bold = True
        tr.font.size = Pt(12)
        tr.font.color.rgb = RGBColor(17, 24, 39)

        # ✅ Resim (varsa) başlığın hemen altına ekle (küçük + çerçeveli)
        karta_gorsel_ekle(right, it.entry, it.link, max_width_inch=4.9)

        # Özet
        sp = right.add_paragraph(it.ozet or "—")
        sp.runs[0].font.size = Pt(10)
        sp.runs[0].font.color.rgb = RGBColor(55, 65, 81)

        # Link
        link_p = right.add_paragraph()
        add_hyperlink(link_p, it.link, "Haberi aç")

        doc.add_paragraph("")

    doc.save(out_path)


# -----------------------------
# Çekme pipeline
# -----------------------------
def tumunu_cek(gun: int, limit_kaynak: int, sayfadan_getir: bool) -> List[Haber]:
    items: List[Haber] = []

    for kaynak_adi, feed_url in RSS_KAYNAKLAR:
        feed = feedparser.parse(feed_url)
        count = 0

        for entry in feed.entries:
            if count >= limit_kaynak:
                break

            tarih = entry_tarihi(entry)
            if not tarih:
                continue

            baslik = temiz_metin(getattr(entry, "title", ""))
            link = temiz_metin(getattr(entry, "link", ""))
            if not baslik or not link:
                continue

            ozet = ozet_uret(entry, link, sayfadan_getir=sayfadan_getir)

            items.append(
                Haber(
                    kaynak=kaynak_adi,
                    baslik=baslik,
                    link=link,
                    tarih=tarih,
                    ozet=ozet,
                    entry=entry,
                )
            )
            count += 1

    items = son_gunleri_filtrele(items, gun=gun)
    items = tekrarlari_temizle(items)
    items.sort(key=lambda x: x.tarih, reverse=True)
    return items


def main():
    ap = argparse.ArgumentParser(
        description="Son N gündeki siber güvenlik haberlerini RSS ile çekip şık DOCX raporu üretir (görselli)."
    )
    ap.add_argument("--gun", type=int, default=7, help="Kaç gün geriye gidilsin (varsayılan: 7)")
    ap.add_argument("--limit", type=int, default=25, help="Kaynak başına en fazla kaç haber (varsayılan: 25)")
    ap.add_argument("--out", type=str, default="haftalik_siber_haberler.docx", help="Çıktı DOCX dosyası")
    ap.add_argument("--no-fetch", action="store_true", help="Özet/görsel için sayfayı çekme (daha hızlı)")
    args = ap.parse_args()

    items = tumunu_cek(gun=args.gun, limit_kaynak=args.limit, sayfadan_getir=(not args.no_fetch))

    if not items:
        print("Seçilen zaman aralığında haber bulunamadı.")
        sys.exit(0)

    docx_olustur(items, out_path=args.out, gun=args.gun)
    print(f"✅ DOCX raporu oluşturuldu: {args.out} ({len(items)} haber)")


if __name__ == "__main__":
    main()