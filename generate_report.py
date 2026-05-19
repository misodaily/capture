"""
신한카드 검색광고 게재보고 PPT 자동 생성기
============================================

사용법:
    python generate_report.py \
        --pc "<PC URL>" \
        --mo "<MO URL>" \
        --card "신한카드 Deep Once" \
        --out "출력파일.pptx"

또는 인자 없이 실행하면 스크립트 하단의 DEFAULT 값을 사용.

동작:
1. PC URL (card-search.naver.com/item)을 1920x1080 뷰포트로 열어
   상단부터 하단까지 모든 <details>를 펼친 상태로 viewport-height 단위로
   잘라 9~10장의 풀폭 캡처를 만든다 (누락 0).
2. MO URL (m-card-search.naver.com)도 동일 방식으로 모바일 캡처.
3. 샘플 PPT(Deep Once)와 동일한 레이아웃의 PPT로 조립.
"""
from __future__ import annotations
import argparse
import math
import os
import sys
import tempfile
from pathlib import Path
from typing import List

from playwright.sync_api import sync_playwright, Page
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# ---- Capture parameters ---------------------------------------------------

PC_VIEWPORT = (1920, 1080)
PC_SCREENSHOT_H = 989   # 샘플 PPT의 PC 이미지 높이와 동일하게 맞춤
PC_OVERLAP = 60         # 슬라이드 간 약간의 겹침 (누락 방지)

PC_TITLE_VIEWPORT = (801, 1311)  # 슬라이드 1 (vertical) 캡처용 좁은 뷰포트

MO_VIEWPORT = (404, 874)         # 샘플 PPT MO 이미지와 동일
MO_TITLE_DETAIL_VIEWPORT = (391, 1258)  # 샘플의 image12 비율
MO_OVERLAP = 50

MO_USER_AGENT = (
    "Mozilla/5.0 (Linux; Android 13; SM-S918N) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36"
)


# ---- Web capture ----------------------------------------------------------

def _open_all_details(page: Page) -> None:
    """페이지 내 모든 혜택 details/접이식 요소를 펼친다.
    네이버 카드 페이지는 click 이벤트에서 본문을 lazy-load 하므로
    반드시 summary 를 click() 해야 한다 (`.open = true` 만으론 본문이 안 채워짐).
    """
    # 1) 혜택 details (details.details: 렌탈/관리비/문화)
    benefit_summaries = page.locator('details.details > summary')
    cnt = benefit_summaries.count()
    for i in range(cnt):
        try:
            s = benefit_summaries.nth(i)
            s.scroll_into_view_if_needed(timeout=3000)
            s.click(timeout=3000)
            page.wait_for_timeout(350)
        except Exception as e:
            print(f"[WARN] benefit summary[{i}] click failed: {e}")

    # 2) 기타 details (annualFee_detail, baseRecord_detail 등): open 만 처리
    page.evaluate("""() => {
        document.querySelectorAll('details:not(.details)').forEach(d => { d.open = true; });
        document.querySelectorAll('[aria-expanded="false"]').forEach(el => {
            el.setAttribute('aria-expanded', 'true');
        });
    }""")
    # 클릭 후 페이지 리플로우 대기
    page.wait_for_timeout(600)


def _hide_floating_chrome(page: Page) -> None:
    """우측 상단 공유/맨위로 버튼 등 떠다니는 요소가 캡처를 가리지 않도록.
    (샘플 PPT에는 이 버튼들도 그대로 찍혀 있으므로 우선 숨기지 않음)
    """
    pass


def _full_page_height(page: Page) -> int:
    return page.evaluate("() => document.documentElement.scrollHeight")


def _capture_tall(page: Page, out_path: Path, base_w: int, max_h: int = 16000) -> int:
    """
    누락 없는 풀페이지 캡처:
    1) fixed/sticky → absolute (Playwright stitching 시 매 viewport 마다 중복 캡처되는 문제 해결)
    2) lazy-load 트리거 (페이지 하단까지 한 번 스크롤)
    3) 최상단으로 복귀 후 `full_page=True` 로 캡처
       — viewport 를 키우면 모바일 reflow 로 인해 페이지가 무한히 늘어나는
       문제(빈 슬라이스 생성)가 있어 viewport-resize 방식은 폐기.
    """
    page.wait_for_timeout(300)
    page.evaluate("() => window.scrollTo(0, 0)")
    page.wait_for_timeout(200)

    # fixed/sticky → absolute
    page.evaluate("""() => {
        document.querySelectorAll('*').forEach(el => {
            const cs = getComputedStyle(el);
            if (cs.position === 'fixed' || cs.position === 'sticky') {
                el.style.setProperty('position', 'absolute', 'important');
            }
        });
    }""")
    page.wait_for_timeout(200)

    # lazy-load 트리거: 페이지 끝까지 스크롤 → 이미지 로드 대기 → 다시 최상단
    page.evaluate("() => window.scrollTo(0, document.documentElement.scrollHeight)")
    page.wait_for_timeout(800)
    try:
        page.evaluate("""async () => {
            await Promise.all(
              Array.from(document.images).filter(img => !img.complete)
                .map(img => new Promise(res => { img.onload = img.onerror = res; }))
            );
        }""")
    except Exception:
        pass
    page.evaluate("() => window.scrollTo(0, 0)")
    page.wait_for_timeout(400)

    page_h = page.evaluate(
        "() => Math.max(document.documentElement.scrollHeight, document.body.scrollHeight)"
    )
    page.screenshot(path=str(out_path), full_page=True, animations="disabled")
    return page_h


def _slice_full_page(
    full_path: str, slice_h: int, overlap: int, out_dir: Path, prefix: str
) -> List[Path]:
    """긴 전체 페이지 스크린샷을 슬라이드 단위로 자른다 (겹침 포함)."""
    img = Image.open(full_path)
    W, H = img.size
    step = slice_h - overlap
    n = max(1, math.ceil((H - overlap) / step))
    paths: List[Path] = []
    for i in range(n):
        top = i * step
        bottom = min(top + slice_h, H)
        # 마지막 조각이 너무 짧으면 위로 당겨서 slice_h 확보
        if bottom - top < slice_h * 0.5 and i > 0:
            top = max(0, H - slice_h)
            bottom = H
        crop = img.crop((0, top, W, bottom))
        # 정확히 slice_h 가 안 되면 캔버스에 붙여서 패딩
        if crop.size[1] < slice_h:
            canvas = Image.new("RGB", (W, slice_h), (255, 255, 255))
            canvas.paste(crop, (0, 0))
            crop = canvas
        p = out_dir / f"{prefix}_{i + 1:02d}.png"
        crop.save(p, "PNG", optimize=True)
        paths.append(p)
        if bottom >= H:
            break
    return paths


def capture_pc(url: str, work: Path) -> dict:
    """PC 캡처: 1920×1080 뷰포트로 풀페이지 캡처 → 989px 단위 슬라이스 N장."""
    work.mkdir(parents=True, exist_ok=True)
    pc_full = work / "pc_full.png"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            viewport={"width": PC_VIEWPORT[0], "height": PC_VIEWPORT[1]},
            device_scale_factor=1,
        )
        page = ctx.new_page()
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(1500)
        _open_all_details(page)
        page.wait_for_timeout(800)
        # 모든 이미지 로드 대기
        page.evaluate("""async () => {
            await Promise.all(
              Array.from(document.images)
                .filter(img => !img.complete)
                .map(img => new Promise(res => { img.onload = img.onerror = res; }))
            );
        }""")
        page.wait_for_timeout(500)
        height = _capture_tall(page, pc_full, PC_VIEWPORT[0])
        ctx.close()
        browser.close()

    slides = _slice_full_page(str(pc_full), PC_SCREENSHOT_H, PC_OVERLAP, work, "pc_slide")
    return {"slides": slides, "full": pc_full, "height": height}


def capture_mo(url: str, work: Path) -> dict:
    """MO 캡처: 404×874 모바일 뷰포트로 풀페이지 캡처 → 874px 단위 슬라이스 N장.
    PC 와 동일한 로직 — 별도의 검색결과/타이틀 캡처는 만들지 않음.
    """
    work.mkdir(parents=True, exist_ok=True)
    mo_full = work / "mo_full.png"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            viewport={"width": MO_VIEWPORT[0], "height": MO_VIEWPORT[1]},
            user_agent=MO_USER_AGENT,
            is_mobile=True, has_touch=True, device_scale_factor=1,
        )
        page = ctx.new_page()
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(1500)
        _open_all_details(page)
        page.wait_for_timeout(800)
        page.evaluate("""async () => {
            await Promise.all(
              Array.from(document.images)
                .filter(img => !img.complete)
                .map(img => new Promise(res => { img.onload = img.onerror = res; }))
            );
        }""")
        page.wait_for_timeout(500)
        height = _capture_tall(page, mo_full, MO_VIEWPORT[0])
        ctx.close()
        browser.close()

    slides = _slice_full_page(str(mo_full), MO_VIEWPORT[1], MO_OVERLAP, work, "mo_slide")
    return {"slides": slides, "full": mo_full, "height": height}


# ---- PPT building ---------------------------------------------------------

# 슬라이드 크기 (16:9, 12192000 x 6858000 EMU = 13.33" x 7.5")
SLIDE_W_EMU = Emu(12192000)
SLIDE_H_EMU = Emu(6858000)

TITLE_BLUE = RGBColor(0x44, 0x72, 0xC4)   # 좌상단 보조 텍스트 컬러
ACCENT_BLUE = RGBColor(0x5B, 0x9B, 0xD5)  # 라인 색

HEADER_TEXT_LEFT = "신한카드 멤버십 영업팀"   # 상단 좌측 (작은 청색)
HEADER_TEXT_RIGHT_PC = "네이버 신용카드 검색광고_PC"
HEADER_TEXT_RIGHT_MO = "네이버 신용카드 검색광고_MO"


def _add_header(slide, prs, right_text: str):
    """샘플 PPT 스타일 상단 헤더 (좌: 회사명, 우: PC/MO 표시)."""
    # 좌측 회사명
    tb = slide.shapes.add_textbox(Emu(234953), Emu(120000), Emu(6000000), Emu(420000))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = HEADER_TEXT_LEFT
    r = p.runs[0]
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = TITLE_BLUE

    # 우측 PC/MO 라벨
    tb2 = slide.shapes.add_textbox(Emu(8200000), Emu(120000), Emu(3800000), Emu(420000))
    tf2 = tb2.text_frame
    tf2.margin_left = tf2.margin_right = tf2.margin_top = tf2.margin_bottom = 0
    p2 = tf2.paragraphs[0]
    p2.alignment = 2  # right
    p2.text = right_text
    r2 = p2.runs[0]
    r2.font.size = Pt(14)
    r2.font.bold = False
    r2.font.color.rgb = TITLE_BLUE


def _add_image_centered(slide, img_path: Path, top_emu: int = 700000,
                        max_w_emu: int = None, max_h_emu: int = None):
    """슬라이드 가운데 상단 정렬로 이미지 배치 (비율 유지)."""
    img = Image.open(img_path)
    w, h = img.size
    max_w_emu = max_w_emu or int(SLIDE_W_EMU * 0.95)
    max_h_emu = max_h_emu or (int(SLIDE_H_EMU) - top_emu - 200000)
    # EMU per pixel at 96 DPI: 9525
    px_to_emu = 9525
    target_w = w * px_to_emu
    target_h = h * px_to_emu
    scale = min(max_w_emu / target_w, max_h_emu / target_h, 1.0)
    final_w = int(target_w * scale)
    final_h = int(target_h * scale)
    left = (int(SLIDE_W_EMU) - final_w) // 2
    slide.shapes.add_picture(str(img_path), left, top_emu, width=final_w, height=final_h)


def _add_image_grid(slide, paths: List[Path], top_emu: int = 700000):
    """모바일 썸네일들을 한 줄로 가운데 정렬 배치."""
    if not paths:
        return
    gap_emu = Emu(150000)
    # 슬라이드 폭 95% 사용 / 길이만큼 분할
    avail_w = int(SLIDE_W_EMU * 0.95)
    avail_h = int(SLIDE_H_EMU) - top_emu - 200000
    # 첫 이미지로 비율 계산
    sample = Image.open(paths[0])
    sw, sh = sample.size
    # 각 카드 폭 추정 후 높이 맞추기
    n = len(paths)
    per_w_emu = (avail_w - int(gap_emu) * (n - 1)) // n
    # px→emu
    px_to_emu = 9525
    target_w = sw * px_to_emu
    target_h = sh * px_to_emu
    scale_w = per_w_emu / target_w
    scale_h = avail_h / target_h
    scale = min(scale_w, scale_h, 1.0)
    final_w = int(target_w * scale)
    final_h = int(target_h * scale)
    total_w = final_w * n + int(gap_emu) * (n - 1)
    left0 = (int(SLIDE_W_EMU) - total_w) // 2
    for i, pth in enumerate(paths):
        left = left0 + i * (final_w + int(gap_emu))
        slide.shapes.add_picture(str(pth), left, top_emu, width=final_w, height=final_h)


def build_pptx(
    pc: dict, mo: dict, out_path: Path,
    card_name: str, slides_per_mo_grid: int = 5, **_ignored,
) -> Path:
    """PC, MO 캡처를 PPT 로 조립.
    - PC: 1슬라이스 = 1슬라이드 (풀폭 가운데 정렬)
    - MO: 1슬라이드 당 5장씩 가로 그리드 정렬
    """
    prs = Presentation()
    prs.slide_width = SLIDE_W_EMU
    prs.slide_height = SLIDE_H_EMU
    blank_layout = prs.slide_layouts[6]

    # ---- PC ----
    for slice_path in pc["slides"]:
        s = prs.slides.add_slide(blank_layout)
        _add_header(s, prs, HEADER_TEXT_RIGHT_PC)
        _add_image_centered(s, slice_path, top_emu=650000)

    # ---- MO: 5장씩 그리드 ----
    mo_slices = mo["slides"]
    for i in range(0, len(mo_slices), slides_per_mo_grid):
        chunk = mo_slices[i: i + slides_per_mo_grid]
        s = prs.slides.add_slide(blank_layout)
        _add_header(s, prs, HEADER_TEXT_RIGHT_MO)
        _add_image_grid(s, chunk, top_emu=700000)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


# ---- CLI ------------------------------------------------------------------

DEFAULT_PC = (
    "https://card-search.naver.com/item?cardAdId=3257"
    "&query=%EC%8B%A0%ED%95%9C%EC%B9%B4%EB%93%9C%20Deep%20Once"
    "&cx=LNFvYD7ZLUB2W2S%2FZrj0TA%3D%3D"
)
DEFAULT_MO = (
    "https://m-card-search.naver.com/item?cardAdId=3257"
    "&query=%EC%8B%A0%ED%95%9C%EC%B9%B4%EB%93%9C%20Deep%20Once"
    "&cx=LNFvYD7ZLUB2W2S%2FZrj0TA%3D%3D"
)


def main():
    ap = argparse.ArgumentParser(description="신한카드 검색광고 게재보고 PPT 자동 생성")
    ap.add_argument("--pc", default=DEFAULT_PC, help="PC card-search.naver.com/item URL")
    ap.add_argument("--mo", default=DEFAULT_MO, help="MO m-card-search.naver.com/item URL")
    ap.add_argument("--card", default="신한카드 Deep Once", help="카드명 (출력 파일명/타이틀용)")
    ap.add_argument(
        "--out",
        default=None,
        help="출력 PPT 경로 (생략 시 바탕화면에 자동 생성)",
    )
    ap.add_argument("--work", default=None, help="중간 캡처 저장 폴더 (생략 시 자동)")
    args = ap.parse_args()

    work = Path(args.work) if args.work else Path(tempfile.mkdtemp(prefix="report_"))
    out_default = Path(
        rf"C:\Users\FIN\Desktop\(핀플로우) 신한카드 멤버십 영업팀 신용카드 검색광고 게재보고_{args.card}.pptx"
    )
    out = Path(args.out) if args.out else out_default

    print(f"[1/3] PC 캡처 중... (work={work})")
    pc = capture_pc(args.pc, work / "pc")
    print(f"      slices: {len(pc['slides'])}, page H={pc['height']}px")

    print("[2/3] MO 캡처 중...")
    mo = capture_mo(args.mo, work / "mo")
    print(f"      slices: {len(mo['slides'])}, page H={mo['height']}px")

    print("[3/3] PPT 생성 중...")
    p = build_pptx(pc, mo, out, args.card)
    print(f"DONE → {p}")
    print(f"      (중간 캡처는 {work} 에 보존)")


if __name__ == "__main__":
    main()
