"""
신한카드 검색광고 게재보고 자동 생성 - 로컬 웹 프론트엔드
============================================================

실행: python server.py
브라우저에서 http://127.0.0.1:5180 자동 열림.

API:
- POST /api/generate { pc_url, mo_url, card_name }  -> { job_id }
- GET  /api/status/<job_id>                          -> { status, step, progress, file?, error? }
- GET  /api/download/<job_id>                        -> PPT 파일 다운로드
- POST /api/open/<job_id>                            -> 로컬 PowerPoint 로 파일 열기
- GET  /api/history                                  -> 최근 생성 이력
- POST /api/extract-card { url }                     -> URL 의 query 파라미터에서 카드명 추출
"""
from __future__ import annotations
import os
import sys
import json
import threading
import traceback
import uuid
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Dict, Any
from urllib.parse import urlparse, parse_qs, unquote

from flask import Flask, request, jsonify, send_file, abort

import generate_report  # 같은 폴더의 메인 모듈

APP_DIR = Path(__file__).parent.resolve()
OUTPUT_DIR = Path.home() / "Desktop"
WORK_ROOT = APP_DIR / "_work"
WORK_ROOT.mkdir(exist_ok=True)
HISTORY_PATH = APP_DIR / "_history.json"

app = Flask(__name__, static_folder=str(APP_DIR / "static"), static_url_path="/static")

# job_id -> { status, step, progress(0-100), file_path, error, started_at, finished_at, card_name }
JOBS: Dict[str, Dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()


def _load_history() -> list:
    if HISTORY_PATH.exists():
        try:
            return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _save_history(items: list) -> None:
    try:
        HISTORY_PATH.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _push_history(job: Dict[str, Any]) -> None:
    if job.get("status") != "done" or not job.get("file_path"):
        return
    items = _load_history()
    items.insert(0, {
        "card_name": job.get("card_name"),
        "file_path": job.get("file_path"),
        "file_name": Path(job["file_path"]).name,
        "finished_at": job.get("finished_at"),
        "pc_url": job.get("pc_url"),
        "mo_url": job.get("mo_url"),
    })
    _save_history(items[:30])  # 최근 30 건만


def _update(job_id: str, **kwargs) -> None:
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(kwargs)


def _build_out_path(card_name: str) -> Path:
    safe = "".join(c for c in card_name if c not in '\\/:*?"<>|').strip() or "신한카드"
    base = f"(핀플로우) 신한카드 멤버십 영업팀 신용카드 검색광고 게재보고_{safe}.pptx"
    out = OUTPUT_DIR / base
    i = 1
    while out.exists():
        out = OUTPUT_DIR / f"(핀플로우) 신한카드 멤버십 영업팀 신용카드 검색광고 게재보고_{safe}_{i}.pptx"
        i += 1
    return out


def _run_job(
    job_id: str, pc_url: str, mo_url: str, card_name: str,
    search_pc_url: str = "", search_mo_url: str = "",
    landing_url: str = "",
) -> None:
    work = WORK_ROOT / job_id
    work.mkdir(parents=True, exist_ok=True)
    out_path = _build_out_path(card_name)

    try:
        # Playwright sync API 는 한 process 당 한 인스턴스만 허용
        # (두 thread 동시 진입 시 hang) → 순차 실행
        _update(job_id, status="running", step="PC 상세 페이지 캡처 중...", progress=8)
        pc = generate_report.capture_pc(pc_url, work / "pc")
        _update(job_id, step=f"PC 완료 ({len(pc['slides'])}장)", progress=28)

        _update(job_id, step="MO 상세 페이지 캡처 중...", progress=32)
        mo = generate_report.capture_mo(mo_url, work / "mo")
        _update(job_id, step=f"MO 완료 ({len(mo['slides'])}장)", progress=56)

        search_pc_data = None
        if search_pc_url:
            _update(job_id, step="PC 검색결과 OneBox 캡처 중...", progress=60)
            search_pc_data = generate_report.capture_search_onebox(
                search_pc_url, work / "search_pc", is_mobile=False
            )

        search_mo_data = None
        if search_mo_url:
            _update(job_id, step="MO 검색결과 OneBox 캡처 중...", progress=68)
            search_mo_data = generate_report.capture_search_onebox(
                search_mo_url, work / "search_mo", is_mobile=True
            )

        landing_pc_data = None
        landing_mo_data = None
        if landing_url:
            _update(job_id, step="안내 페이지 캡처 중 (PC)...", progress=75)
            landing_pc_data = generate_report.capture_landing(
                landing_url, work / "landing_pc", is_mobile=False
            )
            _update(job_id, step="안내 페이지 캡처 중 (MO)...", progress=83)
            landing_mo_data = generate_report.capture_landing(
                landing_url, work / "landing_mo", is_mobile=True
            )

        _update(job_id, step="PPT 조립 중...", progress=92)
        generate_report.build_pptx(
            pc, mo, out_path, card_name,
            search_pc=search_pc_data, search_mo=search_mo_data,
            landing_pc=landing_pc_data, landing_mo=landing_mo_data,
        )

        with JOBS_LOCK:
            JOBS[job_id].update({
                "status": "done",
                "step": "완료",
                "progress": 100,
                "file_path": str(out_path),
                "finished_at": datetime.now().isoformat(timespec="seconds"),
                "pc_slides": len(pc["slides"]),
                "mo_slides": len(mo["slides"]),
                "pc_height": pc["height"],
                "mo_height": mo["height"],
                "search_pc_slides": len(search_pc_data["slides"]) if search_pc_data else 0,
                "search_mo_slides": len(search_mo_data["slides"]) if search_mo_data else 0,
                "landing_pc_slides": len(landing_pc_data["slides"]) if landing_pc_data else 0,
                "landing_mo_slides": len(landing_mo_data["slides"]) if landing_mo_data else 0,
            })
        _push_history(JOBS[job_id])

    except Exception as e:
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)
        _update(
            job_id,
            status="error",
            step="실패",
            progress=100,
            error=f"{type(e).__name__}: {e}",
            traceback=tb,
            finished_at=datetime.now().isoformat(timespec="seconds"),
        )


# -------------------- API --------------------

@app.route("/")
def index():
    return send_file(APP_DIR / "static" / "index.html")


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True) or {}
    pc_url = (data.get("pc_url") or "").strip()
    mo_url = (data.get("mo_url") or "").strip()
    search_pc_url = (data.get("search_pc_url") or "").strip()
    search_mo_url = (data.get("search_mo_url") or "").strip()
    landing_url = (data.get("landing_url") or "").strip()
    card_name = (data.get("card_name") or "").strip() or _extract_card_from_url(pc_url) or "신한카드"

    if not pc_url or not mo_url:
        return jsonify({"error": "PC URL 과 MO URL 을 모두 입력해 주세요."}), 400
    if "card-search.naver.com" not in pc_url:
        return jsonify({"error": "PC URL 은 card-search.naver.com 도메인이어야 합니다."}), 400
    if "m-card-search.naver.com" not in mo_url:
        return jsonify({"error": "MO URL 은 m-card-search.naver.com 도메인이어야 합니다."}), 400
    if search_pc_url and "search.naver.com" not in search_pc_url:
        return jsonify({"error": "PC 검색결과 URL 은 search.naver.com 도메인이어야 합니다."}), 400
    if search_mo_url and "m.search.naver.com" not in search_mo_url:
        return jsonify({"error": "MO 검색결과 URL 은 m.search.naver.com 도메인이어야 합니다."}), 400
    if landing_url and not landing_url.startswith(("http://", "https://")):
        return jsonify({"error": "안내 페이지 URL 은 http:// 또는 https:// 로 시작해야 합니다."}), 400

    job_id = uuid.uuid4().hex[:12]
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "queued",
            "step": "대기 중",
            "progress": 0,
            "card_name": card_name,
            "pc_url": pc_url,
            "mo_url": mo_url,
            "search_pc_url": search_pc_url,
            "search_mo_url": search_mo_url,
            "landing_url": landing_url,
            "started_at": datetime.now().isoformat(timespec="seconds"),
        }
    threading.Thread(
        target=_run_job,
        args=(job_id, pc_url, mo_url, card_name, search_pc_url, search_mo_url, landing_url),
        daemon=True,
    ).start()
    return jsonify({"job_id": job_id, "card_name": card_name})


@app.route("/api/status/<job_id>")
def api_status(job_id: str):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "not_found"}), 404
        # file_name 만 노출 (다운로드 링크 구성용)
        out = dict(job)
        if out.get("file_path"):
            out["file_name"] = Path(out["file_path"]).name
        return jsonify(out)


@app.route("/api/download/<job_id>")
def api_download(job_id: str):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job or not job.get("file_path"):
        abort(404)
    p = Path(job["file_path"])
    if not p.exists():
        abort(404)
    return send_file(p, as_attachment=True, download_name=p.name)


@app.route("/api/open/<job_id>", methods=["POST"])
def api_open(job_id: str):
    """Windows 에서 로컬 PowerPoint 로 파일 열기."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job or not job.get("file_path"):
        return jsonify({"error": "not_found"}), 404
    p = Path(job["file_path"])
    if not p.exists():
        return jsonify({"error": "file_missing"}), 404
    try:
        os.startfile(str(p))  # type: ignore[attr-defined]
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/reveal/<job_id>", methods=["POST"])
def api_reveal(job_id: str):
    """탐색기에서 파일 위치 열기."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job or not job.get("file_path"):
        return jsonify({"error": "not_found"}), 404
    p = Path(job["file_path"])
    if not p.exists():
        return jsonify({"error": "file_missing"}), 404
    try:
        import subprocess
        subprocess.Popen(["explorer.exe", "/select,", str(p)])
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/open-path", methods=["POST"])
def api_open_path():
    data = request.get_json(force=True) or {}
    p = Path(data.get("path", ""))
    if not p.exists():
        return jsonify({"error": "file_missing"}), 404
    try:
        os.startfile(str(p))  # type: ignore[attr-defined]
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/reveal-path", methods=["POST"])
def api_reveal_path():
    data = request.get_json(force=True) or {}
    p = Path(data.get("path", ""))
    if not p.exists():
        return jsonify({"error": "file_missing"}), 404
    try:
        import subprocess
        subprocess.Popen(["explorer.exe", "/select,", str(p)])
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/history")
def api_history():
    items = _load_history()
    # 실제 파일이 존재하는 것만
    items = [i for i in items if Path(i.get("file_path", "")).exists()]
    return jsonify({"items": items})


@app.route("/api/extract-card", methods=["POST"])
def api_extract():
    data = request.get_json(force=True) or {}
    url = (data.get("url") or "").strip()
    name = _extract_card_from_url(url)
    return jsonify({"card_name": name})


def _extract_card_from_url(url: str) -> str | None:
    try:
        qs = parse_qs(urlparse(url).query)
        q = qs.get("query", [None])[0]
        if q:
            return unquote(q).strip()
    except Exception:
        pass
    return None


# -------------------- main --------------------

def main():
    host = "127.0.0.1"
    port = 5180
    url = f"http://{host}:{port}/"
    print(f"\n게재보고 자동화 - 로컬 서버 시작")
    print(f"  → 브라우저에서 {url}")
    print(f"  → 종료: Ctrl+C\n")
    # 별도 스레드에서 1초 후 브라우저 열기 (서버가 listen 시작한 뒤)
    def _open():
        import time
        time.sleep(1.2)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=_open, daemon=True).start()
    # debug=False 이고 use_reloader=False 라 부모 프로세스 안에서 한 번만 떠야 함
    app.run(host=host, port=port, debug=False, use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()
