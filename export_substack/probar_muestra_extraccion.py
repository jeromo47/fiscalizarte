#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import re
from html import unescape
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from faster_whisper import WhisperModel

BASE_URL = "https://fiscalizarte.substack.com"
WORKDIR = Path.home() / "fiscalizarte_export"
COOKIES_PATH = WORKDIR / "cookies.txt"
SAMPLE_PATH = WORKDIR / "muestra_tipos.csv"
OUT_CSV = WORKDIR / "muestra_tipos_resultado.csv"
OUT_XLSX = WORKDIR / "muestra_tipos_resultado.xlsx"


def load_netscape_cookies(path: Path, session: requests.Session) -> None:
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("\t") if "\t" in line else line.split()
        if len(parts) < 7:
            continue
        domain, _flag, cookie_path, secure, expires, name, value = parts[:7]
        session.cookies.set(name, value, domain=domain, path=cookie_path or "/", secure=(secure.upper() == "TRUE"))


def extract_preloads_json(html: str) -> dict:
    m = re.search(r'window\._preloads\s*=\s*JSON\.parse\((".*?")\)\s*</script>', html, re.S)
    if not m:
        m = re.search(r'window\._preloads\s*=\s*JSON\.parse\((".*?")\)', html, re.S)
    if not m:
        return {}
    inner = json.loads(m.group(1))
    return json.loads(inner)


def html_to_text(body_html: str) -> str:
    if not body_html:
        return ""
    soup = BeautifulSoup(body_html, "html.parser")
    return soup.get_text("\n", strip=True)


def parse_embed_media_ids(body_html: str, class_name: str) -> list[str]:
    pattern = rf'<div class="{re.escape(class_name)}"[^>]*data-attrs="([^"]+)"'
    found = []
    for attrs in re.findall(pattern, body_html or ""):
        try:
            data = json.loads(unescape(attrs))
            media_id = data.get("mediaUploadId")
            if media_id:
                found.append(str(media_id))
        except Exception:
            pass
    return found


def fetch_json_text(session: requests.Session, url: str) -> str:
    r = session.get(url, timeout=30)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, list):
        return " ".join((item.get("text") or "").strip() for item in data if isinstance(item, dict) and item.get("text"))
    return ""


def asr_from_video_url(src: str, cache_name: str, model: WhisperModel) -> str:
    out = WORKDIR / cache_name
    if not out.exists():
        with requests.get(src, stream=True, timeout=60) as r:
            r.raise_for_status()
            with out.open("wb") as f:
                for chunk in r.iter_content(1024 * 1024):
                    if chunk:
                        f.write(chunk)
    segments, info = model.transcribe(str(out), language="es", beam_size=1, vad_filter=True)
    parts = []
    for seg in segments:
        t = seg.text.strip()
        if t:
            parts.append(t)
    return " ".join(parts).strip()


def main() -> int:
    df = pd.read_csv(SAMPLE_PATH)
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
        "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
        "Accept": "application/json, text/plain, */*",
    })
    load_netscape_cookies(COOKIES_PATH, session)
    model = WhisperModel("tiny", device="cpu", compute_type="int8")

    rows = []
    for _, row in df.iterrows():
        url = row["url"]
        tipo = row["tipo_clasificado"]
        html = session.get(url, timeout=30).text
        data = extract_preloads_json(html)
        post = data.get("post", {})
        body_text = html_to_text(post.get("body_html") or "")
        transcript_text = ""
        transcript_source = "none"
        notes = ""

        if tipo == "video_o_podcast_principal_con_transcripcion_substack":
            vu = post.get("videoUpload") or {}
            tr = (vu.get("extractedAudio") or {}).get("transcription") or vu.get("transcription") or {}
            if tr.get("cdn_url"):
                transcript_text = fetch_json_text(session, tr["cdn_url"])
                transcript_source = "substack_video"
        elif tipo == "podcast_audio_con_transcripcion_substack":
            pu = post.get("podcastUpload") or {}
            tr = pu.get("transcription") or {}
            if tr.get("cdn_url"):
                transcript_text = fetch_json_text(session, tr["cdn_url"])
                transcript_source = "substack_podcast"
        elif tipo == "newsletter_con_video_embebido":
            body_html = post.get("body_html") or ""
            media_ids = parse_embed_media_ids(body_html, "native-video-embed")
            if media_ids:
                mid = media_ids[0]
                src = session.get(f"{BASE_URL}/api/v1/video/upload/{mid}/src.json?type=original", timeout=30).json()["src"]
                transcript_text = asr_from_video_url(src, f"sample_{mid}.mp4", model)
                transcript_source = "local_asr_video"
        elif tipo == "podcast_video_sin_transcripcion_substack":
            vu = post.get("videoUpload") or {}
            if vu.get("id"):
                src = session.get(f"{BASE_URL}/api/v1/video/upload/{vu['id']}/src.json?type=original", timeout=30).json()["src"]
                transcript_text = asr_from_video_url(src, f"sample_{vu['id']}.mp4", model)
                transcript_source = "local_asr_video"
        elif tipo == "texto_normal":
            transcript_source = "none"
        else:
            notes = "Tipo no manejado en muestra"

        combined = (body_text + "\n\n" + transcript_text).strip() if transcript_text else body_text.strip()
        rows.append({
            "titulo": post.get("title") or row["titulo"],
            "url": url,
            "tipo_clasificado": tipo,
            "body_text": body_text,
            "body_length": len(body_text),
            "transcript_text": transcript_text,
            "transcript_length": len(transcript_text),
            "transcript_source": transcript_source,
            "combined_text": combined,
            "combined_length": len(combined),
            "estado_extraccion": "OK" if combined else "ERROR",
            "notas": notes,
        })

    out = pd.DataFrame(rows)
    out.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    out.to_excel(OUT_XLSX, index=False)
    print(out[["tipo_clasificado", "titulo", "body_length", "transcript_length", "transcript_source", "estado_extraccion"]].to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
