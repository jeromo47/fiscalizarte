#!/usr/bin/env python3
from __future__ import annotations

import json
import re
from dataclasses import asdict
from html import unescape
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from faster_whisper import WhisperModel

BASE_URL = "https://fiscalizarte.substack.com"
WORKDIR = Path.home() / "fiscalizarte_export"
COOKIES_PATH = WORKDIR / "cookies.txt"
CLASSIFIED_CSV_PATH = WORKDIR / "fiscalizarte_posts_clasificados.csv"
FINAL_CSV_PATH = WORKDIR / "fiscalizarte_posts_final.csv"
FINAL_XLSX_PATH = WORKDIR / "fiscalizarte_posts_final_resumen.xlsx"
FINAL_MD_INDEX_PATH = WORKDIR / "fiscalizarte_posts_final_index.md"
POSTS_DIR = WORKDIR / "posts_texto"
JSON_DIR = WORKDIR / "posts_json"
LOG_PATH = WORKDIR / "log_extraccion_final.txt"


def log(msg: str) -> None:
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(msg.rstrip() + "\n")
    print(msg)


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


def slugify(url: str) -> str:
    slug = url.rstrip("/").split("/")[-1]
    slug = re.sub(r"[^a-zA-Z0-9_-]+", "_", slug)
    return slug or "post"


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
    r = session.get(url, timeout=60)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, list):
        return " ".join((item.get("text") or "").strip() for item in data if isinstance(item, dict) and item.get("text"))
    return ""


def asr_from_video_url(src: str, cache_name: str, model: WhisperModel) -> str:
    out = WORKDIR / cache_name
    if not out.exists():
        with requests.get(src, stream=True, timeout=120) as r:
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
    POSTS_DIR.mkdir(parents=True, exist_ok=True)
    JSON_DIR.mkdir(parents=True, exist_ok=True)
    LOG_PATH.write_text("", encoding="utf-8")

    df = pd.read_csv(CLASSIFIED_CSV_PATH)
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
        "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
        "Accept": "application/json, text/plain, */*",
    })
    load_netscape_cookies(COOKIES_PATH, session)
    model = WhisperModel("tiny", device="cpu", compute_type="int8")

    rows = []
    total = len(df)
    for idx, item in enumerate(df.to_dict(orient="records"), 1):
        url = item["url"]
        tipo = item["tipo_clasificado"]
        log(f"[{idx}/{total}] {tipo} -> {url}")
        body_text = ""
        transcript_text = ""
        transcript_source = "none"
        estado = "OK"
        notas = ""
        media_ids = item.get("media_upload_ids") or ""
        duracion_media = item.get("duracion_media") or ""
        titulo = item.get("titulo") or ""
        subtitulo = item.get("subtitulo") or ""
        fecha = item.get("fecha") or ""
        audience = item.get("audience") or ""
        post_type = item.get("post_type") or ""
        metodo = item.get("metodo_extraccion_previsto") or ""

        try:
            html = session.get(url, timeout=60).text
            data = extract_preloads_json(html)
            post = data.get("post", {})
            body_html = post.get("body_html") or ""
            body_text = html_to_text(body_html)

            if tipo == "video_o_podcast_principal_con_transcripcion_substack":
                vu = post.get("videoUpload") or {}
                tr = (vu.get("extractedAudio") or {}).get("transcription") or vu.get("transcription") or {}
                if tr.get("cdn_url"):
                    transcript_text = fetch_json_text(session, tr["cdn_url"])
                    transcript_source = "substack_video"
                else:
                    estado = "ERROR"
                    notas = "No apareció cdn_url de transcript Substack"
            elif tipo == "podcast_audio_con_transcripcion_substack":
                pu = post.get("podcastUpload") or {}
                tr = pu.get("transcription") or {}
                if tr.get("cdn_url"):
                    transcript_text = fetch_json_text(session, tr["cdn_url"])
                    transcript_source = "substack_podcast"
                else:
                    estado = "ERROR"
                    notas = "No apareció cdn_url de transcript podcast"
            elif tipo == "newsletter_con_video_embebido":
                media_list = parse_embed_media_ids(body_html, "native-video-embed")
                if media_list:
                    mid = media_list[0]
                    src = session.get(f"{BASE_URL}/api/v1/video/upload/{mid}/src.json?type=original", timeout=60).json()["src"]
                    transcript_text = asr_from_video_url(src, f"asr_{mid}.mp4", model)
                    transcript_source = "local_asr_video"
                else:
                    estado = "ERROR"
                    notas = "No se encontró mediaUploadId en native-video-embed"
            elif tipo == "podcast_video_sin_transcripcion_substack":
                vu = post.get("videoUpload") or {}
                if vu.get("id"):
                    src = session.get(f"{BASE_URL}/api/v1/video/upload/{vu['id']}/src.json?type=original", timeout=60).json()["src"]
                    transcript_text = asr_from_video_url(src, f"asr_{vu['id']}.mp4", model)
                    transcript_source = "local_asr_video"
                else:
                    estado = "ERROR"
                    notas = "No se encontró videoUpload.id"
            elif tipo == "texto_normal":
                transcript_source = "none"
            else:
                estado = "ERROR"
                notas = f"Tipo no contemplado: {tipo}"

        except Exception as e:
            estado = "ERROR"
            notas = str(e)
            log(f"ERROR en {url}: {e}")

        combined_text = (body_text + "\n\n" + transcript_text).strip() if transcript_text else body_text.strip()
        row = {
            "titulo": titulo,
            "subtitulo": subtitulo,
            "fecha": fecha,
            "url": url,
            "audience": audience,
            "post_type": post_type,
            "tipo_clasificado": tipo,
            "metodo_extraccion_usado": metodo,
            "body_text": body_text,
            "body_length": len(body_text),
            "transcript_text": transcript_text,
            "transcript_length": len(transcript_text),
            "transcript_source": transcript_source,
            "combined_text": combined_text,
            "combined_length": len(combined_text),
            "media_upload_ids": media_ids,
            "duracion_media": duracion_media,
            "estado_extraccion": estado,
            "notas": notas,
        }
        rows.append(row)

        slug = slugify(url)
        (POSTS_DIR / f"{slug}.md").write_text(
            "# " + (titulo or slug) + "\n\n"
            + f"- URL: {url}\n"
            + f"- tipo_clasificado: {tipo}\n"
            + f"- transcript_source: {transcript_source}\n"
            + f"- estado_extraccion: {estado}\n\n"
            + "## body_text\n\n" + (body_text or "") + "\n\n"
            + "## transcript_text\n\n" + (transcript_text or "") + "\n",
            encoding="utf-8",
        )
        (JSON_DIR / f"{slug}.json").write_text(json.dumps(row, ensure_ascii=False, indent=2), encoding="utf-8")

    full_df = pd.DataFrame(rows)
    full_df.to_csv(FINAL_CSV_PATH, index=False, encoding="utf-8-sig")

    summary_df = full_df[[
        "titulo", "fecha", "url", "tipo_clasificado", "metodo_extraccion_usado",
        "body_length", "transcript_length", "transcript_source", "combined_length",
        "estado_extraccion", "notas"
    ]].copy()
    summary_df.to_excel(FINAL_XLSX_PATH, index=False)

    with FINAL_MD_INDEX_PATH.open("w", encoding="utf-8") as f:
        f.write("# Fiscalizarte extracción final\n\n")
        for row in rows:
            slug = slugify(row["url"])
            f.write(f"- [{row['titulo'] or slug}](posts_texto/{slug}.md) | {row['tipo_clasificado']} | {row['estado_extraccion']}\n")

    counts = full_df["estado_extraccion"].value_counts().to_dict()
    log(f"Resumen final: {counts}")
    log(f"Salidas: {FINAL_CSV_PATH} | {FINAL_XLSX_PATH} | {FINAL_MD_INDEX_PATH} | {POSTS_DIR} | {JSON_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
