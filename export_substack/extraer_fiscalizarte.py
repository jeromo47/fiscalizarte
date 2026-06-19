#!/usr/bin/env python3
from __future__ import annotations

import os
import sys
from dataclasses import dataclass, asdict
from html import unescape
from pathlib import Path
from typing import Any
import json
import re

import pandas as pd
import requests
import trafilatura
from bs4 import BeautifulSoup

BASE_URL = "https://fiscalizarte.substack.com"
ARCHIVE_URL = BASE_URL + "/api/v1/archive?sort=new&search=&offset={offset}&limit={limit}"
WORKDIR = Path.home() / "fiscalizarte_export"
COOKIES_PATH = WORKDIR / "cookies.txt"
CSV_PATH = WORKDIR / "fiscalizarte_posts.csv"
XLSX_PATH = WORKDIR / "fiscalizarte_posts.xlsx"
MD_PATH = WORKDIR / "fiscalizarte_posts.md"
LOG_PATH = WORKDIR / "log_extraccion.txt"
CLASSIFIED_CSV_PATH = WORKDIR / "fiscalizarte_posts_clasificados.csv"
CLASSIFIED_XLSX_PATH = WORKDIR / "fiscalizarte_posts_clasificados.xlsx"
CLASSIFIED_MD_PATH = WORKDIR / "fiscalizarte_posts_clasificados.md"
FINAL_SCHEMA_CSV_PATH = WORKDIR / "fiscalizarte_posts_final_schema.csv"
FINAL_SCHEMA_XLSX_PATH = WORKDIR / "fiscalizarte_posts_final_schema.xlsx"
FINAL_SCHEMA_MD_PATH = WORKDIR / "fiscalizarte_posts_final_schema.md"
LIMIT = 12


@dataclass
class PostRow:
    titulo: str
    subtitulo: str
    fecha: str
    url: str
    tipo_audiencia: str
    longitud_texto: int
    texto_completo: str
    estado: str


@dataclass
class ClassifiedPostRow:
    titulo: str
    subtitulo: str
    fecha: str
    url: str
    audience: str
    post_type: str
    tipo_clasificado: str
    metodo_extraccion_previsto: str
    has_body_html: bool
    has_video_upload: bool
    has_podcast_upload: bool
    has_native_video_embed: bool
    has_native_audio_embed: bool
    has_substack_transcript: bool
    transcript_source_candidate: str
    media_upload_ids: str
    duracion_media: str
    estado_revision: str


@dataclass
class FinalSchemaRow:
    titulo: str
    subtitulo: str
    fecha: str
    url: str
    audience: str
    post_type: str
    tipo_clasificado: str
    metodo_extraccion_usado: str
    body_text: str
    body_length: int
    transcript_text: str
    transcript_length: int
    transcript_source: str
    combined_text: str
    combined_length: int
    media_upload_ids: str
    duracion_media: str
    estado_extraccion: str
    notas: str


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
        session.cookies.set(
            name,
            value,
            domain=domain,
            path=cookie_path or "/",
            secure=(secure.upper() == "TRUE"),
        )


def build_url(post: dict[str, Any]) -> str:
    for key in ("canonical_url", "canonicalUrl", "url"):
        val = post.get(key)
        if isinstance(val, str) and val.startswith("http"):
            return val
    slug = post.get("slug") or post.get("id") or ""
    return f"{BASE_URL}/p/{slug}"


def extract_with_bs4(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    article = soup.find("article")
    if article:
        return article.get_text("\n", strip=True)
    body = soup.find("body")
    if body:
        return body.get_text("\n", strip=True)
    return soup.get_text("\n", strip=True)


def fetch_archive(session: requests.Session) -> list[dict[str, Any]]:
    posts: list[dict[str, Any]] = []
    offset = 0
    while True:
        url = ARCHIVE_URL.format(offset=offset, limit=LIMIT)
        log(f"Consultando archivo: {url}")
        r = session.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()
        if not data or not isinstance(data, list):
            break
        posts.extend(data)
        if len(data) < LIMIT:
            break
        offset += LIMIT
    return posts


def export_rows(rows: list[PostRow]) -> None:
    df = pd.DataFrame([asdict(r) for r in rows])
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    df.to_excel(XLSX_PATH, index=False)
    with MD_PATH.open("w", encoding="utf-8") as f:
        f.write("# Fiscalizarte export\n\n")
        for i, row in enumerate(rows, 1):
            f.write(f"## {i}. {row.titulo or '(sin título)'}\n\n")
            f.write(f"- Fecha: {row.fecha}\n")
            f.write(f"- URL: {row.url}\n")
            f.write(f"- Tipo/audiencia: {row.tipo_audiencia}\n")
            f.write(f"- Estado: {row.estado}\n")
            f.write(f"- Longitud: {row.longitud_texto}\n\n")
            if row.subtitulo:
                f.write(f"**Subtítulo:** {row.subtitulo}\n\n")
            f.write(row.texto_completo + "\n\n---\n\n")


def export_classified_rows(rows: list[ClassifiedPostRow]) -> None:
    df = pd.DataFrame([asdict(r) for r in rows])
    df.to_csv(CLASSIFIED_CSV_PATH, index=False, encoding="utf-8-sig")
    df.to_excel(CLASSIFIED_XLSX_PATH, index=False)
    with CLASSIFIED_MD_PATH.open("w", encoding="utf-8") as f:
        f.write("# Fiscalizarte inventario clasificado\n\n")
        for i, row in enumerate(rows, 1):
            f.write(f"## {i}. {row.titulo or '(sin título)'}\n\n")
            f.write(f"- Fecha: {row.fecha}\n")
            f.write(f"- URL: {row.url}\n")
            f.write(f"- audience: {row.audience}\n")
            f.write(f"- post_type: {row.post_type}\n")
            f.write(f"- tipo_clasificado: {row.tipo_clasificado}\n")
            f.write(f"- metodo_extraccion_previsto: {row.metodo_extraccion_previsto}\n")
            f.write(f"- transcript_source_candidate: {row.transcript_source_candidate}\n")
            f.write(f"- media_upload_ids: {row.media_upload_ids}\n")
            f.write(f"- duracion_media: {row.duracion_media}\n")
            f.write(f"- estado_revision: {row.estado_revision}\n\n")


def export_final_schema_rows(rows: list[FinalSchemaRow]) -> None:
    df = pd.DataFrame([asdict(r) for r in rows])
    df.to_csv(FINAL_SCHEMA_CSV_PATH, index=False, encoding="utf-8-sig")
    df.to_excel(FINAL_SCHEMA_XLSX_PATH, index=False)
    with FINAL_SCHEMA_MD_PATH.open("w", encoding="utf-8") as f:
        f.write("# Fiscalizarte esquema final de extracción\n\n")
        for i, row in enumerate(rows, 1):
            f.write(f"## {i}. {row.titulo or '(sin título)'}\n\n")
            f.write(f"- Fecha: {row.fecha}\n")
            f.write(f"- URL: {row.url}\n")
            f.write(f"- tipo_clasificado: {row.tipo_clasificado}\n")
            f.write(f"- metodo_extraccion_usado: {row.metodo_extraccion_usado}\n")
            f.write(f"- transcript_source: {row.transcript_source}\n")
            f.write(f"- media_upload_ids: {row.media_upload_ids}\n")
            f.write(f"- duracion_media: {row.duracion_media}\n")
            f.write(f"- estado_extraccion: {row.estado_extraccion}\n")
            f.write(f"- notas: {row.notas}\n\n")


def extract_preloads_json(html: str) -> dict[str, Any]:
    m = re.search(r'window\._preloads\s*=\s*JSON\.parse\((".*?")\)\s*</script>', html, re.S)
    if not m:
        m = re.search(r'window\._preloads\s*=\s*JSON\.parse\((".*?")\)', html, re.S)
    if not m:
        return {}
    try:
        inner = json.loads(m.group(1))
        return json.loads(inner)
    except Exception:
        return {}


def parse_embed_media_ids(body_html: str, class_name: str) -> list[str]:
    if not body_html:
        return []
    pattern = rf'<div class="{re.escape(class_name)}"[^>]*data-attrs="([^"]+)"'
    found = []
    for attrs in re.findall(pattern, body_html):
        try:
            decoded = unescape(attrs)
            data = json.loads(decoded)
            media_id = data.get("mediaUploadId")
            if media_id:
                found.append(str(media_id))
        except Exception:
            continue
    return found


def classify_post_from_preloads(data: dict[str, Any]) -> ClassifiedPostRow:
    post = data.get("post", {}) if isinstance(data, dict) else {}
    title = str(post.get("title") or "")
    subtitle = str(post.get("subtitle") or post.get("description") or "")
    date = str(post.get("post_date") or post.get("created_at") or post.get("published_at") or "")
    url = build_url(post)
    audience = str(post.get("audience") or "")
    post_type = str(post.get("type") or "")
    body_html = str(post.get("body_html") or "")
    has_body_html = bool(body_html)
    video_upload = post.get("videoUpload") or {}
    podcast_upload = post.get("podcastUpload") or {}
    extracted_audio = video_upload.get("extractedAudio") or {}
    video_transcription = extracted_audio.get("transcription") or video_upload.get("transcription") or {}
    podcast_transcription = podcast_upload.get("transcription") or {}
    native_video_ids = parse_embed_media_ids(body_html, "native-video-embed")
    native_audio_ids = parse_embed_media_ids(body_html, "native-audio-embed")
    has_video_upload = bool(video_upload)
    has_podcast_upload = bool(podcast_upload)
    has_native_video_embed = bool(native_video_ids)
    has_native_audio_embed = bool(native_audio_ids)
    has_substack_transcript = bool(video_transcription or podcast_transcription)

    transcript_source_candidate = ""
    tipo_clasificado = "otro"
    metodo = "revision_manual"
    media_ids: list[str] = []
    durations: list[str] = []

    if has_video_upload:
        media_ids.append(str(video_upload.get("id") or post.get("video_upload_id") or ""))
        if video_upload.get("duration"):
            durations.append(str(video_upload.get("duration")))
    if has_podcast_upload:
        media_ids.append(str(podcast_upload.get("id") or post.get("podcast_upload_id") or ""))
        if podcast_upload.get("duration"):
            durations.append(str(podcast_upload.get("duration")))
    media_ids.extend(native_video_ids)
    media_ids.extend(native_audio_ids)

    if has_body_html and not has_video_upload and not has_podcast_upload and not has_native_video_embed and not has_native_audio_embed:
        tipo_clasificado = "texto_normal"
        metodo = "body_html"
        transcript_source_candidate = "body_html"
    elif has_video_upload and has_substack_transcript:
        tipo_clasificado = "video_o_podcast_principal_con_transcripcion_substack"
        metodo = "body_html + videoUpload.extractedAudio.transcription"
        transcript_source_candidate = "videoUpload.extractedAudio.transcription"
    elif has_podcast_upload and has_substack_transcript:
        tipo_clasificado = "podcast_audio_con_transcripcion_substack"
        metodo = "body_html + podcastUpload.transcription"
        transcript_source_candidate = "podcastUpload.transcription"
    elif has_video_upload and has_podcast_upload and not has_substack_transcript:
        tipo_clasificado = "podcast_video_sin_transcripcion_substack"
        metodo = "body_html + videoUpload + ASR_local"
        transcript_source_candidate = "video_upload_local_asr"
    elif has_native_video_embed:
        tipo_clasificado = "newsletter_con_video_embebido"
        metodo = "body_html + native-video-embed + ASR_local"
        transcript_source_candidate = "native_video_embed_local_asr"
    elif has_native_audio_embed:
        tipo_clasificado = "newsletter_con_audio_embebido"
        metodo = "body_html + native-audio-embed"
        transcript_source_candidate = "native_audio_embed"

    estado_revision = "REVISAR" if tipo_clasificado == "otro" else "OK"

    return ClassifiedPostRow(
        titulo=title,
        subtitulo=subtitle,
        fecha=date,
        url=url,
        audience=audience,
        post_type=post_type,
        tipo_clasificado=tipo_clasificado,
        metodo_extraccion_previsto=metodo,
        has_body_html=has_body_html,
        has_video_upload=has_video_upload,
        has_podcast_upload=has_podcast_upload,
        has_native_video_embed=has_native_video_embed,
        has_native_audio_embed=has_native_audio_embed,
        has_substack_transcript=has_substack_transcript,
        transcript_source_candidate=transcript_source_candidate,
        media_upload_ids=" | ".join(x for x in media_ids if x),
        duracion_media=" | ".join(durations),
        estado_revision=estado_revision,
    )


def build_final_schema_row(classified: ClassifiedPostRow) -> FinalSchemaRow:
    return FinalSchemaRow(
        titulo=classified.titulo,
        subtitulo=classified.subtitulo,
        fecha=classified.fecha,
        url=classified.url,
        audience=classified.audience,
        post_type=classified.post_type,
        tipo_clasificado=classified.tipo_clasificado,
        metodo_extraccion_usado=classified.metodo_extraccion_previsto,
        body_text="",
        body_length=0,
        transcript_text="",
        transcript_length=0,
        transcript_source=(
            "substack_video" if classified.tipo_clasificado == "video_o_podcast_principal_con_transcripcion_substack"
            else "substack_podcast" if classified.tipo_clasificado == "podcast_audio_con_transcripcion_substack"
            else "local_asr_video" if classified.tipo_clasificado in {"newsletter_con_video_embebido", "podcast_video_sin_transcripcion_substack"}
            else "none"
        ),
        combined_text="",
        combined_length=0,
        media_upload_ids=classified.media_upload_ids,
        duracion_media=classified.duracion_media,
        estado_extraccion="PENDIENTE",
        notas="Preparado para extracción final",
    )


def main() -> int:
    WORKDIR.mkdir(parents=True, exist_ok=True)
    LOG_PATH.write_text("", encoding="utf-8")

    if not COOKIES_PATH.exists():
        log("No existe cookies.txt en la carpeta de trabajo.")
        return 2

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
        "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    })
    load_netscape_cookies(COOKIES_PATH, session)

    try:
        posts = fetch_archive(session)
    except Exception as e:
        log(f"ERROR consultando archive autenticado: {e}")
        return 1

    rows: list[PostRow] = []
    classified_rows: list[ClassifiedPostRow] = []
    final_schema_rows: list[FinalSchemaRow] = []
    log(f"Posts detectados: {len(posts)}")
    for idx, post in enumerate(posts, 1):
        url = build_url(post)
        titulo = post.get("title") or ""
        subtitulo = post.get("subtitle") or post.get("description") or ""
        fecha = str(post.get("post_date") or post.get("created_at") or post.get("published_at") or "")
        tipo = str(post.get("audience") or post.get("type") or post.get("visibility") or "")
        estado = "OK"
        texto = ""
        log(f"[{idx}/{len(posts)}] Extrayendo {url}")
        try:
            resp = session.get(url, timeout=30)
            html = resp.text
            if resp.status_code >= 400:
                estado = "ERROR"
            downloaded = trafilatura.extract(html, include_comments=False, include_tables=False)
            texto = downloaded.strip() if downloaded else extract_with_bs4(html).strip()
            low = html.lower()
            if not texto:
                estado = "ERROR"
            elif len(texto) < 1200 and any(marker in low for marker in ["subscribe to read", "paid subscribers", "become a subscriber", "this post is for subscribers"]):
                estado = "POSIBLE_PAYWALL_O_NO_ACCESO"
            else:
                estado = "OK"
            preloads = extract_preloads_json(html)
            if preloads:
                classified = classify_post_from_preloads(preloads)
                classified_rows.append(classified)
                final_schema_rows.append(build_final_schema_row(classified))
            else:
                classified = ClassifiedPostRow(
                    titulo=titulo,
                    subtitulo=subtitulo,
                    fecha=fecha,
                    url=url,
                    audience=str(post.get("audience") or ""),
                    post_type=str(post.get("type") or ""),
                    tipo_clasificado="sin_preloads",
                    metodo_extraccion_previsto="revision_manual",
                    has_body_html=False,
                    has_video_upload=False,
                    has_podcast_upload=False,
                    has_native_video_embed=False,
                    has_native_audio_embed=False,
                    has_substack_transcript=False,
                    transcript_source_candidate="",
                    media_upload_ids="",
                    duracion_media="",
                    estado_revision="REVISAR",
                )
                classified_rows.append(classified)
                final_schema_rows.append(build_final_schema_row(classified))
        except Exception as e:
            estado = "ERROR"
            texto = f"ERROR: {e}"
            classified = ClassifiedPostRow(
                titulo=titulo,
                subtitulo=subtitulo,
                fecha=fecha,
                url=url,
                audience=str(post.get("audience") or ""),
                post_type=str(post.get("type") or ""),
                tipo_clasificado="error",
                metodo_extraccion_previsto="revision_manual",
                has_body_html=False,
                has_video_upload=False,
                has_podcast_upload=False,
                has_native_video_embed=False,
                has_native_audio_embed=False,
                has_substack_transcript=False,
                transcript_source_candidate="",
                media_upload_ids="",
                duracion_media="",
                estado_revision="REVISAR",
            )
            classified_rows.append(classified)
            final_schema_rows.append(build_final_schema_row(classified))
            log(f"Error en {url}: {e}")

        rows.append(PostRow(
            titulo=titulo,
            subtitulo=subtitulo,
            fecha=fecha,
            url=url,
            tipo_audiencia=tipo,
            longitud_texto=len(texto),
            texto_completo=texto,
            estado=estado,
        ))

    export_rows(rows)
    export_classified_rows(classified_rows)
    export_final_schema_rows(final_schema_rows)
    total = len(rows)
    ok = sum(1 for r in rows if r.estado == "OK")
    pay = sum(1 for r in rows if r.estado == "POSIBLE_PAYWALL_O_NO_ACCESO")
    log(f"Resumen: total={total}, OK={ok}, POSIBLE_PAYWALL_O_NO_ACCESO={pay}")
    counts = pd.Series([r.tipo_clasificado for r in classified_rows]).value_counts().to_dict() if classified_rows else {}
    log(f"Resumen clasificación: {counts}")
    log(f"Archivos: {CSV_PATH} | {XLSX_PATH} | {MD_PATH} | {CLASSIFIED_CSV_PATH} | {CLASSIFIED_XLSX_PATH} | {CLASSIFIED_MD_PATH} | {FINAL_SCHEMA_CSV_PATH} | {FINAL_SCHEMA_XLSX_PATH} | {FINAL_SCHEMA_MD_PATH} | {LOG_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
