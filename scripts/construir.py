#!/usr/bin/env python3
"""
Convierte las presentaciones de `fuentes/` en un sitio estático dentro de `sitio/`.

Para cada archivo fuentes/<nombre>.pptx (o .pdf) produce:
  sitio/<slug>/index.html      visor web (flechas, pantalla completa, transiciones)
  sitio/<slug>/diapositivas/   una imagen JPG por diapositiva (ancho 1920 px)
  sitio/<slug>/<slug>.pdf      el PDF de la presentación
  sitio/<slug>/<slug>.pptx     copia del original para descargar
y al final sitio/index.html con la lista de todas.

Los metadatos opcionales (título, fecha, autor, descripción, enlaces) se leen de
`presentaciones.yml`, indexados por el nombre del archivo sin extensión.

Uso: python3 scripts/construir.py [--solo-nuevas]
"""
import argparse
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from datetime import date
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parent.parent
FUENTES = RAIZ / "fuentes"
SITIO = RAIZ / "sitio"
PLANTILLAS = RAIZ / "plantillas"
METADATOS = RAIZ / "presentaciones.yml"
ANCHO = 1920          # ancho de cada diapositiva en píxeles
CALIDAD_JPG = 90      # 90 es visualmente indistinguible del original y pesa un tercio


def slugificar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    texto = re.sub(r"[^a-zA-Z0-9]+", "-", texto).strip("-").lower()
    return texto or "presentacion"


def hash_archivo(ruta: Path) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()[:16]


def cargar_metadatos() -> dict:
    if METADATOS.exists():
        with open(METADATOS, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    return {}


def pptx_a_pdf(pptx: Path, destino: Path) -> Path:
    """Convierte con LibreOffice en modo headless. Devuelve la ruta del PDF."""
    perfil = Path("/tmp/lo-perfil")
    subprocess.run(
        [
            "soffice", "--headless", "--norestore",
            f"-env:UserInstallation=file://{perfil}",
            "--convert-to", "pdf", "--outdir", str(destino), str(pptx),
        ],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600,
    )
    pdf = destino / (pptx.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice no produjo el PDF de {pptx.name}")
    return pdf


def alto_para(pdf: Path) -> int:
    """Alto en píxeles que corresponde a ANCHO según la proporción de la página (1080 para 16:9)."""
    salida = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, check=True).stdout
    m = re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", salida)
    if not m:
        return -1
    ancho_pt, alto_pt = float(m.group(1)), float(m.group(2))
    return round(ANCHO * alto_pt / ancho_pt)


def pdf_a_imagenes(pdf: Path, carpeta: Path) -> list[str]:
    carpeta.mkdir(parents=True, exist_ok=True)
    for viejo in carpeta.glob("*.jpg"):
        viejo.unlink()
    subprocess.run(
        [
            "pdftoppm", "-jpeg", "-jpegopt", f"quality={CALIDAD_JPG}",
            "-scale-to-x", str(ANCHO), "-scale-to-y", str(alto_para(pdf)),
            str(pdf), str(carpeta / "d"),
        ],
        check=True, timeout=600,
    )
    imagenes = sorted(carpeta.glob("d-*.jpg"))
    # Renombrar a 001.jpg, 002.jpg... para que el visor las encuentre sin listar
    nombres = []
    for i, img in enumerate(imagenes, 1):
        nuevo = carpeta / f"{i:03d}.jpg"
        img.rename(nuevo)
        nombres.append(nuevo.name)
    return nombres


def renderizar(plantilla: str, **campos) -> str:
    texto = (PLANTILLAS / plantilla).read_text(encoding="utf-8")
    for clave, valor in campos.items():
        texto = texto.replace("{{" + clave + "}}", str(valor))
    return texto


def fecha_legible(valor) -> str:
    if isinstance(valor, date):
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                 "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        return f"{valor.day} de {meses[valor.month - 1]} de {valor.year}"
    return str(valor) if valor else ""


def construir_una(fuente: Path, meta: dict, solo_nuevas: bool) -> dict:
    slug = meta.get("slug") or slugificar(fuente.stem)
    carpeta = SITIO / slug
    carpeta.mkdir(parents=True, exist_ok=True)
    marca = carpeta / ".fuente.json"
    firma = hash_archivo(fuente)

    titulo = meta.get("titulo") or fuente.stem.replace("_", " ").replace("-", " ")
    info = {
        "slug": slug,
        "titulo": titulo,
        "autor": meta.get("autor", ""),
        "fecha": meta.get("fecha"),
        "descripcion": meta.get("descripcion", ""),
        "enlaces": meta.get("enlaces", []),
        "archivo": fuente.name,
    }

    previa = json.loads(marca.read_text()) if marca.exists() else {}
    if solo_nuevas and previa.get("firma") == firma and (carpeta / "diapositivas").exists():
        print(f"  = {fuente.name}: sin cambios")
        info["n"] = previa.get("n", 0)
    else:
        print(f"  > {fuente.name}: convirtiendo…")
        if fuente.suffix.lower() == ".pdf":
            pdf = carpeta / f"{slug}.pdf"
            shutil.copy2(fuente, pdf)
        else:
            pdf_tmp = pptx_a_pdf(fuente, carpeta)
            pdf = carpeta / f"{slug}.pdf"
            if pdf_tmp != pdf:
                pdf_tmp.replace(pdf)
            shutil.copy2(fuente, carpeta / f"{slug}{fuente.suffix.lower()}")
        nombres = pdf_a_imagenes(pdf, carpeta / "diapositivas")
        info["n"] = len(nombres)
        marca.write_text(json.dumps({"firma": firma, "n": info["n"]}))
        print(f"    {info['n']} diapositivas")

    descargable = f"{slug}{fuente.suffix.lower()}" if fuente.suffix.lower() != ".pdf" else ""
    enlaces_html = "".join(
        f'<li><a href="{html.escape(e["url"])}" target="_blank" rel="noopener">{html.escape(e.get("texto", e["url"]))}</a></li>'
        for e in info["enlaces"]
    )
    (carpeta / "index.html").write_text(
        renderizar(
            "presentacion.html",
            titulo=html.escape(titulo),
            autor=html.escape(info["autor"]),
            fecha=html.escape(fecha_legible(info["fecha"])),
            descripcion=html.escape(info["descripcion"]),
            n=info["n"],
            pdf=f"{slug}.pdf",
            pptx=descargable,
            boton_pptx=(f'<a class="boton" href="{descargable}" download>Descargar PowerPoint</a>' if descargable else ""),
            enlaces=(f"<h2>Enlaces</h2><ul>{enlaces_html}</ul>" if enlaces_html else ""),
        ),
        encoding="utf-8",
    )
    return info


def construir_indice(presentaciones: list[dict], config: dict):
    def clave(p):
        f = p.get("fecha")
        return (f.isoformat() if isinstance(f, date) else str(f or "")), p["titulo"]

    presentaciones.sort(key=clave, reverse=True)
    tarjetas = []
    for p in presentaciones:
        portada = f'{p["slug"]}/diapositivas/001.jpg'
        detalles = " · ".join(x for x in [fecha_legible(p["fecha"]), p["autor"], f'{p["n"]} diapositivas'] if x)
        tarjetas.append(f"""
      <article class="tarjeta">
        <a href="{p['slug']}/" class="portada"><img src="{portada}" alt="" loading="lazy"></a>
        <div class="texto">
          <h2><a href="{p['slug']}/">{html.escape(p['titulo'])}</a></h2>
          <p class="detalles">{html.escape(detalles)}</p>
          {f'<p>{html.escape(p["descripcion"])}</p>' if p['descripcion'] else ''}
          <p class="acciones">
            <a class="boton primario" href="{p['slug']}/visor.html">Presentar</a>
            <a class="boton" href="{p['slug']}/{p['slug']}.pdf">PDF</a>
          </p>
        </div>
      </article>""")
    (SITIO / "index.html").write_text(
        renderizar(
            "indice.html",
            titulo=html.escape(config.get("titulo", "Presentaciones")),
            subtitulo=html.escape(config.get("subtitulo", "")),
            pie=config.get("pie", ""),
            tarjetas="\n".join(tarjetas) or "<p>Aún no hay presentaciones.</p>",
            total=len(presentaciones),
        ),
        encoding="utf-8",
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--solo-nuevas", action="store_true",
                    help="no reconvertir las presentaciones cuyo archivo no cambió")
    args = ap.parse_args()

    SITIO.mkdir(exist_ok=True)
    metadatos = cargar_metadatos()
    config = metadatos.pop("_sitio", {}) or {}

    fuentes = sorted(p for p in FUENTES.iterdir()
                     if p.suffix.lower() in (".pptx", ".ppt", ".odp", ".pdf") and not p.name.startswith("."))
    if not fuentes:
        print("No hay archivos en fuentes/")
    presentaciones = []
    for fuente in fuentes:
        meta = metadatos.get(fuente.stem, {}) or {}
        presentaciones.append(construir_una(fuente, meta, args.solo_nuevas))

    # Borrar carpetas del sitio cuya fuente ya no existe
    vivos = {p["slug"] for p in presentaciones}
    for carpeta in SITIO.iterdir():
        if carpeta.is_dir() and carpeta.name not in vivos:
            shutil.rmtree(carpeta)
            print(f"  - {carpeta.name}: fuente eliminada, carpeta borrada")

    # Visor compartido (una sola copia por presentación, apunta a sus imágenes)
    for p in presentaciones:
        shutil.copy2(PLANTILLAS / "visor.html", SITIO / p["slug"] / "visor.html")
        (SITIO / p["slug"] / "visor.json").write_text(
            json.dumps({"titulo": p["titulo"], "n": p["n"]}, ensure_ascii=False)
        )
    shutil.copy2(PLANTILLAS / "estilos.css", SITIO / "estilos.css")
    (SITIO / ".nojekyll").touch()
    construir_indice(presentaciones, config)
    print(f"Listo: {len(presentaciones)} presentaciones en {SITIO}")


if __name__ == "__main__":
    main()
