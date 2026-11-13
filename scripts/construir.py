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
    if (carpeta / "m").exists():
        shutil.rmtree(carpeta / "m")  # miniaturas viejas: se regeneran después
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


ANCHO_MINI = 320  # ancho de las miniaturas del menú lateral


def generar_miniaturas(carpeta: Path, forzar: bool = False) -> None:
    """Crea diapositivas/m/NNN.jpg (miniaturas) a partir de las imágenes grandes."""
    from PIL import Image
    mini = carpeta / "m"
    mini.mkdir(exist_ok=True)
    for img in sorted(carpeta.glob("[0-9][0-9][0-9].jpg")):
        destino = mini / img.name
        if destino.exists() and not forzar:
            continue
        with Image.open(img) as im:
            alto = round(im.height * ANCHO_MINI / im.width)
            im.convert("RGB").resize((ANCHO_MINI, alto), Image.LANCZOS).save(destino, "JPEG", quality=80)
    for sobrante in mini.glob("*.jpg"):
        if not (carpeta / sobrante.name).exists():
            sobrante.unlink()


def leer_secciones(pptx: Path) -> list[dict]:
    """Secciones de PowerPoint (clic derecho > Agregar sección) como [{titulo, desde}] con desde en base 1."""
    import zipfile
    try:
        with zipfile.ZipFile(pptx) as z:
            xml = z.read("ppt/presentation.xml").decode("utf-8")
    except Exception:
        return []
    lista = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", xml, re.S)
    if not lista:
        return []
    orden = {sid: i + 1 for i, sid in enumerate(re.findall(r'<p:sldId [^>]*id="(\d+)"', lista.group(1)))}
    secciones = []
    for nombre, cuerpo in re.findall(r'<p14:section [^>]*name="([^"]*)"[^>]*>(.*?)</p14:section>', xml, re.S):
        ids = [orden[s] for s in re.findall(r'<p14:sldId id="(\d+)"', cuerpo) if s in orden]
        if ids:
            secciones.append({"titulo": html.unescape(nombre), "desde": min(ids)})
    secciones.sort(key=lambda s: s["desde"])
    # Una sola sección por defecto no aporta nada
    if len(secciones) == 1 and secciones[0]["desde"] == 1 and secciones[0]["titulo"].lower() in ("sección predeterminada", "default section"):
        return []
    return secciones


def renderizar(plantilla: str, **campos) -> str:
    texto = (PLANTILLAS / plantilla).read_text(encoding="utf-8")
    campos["vcss"] = hash_archivo(PLANTILLAS / "estilos.css")[:8]
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
        "categoria": meta.get("categoria", "") or "",
        "archivo": fuente.name,
    }

    info["v"] = firma[:10]
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

    generar_miniaturas(carpeta / "diapositivas")
    # Secciones: las de presentaciones.yml mandan; si no hay, las de PowerPoint
    info["secciones"] = meta.get("secciones") or (leer_secciones(fuente) if fuente.suffix.lower() == ".pptx" else [])

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
            v=info["v"],
            pdf=f"{slug}.pdf",
            pptx=descargable,
            boton_pptx=(f'<a class="boton" href="{descargable}" download>Descargar PowerPoint</a>' if descargable else ""),
            enlaces=(f"<h2>Enlaces</h2><ul>{enlaces_html}</ul>" if enlaces_html else ""),
        ),
        encoding="utf-8",
    )
    return info


def tarjeta_html(p: dict, prefijo: str) -> str:
    """Tarjeta de una presentación. `prefijo` es la ruta relativa hasta la raíz del sitio."""
    base = f'{prefijo}{p["slug"]}/'
    detalles = " · ".join(x for x in [fecha_legible(p["fecha"]), p["autor"], f'{p["n"]} diapositivas'] if x)
    return f"""
      <article class="tarjeta">
        <a href="{base}" class="portada"><img src="{base}diapositivas/001.jpg?v={p['v']}" alt="" loading="lazy"></a>
        <div class="texto">
          <h3><a href="{base}">{html.escape(p['titulo'])}</a></h3>
          <p class="detalles">{html.escape(detalles)}</p>
          {f'<p>{html.escape(p["descripcion"])}</p>' if p['descripcion'] else ''}
          <p class="acciones">
            <a class="boton primario" href="{base}visor.html">Presentar</a>
            <a class="boton" href="{base}{p['slug']}.pdf">PDF</a>
          </p>
        </div>
      </article>"""


def contar(carpeta: dict) -> int:
    return len(carpeta["items"]) + sum(contar(s) for s in carpeta["sub"].values())


def primera_portada(carpeta: dict):
    if carpeta["items"]:
        return carpeta["items"][0]
    for s in carpeta["sub"].values():
        p = primera_portada(s)
        if p:
            return p
    return None


def tarjeta_carpeta_html(nombre: str, carpeta: dict, prefijo: str) -> str:
    n = contar(carpeta)
    portada = primera_portada(carpeta)
    img = f'<img src="{prefijo}{portada["slug"]}/diapositivas/001.jpg?v={portada["v"]}" alt="" loading="lazy">' if portada else ""
    nsub = len(carpeta["sub"])
    detalle = f'{n} presentaci{"ón" if n == 1 else "ones"}' + (f' · {nsub} subcategor{"ía" if nsub == 1 else "ías"}' if nsub else "")
    return f"""
      <a class="carpeta" href="{slugificar(nombre)}/">
        <div class="portada">{img}</div>
        <div class="texto"><h3>{html.escape(nombre)}</h3><p class="detalles">{detalle}</p></div>
      </a>"""


def escribir_carpeta(ruta_nombres: list[str], carpeta: dict, config: dict, orden: list[str], destino: Path, prefijo: str):
    """Escribe el index.html de una carpeta (raíz, categoría o subcategoría) y recurre a sus subcarpetas."""
    destino.mkdir(parents=True, exist_ok=True)
    es_raiz = not ruta_nombres
    sin_clasificar = config.get("sin_clasificar", "Sin clasificar")

    def orden_sub(nombre):
        return (orden.index(nombre) if nombre in orden else len(orden), nombre == sin_clasificar, nombre.lower())
    subnombres = sorted(carpeta["sub"], key=orden_sub)

    migas = [f'<a href="{prefijo}">{html.escape(config.get("titulo", "Presentaciones"))}</a>']
    for i, nombre in enumerate(ruta_nombres):
        subida = "../" * (len(ruta_nombres) - 1 - i)
        migas.append(f'<a href="{subida}">{html.escape(nombre)}</a>' if i < len(ruta_nombres) - 1 else f'<span>{html.escape(nombre)}</span>')

    carpetas_html = "".join(tarjeta_carpeta_html(nombre, carpeta["sub"][nombre], prefijo) for nombre in subnombres)
    items_html = "".join(tarjeta_html(p, prefijo) for p in carpeta["items"])
    bloques = []
    if carpetas_html:
        bloques.append(f'<section class="rejilla carpetas">{carpetas_html}</section>')
    if items_html:
        if carpetas_html:
            bloques.append('<h2 class="subtitulo-seccion">Presentaciones</h2>')
        bloques.append(f'<section class="rejilla">{items_html}</section>')
    if not bloques:
        bloques.append("<p>Aún no hay presentaciones.</p>")

    titulo_pagina = ruta_nombres[-1] if ruta_nombres else config.get("titulo", "Presentaciones")
    (destino / "index.html").write_text(
        renderizar(
            "indice.html",
            titulo=html.escape(titulo_pagina),
            titulo_sitio=html.escape(config.get("titulo", "Presentaciones")),
            subtitulo=html.escape(config.get("subtitulo", "")) if es_raiz else f'{contar(carpeta)} presentaci{"ón" if contar(carpeta) == 1 else "ones"}',
            migas="" if es_raiz else '<nav class="migas">' + ' <span class="sep">›</span> '.join(migas) + "</nav>",
            prefijo=prefijo,
            contenido="\n".join(bloques),
            pie=config.get("pie", ""),
        ),
        encoding="utf-8",
    )
    for nombre in subnombres:
        escribir_carpeta(ruta_nombres + [nombre], carpeta["sub"][nombre], config, orden,
                         destino / slugificar(nombre), prefijo + "../")


def normalizar(texto: str) -> str:
    """Minúsculas sin tildes, para buscar 'cushing' o 'farmaceutica' sin importar acentos."""
    texto = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto.lower()).strip()


def buscador_html(presentaciones: list[dict]) -> str:
    """Caja de búsqueda de la portada: filtra por título, archivo, categoría, autor, descripción y fecha."""
    if not presentaciones:
        return ""
    tarjetas = []
    for p in presentaciones:
        texto = " ".join(str(x) for x in [p["titulo"], p["archivo"], Path(p["archivo"]).stem.replace("_", " "),
                                          p["categoria"], p["autor"], p["descripcion"], fecha_legible(p["fecha"])] if x)
        tarjetas.append(tarjeta_html(p, "").replace('<article class="tarjeta">',
                        f'<article class="tarjeta" data-busqueda="{html.escape(normalizar(texto))}" hidden>', 1))
    return f"""
  <div class="buscador">
    <input id="buscar" type="search" placeholder="Buscar por título, archivo, categoría o autor…" autocomplete="off" aria-label="Buscar presentaciones">
  </div>
  <p id="sinResultados" class="sin-resultados" hidden>No hay presentaciones que coincidan.</p>
  <section id="resultados" class="rejilla" hidden>{"".join(tarjetas)}</section>
  <script>
  document.addEventListener('DOMContentLoaded', () => {{
    const caja = document.getElementById('buscar');
    const resultados = document.getElementById('resultados');
    const carpetas = document.getElementById('carpetas');
    const vacio = document.getElementById('sinResultados');
    const tarjetas = [...resultados.querySelectorAll('.tarjeta')];
    const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
    function filtrar() {{
      const palabras = norm(caja.value).split(/\\s+/).filter(Boolean);
      const activo = palabras.length > 0;
      let n = 0;
      for (const t of tarjetas) {{
        const ok = activo && palabras.every(w => t.dataset.busqueda.includes(w));
        t.hidden = !ok; if (ok) n++;
      }}
      resultados.hidden = !activo || n === 0;
      vacio.hidden = !activo || n > 0;
      if (carpetas) carpetas.hidden = activo;
      history.replaceState(null, '', activo ? '#buscar=' + encodeURIComponent(caja.value) : location.pathname);
    }}
    caja.addEventListener('input', filtrar);
    caja.addEventListener('keydown', e => {{ if (e.key === 'Escape') {{ caja.value = ''; filtrar(); }} }});
    const previo = location.hash.match(/^#buscar=(.*)$/);
    if (previo) {{ caja.value = decodeURIComponent(previo[1]); filtrar(); }}
    // "/" enfoca el buscador, como en muchos sitios
    document.addEventListener('keydown', e => {{
      if (e.key === '/' && document.activeElement !== caja) {{ e.preventDefault(); caja.focus(); }}
    }});
  }});
  </script>"""


def construir_indice(presentaciones: list[dict], config: dict):
    def clave(p):
        f = p.get("fecha")
        return (f.isoformat() if isinstance(f, date) else str(f or "")), p["titulo"]

    presentaciones.sort(key=clave, reverse=True)
    sin_clasificar = config.get("sin_clasificar", "Sin clasificar")

    # Árbol de carpetas: "A / B / C" anida A > B > C. Las presentaciones sin categoría van a "Sin clasificar".
    raiz = {"sub": {}, "items": []}
    for p in presentaciones:
        partes = [x.strip() for x in p["categoria"].split("/") if x.strip()] or [sin_clasificar]
        nodo = raiz
        for parte in partes:
            nodo = nodo["sub"].setdefault(parte, {"sub": {}, "items": []})
        nodo["items"].append(p)

    # Las carpetas viven en sitio/c/<cat>/<sub>/ para no chocar con los slugs de las presentaciones
    carpeta_c = SITIO / "c"
    if carpeta_c.exists():
        shutil.rmtree(carpeta_c)
    orden = config.get("orden_categorias") or []
    # La raíz se escribe en sitio/index.html, pero sus subcarpetas cuelgan de sitio/c/
    raiz_destino = SITIO
    escribir_carpeta([], {"sub": {}, "items": raiz["items"]}, config, orden, raiz_destino, "")
    # Reescribir la raíz con enlaces a c/<cat>/
    sub_html = "".join(
        tarjeta_carpeta_html(n, raiz["sub"][n], "").replace(f'href="{slugificar(n)}/"', f'href="c/{slugificar(n)}/"')
        for n in sorted(raiz["sub"], key=lambda n: (orden.index(n) if n in orden else len(orden), n == sin_clasificar, n.lower()))
    )
    contenido_raiz = buscador_html(presentaciones) + (
        f'<section id="carpetas" class="rejilla carpetas">{sub_html}</section>' if sub_html else "<p>Aún no hay presentaciones.</p>")
    (SITIO / "index.html").write_text(
        renderizar(
            "indice.html",
            titulo=html.escape(config.get("titulo", "Presentaciones")),
            titulo_sitio=html.escape(config.get("titulo", "Presentaciones")),
            subtitulo=html.escape(config.get("subtitulo", "")),
            migas="",
            prefijo="",
            contenido=contenido_raiz,
            pie=config.get("pie", ""),
        ),
        encoding="utf-8",
    )
    for nombre in raiz["sub"]:
        escribir_carpeta([nombre], raiz["sub"][nombre], config, orden, carpeta_c / slugificar(nombre), "../../")


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
        if carpeta.is_dir() and carpeta.name not in vivos and carpeta.name != "c":
            shutil.rmtree(carpeta)
            print(f"  - {carpeta.name}: fuente eliminada, carpeta borrada")

    # Visor compartido (una sola copia por presentación, apunta a sus imágenes)
    for p in presentaciones:
        shutil.copy2(PLANTILLAS / "visor.html", SITIO / p["slug"] / "visor.html")
        (SITIO / p["slug"] / "visor.json").write_text(
            json.dumps({"titulo": p["titulo"], "n": p["n"], "v": p["v"], "secciones": p.get("secciones", [])}, ensure_ascii=False)
        )
    shutil.copy2(PLANTILLAS / "estilos.css", SITIO / "estilos.css")
    shutil.copy2(PLANTILLAS / "sw.js", SITIO / "sw.js")
    (SITIO / ".nojekyll").touch()
    construir_indice(presentaciones, config)
    print(f"Listo: {len(presentaciones)} presentaciones en {SITIO}")


if __name__ == "__main__":
    main()
