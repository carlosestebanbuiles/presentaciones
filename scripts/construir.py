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


def indices_visibles(pptx: Path) -> dict[int, int | None]:
    """Número de diapositiva en el .pptx (base 1) -> número en el sitio, o None si está oculta.
    LibreOffice no exporta las diapositivas ocultas, así que el sitio solo cuenta las visibles."""
    try:
        from pptx import Presentation
        prs = Presentation(str(pptx))
    except Exception:
        return {}
    mapa, n = {}, 0
    for i, s in enumerate(prs.slides, 1):
        if s._element.get("show") == "0":
            mapa[i] = None
        else:
            n += 1
            mapa[i] = n
    return mapa


def leer_enlaces(pptx: Path) -> dict[int, list[dict]]:
    """Hipervínculos del .pptx por diapositiva (numeración del sitio).
    Cada enlace: {url | ir, texto, caja: [x, y, ancho, alto] en fracción de la diapositiva, o None}.
    - Forma o imagen con acción de clic: caja = la forma.
    - Texto: si el texto enlazado ocupa casi todo el cuadro, caja = el cuadro; si es parte de un párrafo
      más largo (una referencia, por ejemplo), solo se lista, para no volver clicable todo el cuadro.
    - Saltos a otra diapositiva: {ir: número}."""
    try:
        from pptx import Presentation
        from pptx.enum.action import PP_ACTION
        from pptx.oxml.ns import qn
        prs = Presentation(str(pptx))
    except Exception:
        return {}
    W, H = prs.slide_width or 1, prs.slide_height or 1
    visibles = indices_visibles(pptx)
    id_a_numero = {s.slide_id: i for i, s in enumerate(prs.slides, 1)}
    valido = re.compile(r"^(https?://|mailto:)", re.I)

    def caja_de(x, y, w, h):
        if None in (x, y, w, h) or w <= 0 or h <= 0:
            return None
        return [round(x / W, 4), round(y / H, 4), round(w / W, 4), round(h / H, 4)]

    def recorrer(formas, transformar, salida):
        for sh in formas:
            x, y, w, h = sh.left, sh.top, sh.width, sh.height
            if None not in (x, y, w, h):
                x, y, w, h = transformar(x, y, w, h)
            if sh.shape_type == 6:  # grupo: las coordenadas de los hijos están en el espacio del grupo
                xfrm = sh._element.grpSpPr.find(qn("a:xfrm"))
                choff = xfrm.find(qn("a:chOff")) if xfrm is not None else None
                chext = xfrm.find(qn("a:chExt")) if xfrm is not None else None
                if choff is not None and chext is not None and int(chext.get("cx")) and int(chext.get("cy")):
                    cx0, cy0 = int(choff.get("x")), int(choff.get("y"))
                    sx, sy = w / int(chext.get("cx")), h / int(chext.get("cy"))
                    def t_hijo(a, b, c, d, x=x, y=y, cx0=cx0, cy0=cy0, sx=sx, sy=sy):
                        return x + (a - cx0) * sx, y + (b - cy0) * sy, c * sx, d * sy
                else:
                    t_hijo = transformar
                recorrer(sh.shapes, t_hijo, salida)
                continue
            texto_forma = sh.text_frame.text.strip() if getattr(sh, "has_text_frame", False) and sh.has_text_frame else ""
            # Acción de clic sobre la forma completa
            url, destino = None, None
            try:
                accion = sh.click_action
                if accion.action == PP_ACTION.NAMED_SLIDE:
                    destino = accion.target_slide
                else:
                    url = accion.hyperlink.address
            except Exception:
                pass
            if url and valido.match(url):
                salida.append({"url": url, "texto": texto_forma[:120] or url, "caja": caja_de(x, y, w, h)})
            elif destino is not None:
                n = visibles.get(id_a_numero.get(destino.slide_id))
                if n:
                    salida.append({"ir": n, "texto": texto_forma[:120] or f"Diapositiva {n}", "caja": caja_de(x, y, w, h)})
            # Enlaces en el texto
            if texto_forma:
                por_url = {}
                for par in sh.text_frame.paragraphs:
                    for r in par.runs:
                        try:
                            u = r.hyperlink.address
                        except Exception:
                            u = None
                        if u and valido.match(u):
                            por_url.setdefault(u, []).append(r.text)
                total = len(re.sub(r"\s", "", texto_forma)) or 1
                for u, trozos in por_url.items():
                    enlazado = "".join(trozos).strip()
                    ocupa = len(re.sub(r"\s", "", enlazado)) / total
                    caja = caja_de(x, y, w, h) if (len(por_url) == 1 and ocupa >= 0.6) else None
                    salida.append({"url": u, "texto": enlazado[:120] or u, "caja": caja})

    resultado = {}
    for i, s in enumerate(prs.slides, 1):
        n = visibles.get(i)
        if not n:
            continue
        lista = []
        recorrer(s.shapes, lambda a, b, c, d: (a, b, c, d), lista)
        # quitar duplicados exactos
        vistos, unicos = set(), []
        for e in lista:
            clave = (e.get("url"), e.get("ir"), tuple(e["caja"]) if e.get("caja") else None)
            if clave not in vistos:
                vistos.add(clave)
                unicos.append(e)
        if unicos:
            resultado[n] = unicos
    return resultado


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
    ids_ordenados = re.findall(r'<p:sldId [^>]*id="(\d+)"', lista.group(1))
    visibles = indices_visibles(pptx)
    orden = {sid: visibles[i + 1] for i, sid in enumerate(ids_ordenados) if visibles.get(i + 1)}
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
    # Enlaces del PowerPoint, por diapositiva (para el visor y para la ficha)
    info["enlaces_pptx"] = leer_enlaces(fuente) if fuente.suffix.lower() == ".pptx" else {}
    items, vistos = [], set()
    for e in info["enlaces"]:  # los de presentaciones.yml, primero
        if e.get("url") and e["url"] not in vistos:
            vistos.add(e["url"])
            items.append(f'<li><a href="{html.escape(e["url"])}" target="_blank" rel="noopener">'
                         f'{html.escape(e.get("texto", e["url"]))}</a></li>')
    for n in sorted(info["enlaces_pptx"]):
        for e in info["enlaces_pptx"][n]:
            if not e.get("url") or e["url"] in vistos:
                continue
            vistos.add(e["url"])
            texto = e["texto"] if e["texto"] != e["url"] else re.sub(r"^https?://(www\.)?", "", e["url"])
            items.append(f'<li><a href="{html.escape(e["url"])}" target="_blank" rel="noopener">{html.escape(texto)}</a>'
                         f' <a class="en-diapo" href="visor.html#{n}">diapositiva {n}</a></li>')
    enlaces_html = "".join(items)
    detalles = " · ".join(x for x in [fecha_legible(info["fecha"]), info["autor"], f'{info["n"]} diapositivas'] if x)
    (carpeta / "index.html").write_text(
        renderizar(
            "presentacion.html",
            titulo=html.escape(titulo),
            detalles=html.escape(detalles),
            descripcion=html.escape(info["descripcion"]),
            n=info["n"],
            v=info["v"],
            pdf=f"{slug}.pdf",
            pptx=descargable,
            boton_pptx=(f'<a class="boton" href="{descargable}" download>Descargar PowerPoint</a>' if descargable else ""),
            enlaces=(f'<h2>Enlaces</h2><ul class="lista-enlaces">{enlaces_html}</ul>' if enlaces_html else ""),
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


# ---- Portadas de las carpetas ----
# Se configuran en presentaciones.yml (_sitio.portadas), por ruta de carpeta:
#   "Académicas": {color: "#2e7d5b", icono: academico}
#   "Festina Lente": {imagen: festina-lente.jpg}      (archivo dentro de portadas/)
# Una subcarpeta sin configuración hereda el color y el ícono de su carpeta madre, con un matiz propio.
PORTADAS: dict = {}
ICONOS = {
    "academico": '<path d="M2 9.5 12 4.5l10 5-10 5z"/><path d="M6 11.5v4.8c0 1.4 2.7 3.2 6 3.2s6-1.8 6-3.2v-4.8"/><path d="M22 9.5v5.5"/>',
    "industria": '<path d="M9.5 3h5"/><path d="M10.5 3v6.2L5 18.6A1.9 1.9 0 0 0 6.6 21.5h10.8a1.9 1.9 0 0 0 1.6-2.9L13.5 9.2V3"/><path d="M7.4 15h9.2"/>',
    "carpeta": '<path d="M3 7.5A2 2 0 0 1 5 5.5h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
}
COLOR_POR_DEFECTO = "#4f6478"


def _hex_a_rgb(c: str):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _rgb_a_hex(rgb) -> str:
    return "#" + "".join(f"{max(0, min(255, round(x * 255))):02x}" for x in rgb)


def _ajustar(color: str, matiz: float = 0.0, luz: float = 0.0, sat: float = 0.0) -> str:
    import colorsys
    h, l, s = colorsys.rgb_to_hls(*_hex_a_rgb(color))
    return _rgb_a_hex(colorsys.hls_to_rgb((h + matiz) % 1, max(0, min(1, l + luz)), max(0, min(1, s + sat))))


def estilo_carpeta(ruta: list[str], posicion: int) -> dict:
    """Color, ícono e imagen para la carpeta cuya ruta completa es `ruta` (p. ej. ["Industria farmacéutica", "Amgen"])."""
    propia = PORTADAS.get(" / ".join(ruta)) or {}
    if len(ruta) > 1:
        madre = estilo_carpeta(ruta[:-1], 0)
        # matiz propio para cada subcarpeta: alterna hacia un lado y otro del color de la madre
        paso = ((posicion // 2) + 1) * (1 if posicion % 2 == 0 else -1)
        color = propia.get("color") or _ajustar(madre["color"], matiz=0.035 * paso, luz=0.03 * (1 if paso > 0 else -1))
        icono = propia.get("icono") or madre["icono"]
    else:
        color = propia.get("color") or COLOR_POR_DEFECTO
        icono = propia.get("icono") or "carpeta"
    return {"color": color, "icono": icono if icono in ICONOS else "carpeta", "imagen": propia.get("imagen")}


def portada_carpeta_html(ruta: list[str], posicion: int, prefijo: str) -> str:
    e = estilo_carpeta(ruta, posicion)
    if e["imagen"] and (RAIZ / "portadas" / e["imagen"]).exists():
        v = hash_archivo(RAIZ / "portadas" / e["imagen"])[:8]
        return (f'<div class="portada"><img src="{prefijo}portadas/{html.escape(e["imagen"])}?v={v}" '
                f'alt="{html.escape(ruta[-1])}" loading="lazy"></div>')
    c1 = _ajustar(e["color"], luz=0.06)
    c2 = _ajustar(e["color"], luz=-0.10, matiz=0.02)
    madre = f'<span class="ruta-portada">{html.escape(" · ".join(ruta[:-1]))}</span>' if len(ruta) > 1 else ""
    svg = f'<svg viewBox="0 0 24 24" aria-hidden="true">{ICONOS[e["icono"]]}</svg>'
    return (f'<div class="portada portada-tipo" style="--c1:{c1};--c2:{c2}">'
            f'<span class="fondo-icono">{svg}</span>'
            f'<span class="icono-portada">{svg}</span>{madre}'
            f'<span class="nombre-portada">{html.escape(ruta[-1])}</span></div>')


def tarjeta_carpeta_html(nombre: str, carpeta: dict, prefijo: str, ruta_madre: list[str] | None = None, posicion: int = 0) -> str:
    n = contar(carpeta)
    nsub = len(carpeta["sub"])
    detalle = f'{n} presentaci{"ón" if n == 1 else "ones"}' + (f' · {nsub} subcategor{"ía" if nsub == 1 else "ías"}' if nsub else "")
    portada = portada_carpeta_html((ruta_madre or []) + [nombre], posicion, prefijo)
    return f"""
      <a class="carpeta" href="{slugificar(nombre)}/">
        {portada}
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

    carpetas_html = "".join(tarjeta_carpeta_html(nombre, carpeta["sub"][nombre], prefijo, ruta_nombres, k)
                            for k, nombre in enumerate(subnombres))
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
        tarjeta_carpeta_html(n, raiz["sub"][n], "", [], k).replace(f'href="{slugificar(n)}/"', f'href="c/{slugificar(n)}/"')
        for k, n in enumerate(sorted(raiz["sub"], key=lambda n: (orden.index(n) if n in orden else len(orden), n == sin_clasificar, n.lower())))
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
    PORTADAS.clear()
    PORTADAS.update({str(k).strip(): (v or {}) for k, v in (config.get("portadas") or {}).items()})

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
        if carpeta.is_dir() and carpeta.name not in vivos and carpeta.name not in ("c", "portadas"):
            shutil.rmtree(carpeta)
            print(f"  - {carpeta.name}: fuente eliminada, carpeta borrada")

    # Visor compartido (una sola copia por presentación, apunta a sus imágenes)
    for p in presentaciones:
        shutil.copy2(PLANTILLAS / "visor.html", SITIO / p["slug"] / "visor.html")
        (SITIO / p["slug"] / "visor.json").write_text(
            json.dumps({"titulo": p["titulo"], "n": p["n"], "v": p["v"], "secciones": p.get("secciones", []),
                        "enlaces": p.get("enlaces_pptx", {})}, ensure_ascii=False)
        )
    shutil.copy2(PLANTILLAS / "estilos.css", SITIO / "estilos.css")
    shutil.copy2(PLANTILLAS / "sw.js", SITIO / "sw.js")
    if (RAIZ / "portadas").exists():
        shutil.copytree(RAIZ / "portadas", SITIO / "portadas", dirs_exist_ok=True)
    (SITIO / ".nojekyll").touch()
    construir_indice(presentaciones, config)
    print(f"Listo: {len(presentaciones)} presentaciones en {SITIO}")


if __name__ == "__main__":
    main()
