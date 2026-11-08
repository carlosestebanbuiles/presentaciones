# Presentaciones

Sitio para presentar clases y charlas directamente desde el navegador, sin PowerPoint ni visores externos.
Cada presentación se publica como imágenes de alta resolución (1920 px de ancho) con un visor web propio,
más el PDF y el archivo original para descargar.

## Cómo agregar una presentación

1. Copiar el archivo `.pptx` (o `.pdf`) en la carpeta `fuentes/`.
   El nombre del archivo se usa como título si no se indica otro; conviene algo como `Hipotiroidismo_2026.pptx`.
2. (Opcional) Agregar título, fecha, autor, categoría, descripción y enlaces en `presentaciones.yml`.
   La portada muestra las categorías como carpetas; al entrar se ven sus subcategorías y presentaciones.
   La categoría admite niveles con `/`, por ejemplo
   `categoria: Académicas` o `categoria: Industria farmacéutica / Recordati`. El orden de las
   categorías se fija en `_sitio.orden_categorias`; las presentaciones sin categoría van en "Sin clasificar".
   Cambiar la categoría de una presentación no cambia su dirección, así que los enlaces siguen valiendo.
3. Hacer *commit* y *push* a la rama `main`.

GitHub Actions convierte la presentación y publica el sitio en unos minutos. Las presentaciones que no
cambiaron no se vuelven a convertir.

Para quitar una presentación basta con borrar el archivo de `fuentes/`.

## Cómo se presenta

En el índice, el botón **Presentar** abre el visor:

- `→`, espacio o clic en la mitad derecha: siguiente; `←` o clic en la mitad izquierda: anterior.
- `F` (o el botón de iniciar del presentador de mano): pantalla completa; `B` o `.`: pantalla en negro; `Esc`: salir. En celular o tableta se desliza con el dedo.
- La dirección lleva el número de diapositiva (`visor.html#12`), así que se puede compartir un enlace a una diapositiva concreta.

## Estructura

```
fuentes/            los .pptx originales (lo único que hay que tocar cada vez)
presentaciones.yml  metadatos opcionales y configuración del índice
plantillas/         visor, índice, ficha y estilos
scripts/construir.py  conversor: pptx → pdf (LibreOffice) → jpg (pdftoppm) + páginas
sitio/              salida generada; no se sube al repositorio
.github/workflows/publicar.yml  automatización
```

## Probar en el computador

Requiere LibreOffice, poppler (`pdftoppm`) y Python 3 con `pyyaml`.

```
python3 scripts/construir.py
cd sitio && python3 -m http.server 8000
```

y abrir http://localhost:8000.

## Notas

- Las imágenes son JPG con calidad 90, visualmente iguales al original. Para cambiar la resolución o la calidad,
  editar `ANCHO` y `CALIDAD_JPG` al inicio de `scripts/construir.py`.
- La conversión la hace LibreOffice; si una diapositiva usa una fuente poco común que LibreOffice no tiene,
  el texto puede verse ligeramente distinto. Las fuentes de Office (Calibri, Cambria, Arial, Times) tienen
  equivalentes métricos instalados en la automatización.
- Las animaciones dentro de una diapositiva no se conservan (cada diapositiva es una imagen fija);
  las transiciones entre diapositivas las hace el visor con un fundido.
