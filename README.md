# Clasicos Morandua - Sitio Web Estatico, WebAR 3D y Pre-Triaje IA

Prototipo funcional del sitio web de **Clasicos Morandua** (almacen tradicional familiar de muebles en madera maciza Flor Morado y Linea Deko, ebanisteria artesanal y restauracion en Bogota D.C.), desarrollado como evidencia practica de las **Secciones 5.5 (Realidad Aumentada e IA)** y **5.9 (Sitio Web y UX/UI)** del Proyecto Final de Ingenieria de Sistemas (ACA).

## Arquitectura y Cumplimiento Tecnico

1. **5 Secciones Navegables Exigidas:**
   - **Inicio (`#inicio`):** Hero editorial con propuesta de valor en madera maciza y accesos directos a catalogo, visor WebAR 3D y WhatsApp Business.
   - **Quienes Somos (`#quienes-somos`):** Historia artesanal e integracion del almacen principal con los talleres satelites de ebanisteria y pintura.
   - **Productos y Servicios (`#productos-servicios`):** Catalogo interactivo filtrable entre Linea Flor Morado, Linea Deko y Restauracion/Refaccion (Antes y Despues), mas visor 3D / Realidad Aumentada con `<model-viewer>`.
   - **Ubicacion (`#ubicacion`):** Sede fisica en Bogota, horarios de atencion y logistica de transporte aliado.
   - **Contacto (`#contacto`):** Asistente de Cotizacion Rapida (Demo IA) con generacion de enlace a WhatsApp Business (`+57 300 000 0000`), correo comercial (`contacto@clasicosmorandua.ejemplo.co`) y redes Meta.
2. **Estetica Editorial Warm Minimalist y Accesibilidad AAA:**
   - Paleta hexadecimal estricta: Fondo primario `#F7F5F0`, Superficie `#FFFFFF`, Texto primario `#1C1917` (contraste AAA `16.4:1`), Acento Flor Morado `#8B5A3C`, Verde WhatsApp `#1F5E3B`.
   - Tipografia base `18px`, botones e interactivos con area tactil minima de `48px x 48px` y estados `:focus-visible` nitidos.
3. **Eficiencia Algoritmica (Criterio Big O):**
   - Indice precomputado `CATEGORY_INDEX` usando `Map` construido una sola vez en `O(n)` al inicializar `js/data/catalog.js`.
   - Consultas de filtrado por categoria en tiempo constante `O(1)` (`Map.prototype.get`) y renderizado DOM en `O(k)` mediante un unico `DocumentFragment`.
   - Generacion de enlaces de pre-triaje hacia WhatsApp Business en `O(1)`.

## Estructura del Repositorio

```text
ACA/
|-- index.html
|-- css/
|   `-- styles.css
|-- js/
|   |-- main.js
|   |-- data/
|   |   `-- catalog.js
|   `-- modules/
|       `-- filter.js
|-- tests/
|   `-- catalog.test.mjs
|-- assets/
|   `-- screenshots/
|       |-- desktop-1280x800.png
|       |-- desktop-catalogo-1280x800.png
|       |-- mobile-390x844.png
|       `-- mobile-contacto-ia-390x844.png
`-- .gitignore
```

## Ejecucion de Pruebas Automatizadas

El proyecto utiliza el corredor nativo de Node.js (`node:test` y `node:assert/strict`) sin dependencias externas de `npm`:

```bash
node --test tests/catalog.test.mjs
```
