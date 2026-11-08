# Clásicos Morandúa — Sitio Web Estático, WebAR 3D y Asistente de Cotización

**Corporación Unificada Nacional de Educación Superior (CUN)**  
**Asignatura:** Informática y Convergencia Tecnológica (`TI010 / 55522 / PRIMER BLOQUE / 26V05`)  
**Estudiante:** Norman Eduardo Jaimes Mora  
**Año:** 2026  

- **Sitio Web en Vivo (GitHub Pages):** [https://thenormanjames.github.io/CUN-Informatica-y-convergencia/](https://thenormanjames.github.io/CUN-Informatica-y-convergencia/)
- **Informe Escrito (PDF - Normas APA 7):** [`docs/ProyectoFinal_Norman_Jaimes.pdf`](./docs/ProyectoFinal_Norman_Jaimes.pdf)
- **Informe Escrito (Word Editable):** [`docs/ProyectoFinal_Norman_Jaimes.docx`](./docs/ProyectoFinal_Norman_Jaimes.docx)
- **Presentación Ejecutiva (PowerPoint 10 Diapositivas):** [`docs/ProyectoFinal_Norman_Jaimes.pptx`](./docs/ProyectoFinal_Norman_Jaimes.pptx)
- **Presentación Ejecutiva (PDF):** [`docs/ProyectoFinal_Norman_Jaimes_Presentacion.pdf`](./docs/ProyectoFinal_Norman_Jaimes_Presentacion.pdf)

---

## Descripción General del Proyecto

Prototipo funcional del sitio web de **Clásicos Morandúa** (almacén tradicional familiar de muebles en madera maciza Flor Morado y Línea Deko, ebanistería artesanal y restauración en Bogotá D.C.), desarrollado como evidencia práctica de las **Secciones 5.5 (Tecnologías Emergentes: IA y Realidad Aumentada)** y **5.9 (Sitio Web y Experiencia de Usuario UX/UI)** del Proyecto Final de Ingeniería de Sistemas.

## Arquitectura y Cumplimiento Académico

1. **5 Secciones Navegables en Página Única (One-Page):**
   - **Inicio (`#inicio`):** Presentación editorial con propuesta de valor en madera maciza, indicadores de garantía y accesos directos a catálogo, visor WebAR 3D y WhatsApp Business.
   - **Nosotros (`#quienes-somos`):** Historia artesanal e integración del almacén principal con los talleres satélites de ebanistería y pintura catalizada.
   - **Catálogo (`#productos-servicios`):** Catálogo interactivo filtrable al instante entre Línea Flor Morado, Línea Deko y Restauración/Refacción (Antes y Después), más visor 3D / Realidad Aumentada con `<model-viewer>`.
   - **Taller y Sede (`#ubicacion`):** Sede física en Bogotá D.C., horarios de atención y ruta logística de transporte aliado.
   - **Cotizar (`#contacto`):** Cotizador Rápido por WhatsApp Business (`+57 300 000 0000`), correo comercial (`contacto@clasicosmorandua.ejemplo.co`) y redes del ecosistema Meta.
2. **Estética Editorial Warm Minimalist y Accesibilidad para Adultos Mayores (WCAG AAA):**
   - Paleta hexadecimal de alto contraste: Fondo primario `#F7F5F0` (lino cálido), Superficie `#FFFFFF`, Texto primario `#1C1917` (contraste AAA `16.4:1`), Acento Flor Morado `#8B5A3C`, Verde WhatsApp `#1F5E3B`.
   - Tipografía sans-serif moderna (`Outfit` y `Plus Jakarta Sans`) con tamaño base de `18px`, botones e interactivos con área táctil mínima de `48px x 48px` y estados `:focus-visible` nítidos.

## Estructura del Repositorio

```text
.
├── index.html
├── css/
│   └── styles.css
├── js/
│   ├── main.js
│   ├── data/
│   │   └── catalog.js
│   └── modules/
│       └── filter.js
├── assets/
│   ├── images/
│   ├── models/
│   └── screenshots/
├── docs/
│   ├── ProyectoFinal_Norman_Jaimes.pdf
│   ├── ProyectoFinal_Norman_Jaimes.docx
│   ├── ProyectoFinal_Norman_Jaimes.pptx
│   ├── ProyectoFinal_Norman_Jaimes_Presentacion.pdf
│   └── generate_deliverables.py
└── tests/
    └── catalog.test.mjs
```

## Ejecución de Pruebas Automatizadas

El proyecto utiliza el corredor nativo de pruebas de Node.js (`node:test` y `node:assert/strict`) sin dependencias externas:

```bash
node --test tests/catalog.test.mjs
```
