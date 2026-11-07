# DESIGN SYSTEM: WARM ORGANIC MINIMALIST & EDITORIAL CRAFT
## Inspiración Visual: Referencias *ORLA* y *STILL / HOME*
### Aplicado a: Clásicos Morandúa (Sitio Web, Informe APA 7 en Word y Presentación Ejecutiva)

---

## 1. Filosofía de Diseño y Atmósfera Visual
El sistema de diseño traduce la estética editorial de marcas contemporáneas de mobiliario de autor (*ORLA* y *STILL / HOME*) al contexto de la tradición ebanista bogotana de **Clásicos Morandúa**:
* **Atmósfera:** Serena, luminosa, arquitectónica y arraigada en la nobleza de la madera maciza (Flor Morado) y la funcionalidad de la Línea Deko.
* **Tensión Armónica:** Convivencia entre la tipografía clásica editorial y la pulcritud geométrica moderna; amplios espacios negativos (respiro visual) frente a detalles táctiles de materiales (vetas, linos y ensambles).
* **Enfoque Humano y Accesible:** Diseñado para enamorar tanto a audiencias contemporáneas como a adultos mayores, asegurando un contraste visual superior (WCAG AAA), tamaños táctiles generosos (≥ 48px) y lectura sin esfuerzo cognitivo.

---

## 2. Paleta Cromática y Tokens de Color

| Token | Hex | Muestra / Función | Rol en la Interfaz |
| :--- | :--- | :--- | :--- |
| `--color-bg-linen` | `#FAF8F5` | Lino Calizo Cálido | Fondo base de la página, diapositivas e informe; genera calidez sin cansar la vista. |
| `--color-bg-sand` | `#F3EFEA` | Arena de Cantera Suave | Paneles divididos, bloques de historia, tarjetas destacadas y filas alternas. |
| `--color-surface` | `#FFFFFF` | Blanco Yeso Puro | Superficie de tarjetas de productos, modales y lienzos de presentación. |
| `--color-charcoal` | `#1C1917` | Carbón Profundo (95% negro) | Titulares principales, textos de alto contraste (16.4:1 contra el lino) y botones primarios. |
| `--color-text-muted` | `#78716C` | Piedra Taupe / Humo | Párrafos secundarios, descripciones técnicas, notas al pie y kickers superiores. |
| `--color-wood-amber` | `#8B5A3C` | Madera Flor Morado | Acento natural de artesanía, bordes de enfoque, insignias y botones secundarios cálidos. |
| `--color-wood-dark` | `#633B24` | Corteza Nogal Profunda | Estados hover de acentos de madera y cabeceras de tablas en documentos formales. |
| `--color-olive` | `#3D523D` | Verde Olivo / Laurel Silvestre | Botones de WhatsApp Business, estados de confirmación y llamadas a la naturaleza. |
| `--color-border-subtle`| `#E7E2D8` | Trazo Calizo / Línea Fina | Delimitadores de 1px limpios, bordes sutiles de tarjetas y divisiones de navegación. |

---

## 3. Jerarquía Tipográfica y Reglas de Puntuación

Inspirada en *ORLA* y *STILL / HOME*, la tipografía se basa en dos familias complementarias:
1. **Familia Display / Titulares:** *Editorial Serif* (`Cormorant Garamond`, `Playfair Display`, `Georgia`, serif tradicional).
   - **Regla del Punto Final:** Todos los titulares principales de impacto en el Hero y banners editoriales culminan con un punto ortográfico sobrio (`.`):
     - *Ejemplos:* `"Objetos para un habitar más sereno."` | `"Madera maciza, tiempos más lentos."` | `"Muebles concebidos para perdurar."`
   - Tamaño: `2.4rem` a `3.2rem` en escritorio, `1.85rem` a `2.1rem` en móvil.
   - Altura de línea (`line-height`): Compacta y elegante (`1.15` a `1.22`).
2. **Familia de Lectura e Interfaz:** *Modern Clean Sans* (`Inter`, `Plus Jakarta Sans`, `-apple-system`, `system-ui`).
   - Peso regular para párrafos (`400`), semibold para llamadas a la acción (`600`).
   - Tamaño base: `18px` (1.125rem) para garantizar accesibilidad en adultos mayores.
   - Altura de línea (`line-height`): Muy holgada (`1.65` a `1.8`) para dar aire al texto.
3. **Kickers / Eyebrows (Pretítulos):**
   - Mayúsculas sostenidas (`uppercase`), tamaño reducido (`11px` a `13px`), peso `600`, espaciado de letras expandido (`letter-spacing: 0.18em` a `0.22em`), color `--color-text-muted` o `--color-wood-amber`.
   - *Ejemplos:* `MATERIALES NOBLES, ESPACIOS EN CALMA.` | `COLECCIÓN DESTACADA` | `SECCIÓN 5.5 - WEBAR 3D`.

---

## 4. Componentes y Patrones de Interfaz

### 4.1. Barra de Navegación Editorial
- **Logotipo:** Tipografía sans o serif con espaciado amplio: `CLÁSICOS MORANDÚA` o `MORANDÚA / BOGOTÁ`.
- **Menú:** Enlaces en minúscula/título (`Inicio`, `Quiénes Somos`, `Colección`, `Ubicación`, `Contacto`) con indicador sutil de pestaña activa (línea inferior de 2px en carbón).
- **Acceso Rápido:** Botón píldora o enlace sutil hacia WhatsApp Business o cotización asistida.

### 4.2. Tarjetas de Catálogo (Featured Objects Layout)
- **Marco Visual:** Relación de aspecto limpia (4:5 o 1:1), fondo neutro estudio en lino claro con iluminación cenital suave y sombras difusas (sin cajas duras ni brillos plásticos).
- **Muestras de Color / Material (Swatches):** Pequeños círculos cromáticos bajo el título que ilustran los acabados disponibles:
  - Círculo 1: Tono Miel Flor Morado (`#C49A6C`).
  - Círculo 2: Roble Natural Deko (`#D7BA89`).
  - Círculo 3: Nogal Tabaco (`#5C4033`).
  - Círculo 4: Lino Crudo Textil (`#E8E3DA`).
- **Disposición de Metadatos:**
  - Fila 1: Título de la pieza a la izquierda (`Comedor Imperial Sabana`) y flecha sutil a la derecha (`->`).
  - Fila 2: Categoría o madera en color taupe (`Flor Morado Macizo`) y precio o cotización base (`Desde $4.250.000 COP`).
- **Acciones:** Botón píldora negro/carbón para cotizar y botón contorneado de trazo fino para abrir el visor en Realidad Aumentada 3D.

### 4.3. Sección Dividida de Materiales y Oficio (Split Craft Showcase)
Inspirada en el bloque *LASTING MATERIALS* de *STILL / HOME*:
- **Columna Izquierda (Macro Textura Madera):**
  - Kicker: `MATERIALES QUE TRASCIENDEN`
  - Titular: `Creado por la naturaleza, tallado a mano.`
  - Texto: `Maderas macizas secadas al horno. Ensambles de caja y espigo sin tornillería visible. Mobiliario honesto que madura con gracia con el paso de las décadas.`
  - Enlace: `Conocer nuestro proceso artesanal ->`
- **Columna Derecha (Linos y Lacado de Taller):**
  - Kicker: `TALLERES SATÉLITES EN BOGOTÁ`
  - Titular: `Ebanistería y acabados que respiran con tu hogar.`
  - Texto: `Sellado de poro abierto y lacas catalizadas resistentes a la humedad de la Sabana.`

### 4.4. Botones y Microinteracciones
- **Botón Primario:** Fondo sólido carbón (`#1C1917`), texto blanco marfil, bordes redondeados (`4px` a `8px` o píldora suave), flecha direccional integrada (`->`), transición suave (`0.2s`).
- **Botón WhatsApp:** Verde olivo profundo (`#1F5E3B` o `#3D523D`), transmitiendo serenidad en vez de estridencia comercial.
- **Botón Secundario / AR:** Fondo transparente, borde fino de `1.5px` en carbón o madera, texto oscuro.
- **Área Táctil:** Altura mínima de `48px` para garantizar clic fácil en dispositivos táctiles de personas mayores.

---

## 5. Aplicación a los Documentos Académicos (Word y PowerPoint)

### 5.1. Informe APA 7 en Word (`ProyectoFinal_Norman_Jaimes.docx`)
* **Portada:** Maquetación sobria y espaciosa con márgenes simétricos de `2.54 cm` (1 pulgada) y numeración dinámica en el extremo superior derecho.
* **Tabla de Contenido (Norma APA 7 Estricta):**
  - En APA 7ª edición, la Tabla de Contenido **no es una tabla de datos**; por lo tanto, **no lleva bordes, no lleva sombreados oscuros de cabecera ni se numera como `Tabla 1`**.
  - Se estructura como un índice jerárquico limpio con sangría en subniveles, líneas de puntos guía (*dot leaders* `.......`) y números de página alineados a la derecha.
* **Tablas de Datos (`Tabla 1` a `Tabla 6` bajo Norma APA 7 Estricta):**
  - Etiqueta en negrita (`Tabla 1`) en una línea independiente y título descriptivo en *cursiva* en la línea inmediatamente inferior.
  - **Regla de las 3 líneas horizontales:** Únicamente borde superior de tabla, borde inferior del encabezado y borde inferior de tabla. Cero bordes verticales y cero cuadrículas internas pesadas.
  - Cierre con párrafo `Nota.` en cursiva al pie de cada tabla.
* **Figuras con Capturas Reales Incrustadas (`Figura 1` a `Figura 3`):** Numeradas y tituladas en formato APA 7 con las capturas reales del sitio web incrustadas directamente y su respectiva nota descriptiva.

### 5.2. Presentación Ejecutiva en PowerPoint (`ProyectoFinal_Norman_Jaimes.pptx`)
* **Lienzo Panorámico (16:9):** Fondo uniforme en lino cálido (`#FAF8F5`).
* **Regla del Hilo Conductor:** Titulares H1 en `Georgia` (22pt) en color carbón (`#1C1917`) con punto final reflexivo.
* **Estructura de Tarjetas Flotantes:**
  - Tarjeta izquierda blanca (`#FFFFFF`) con borde de `1px` color piedra, conteniendo exactamente tres viñetas limpias con viñeta tipográfica guion (`-`) y acento de madera.
  - Tarjeta lateral derecha blanca con borde cálido (`#8B5A3C`), etiqueta superior contrastada y desglose de métricas o arquitectura.
* **Notas de Orador:** Redacción en tono consultor sobrio en cada una de las 10 diapositivas.

---

## 6. Checklist de Control de Calidad Visual
- [x] Ausencia total de colores saturados estridentes o sombras duras de baja calidad.
- [x] Cumplimiento de contraste WCAG AAA en todos los pares de texto/fondo.
- [x] Titulares con puntuación final deliberada (`.`).
- [x] Espaciado de letras generoso en todos los textos auxiliares (`letter-spacing >= 0.15em`).
- [x] Flechas integradas (`->`) en botones y enlaces de conversión.
- [x] Coherencia 100% simétrica entre la web, el documento Word y las diapositivas de PowerPoint.

