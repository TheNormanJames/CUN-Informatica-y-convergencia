/**
 * @fileoverview Capa de datos del catalogo comercial, indice precomputado Map O(1)
 * y generadores sanitizados de contacto para Clasicos Morandua.
 */

export const BUSINESS_CONTACT = Object.freeze({
  brandName: 'Clasicos Morandua',
  phoneDisplay: '+57 300 000 0000',
  phoneRaw: '573000000000',
  email: 'contacto@clasicosmorandua.ejemplo.co',
  address: 'Distrito Maderero Tradicional (Zona Doce de Octubre / Barrios Unidos), Bogota D.C., Colombia',
  schedule: 'Lunes a Sabado: 8:30 AM - 6:00 PM | Domingos y Festivos: 10:00 AM - 3:00 PM',
  logistics: 'Despacho asegurado con flota de transporte aliado en Bogota, Sabana Centro y municipios de Cundinamarca.',
  socialMeta: {
    instagram: 'https://www.instagram.com/',
    facebook: 'https://www.facebook.com/'
  }
});

export const CATEGORIES = Object.freeze({
  ALL: 'todos',
  FLOR_MORADO: 'flor-morado',
  DEKO: 'deko',
  RESTAURACION: 'restauracion'
});

export const CATEGORY_LABELS = Object.freeze({
  [CATEGORIES.ALL]: 'Todo el Catalogo',
  [CATEGORIES.FLOR_MORADO]: 'Linea Flor Morado',
  [CATEGORIES.DEKO]: 'Linea Deko',
  [CATEGORIES.RESTAURACION]: 'Restauracion y Refaccion'
});

/**
 * Catalogo oficial de Clasicos Morandua (6 items representativos).
 */
export const CATALOG_ITEMS = Object.freeze([
  Object.freeze({
    id: 'fm-comedor-imperial-6p',
    title: 'Comedor Imperial Sabana (6 Puestos)',
    category: CATEGORIES.FLOR_MORADO,
    categoryLabel: 'Linea Flor Morado',
    woodType: 'Flor Morado Macizo Secado al Horno',
    dimensions: '180 cm largo x 100 cm ancho x 78 cm alto',
    priceReference: 'Desde $4.250.000 COP',
    warranty: 'Garantia estructural de 10 anos en ensambles de caja y espigo',
    description: 'Mesa de comedor familiar en madera maciza de Flor Morado con veta natural pronunciada, acabado en laca catalizada semimate resistente al calor y seis sillas ergonomicas tapizadas en lino antimanchas.',
    badge: 'Alta Ebanisteria',
    hasAR: true,
    arModelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
    arModelTitle: 'Silla Maestro Comedor Imperial en Flor Morado (Vista 3D / AR)',
    visualTheme: 'wood-warm',
    swatches: Object.freeze(['#C49A6C', '#8B5A3C', '#E8E3DA'])
  }),
  Object.freeze({
    id: 'fm-cama-santafe-queen',
    title: 'Cama Capitular Santafe (Tamano Queen)',
    category: CATEGORIES.FLOR_MORADO,
    categoryLabel: 'Linea Flor Morado',
    woodType: 'Flor Morado Seleccionado',
    dimensions: '168 cm ancho x 205 cm fondo x 125 cm cabecero',
    priceReference: 'Desde $2.890.000 COP',
    warranty: 'Garantia estructural de 10 anos con largueros reforzados',
    description: 'Estructura de cama clasica con cabecero entablerado a mano y tendido pesado antiruido. Ensamblada en nuestro taller satelite de ebanisteria tradicional sin tornilleria expuesta.',
    badge: 'Madera Maciza',
    hasAR: true,
    arModelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
    arModelTitle: 'Poltrona y Modulo de Alcoba en Flor Morado (Vista 3D / AR)',
    visualTheme: 'wood-deep',
    swatches: Object.freeze(['#8B5A3C', '#5C4033', '#DDD7CA'])
  }),
  Object.freeze({
    id: 'dk-biblioteca-estudio-modular',
    title: 'Biblioteca Modular Teusaquillo',
    category: CATEGORIES.DEKO,
    categoryLabel: 'Linea Deko',
    woodType: 'Madera Deko Estabilizada y Enchapado Natural',
    dimensions: '160 cm ancho x 38 cm fondo x 210 cm alto',
    priceReference: 'Desde $2.150.000 COP',
    warranty: 'Garantia de 5 anos en herrajes y estabilidad dimensional',
    description: 'Biblioteca de diseno contemporaneo calido con entrepanos graduables de alta carga para libros pesados, nicho central para estudio y puertas inferiores con cierre lento.',
    badge: 'Diseno Funcional',
    hasAR: true,
    arModelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
    arModelTitle: 'Silla de Lectura Linea Deko (Vista 3D / AR)',
    visualTheme: 'deko-sand',
    swatches: Object.freeze(['#D7BA89', '#C49A6C', '#5C4033'])
  }),
  Object.freeze({
    id: 'dk-mesa-centro-bacata',
    title: 'Mesa de Centro y Auxiliar Bacata',
    category: CATEGORIES.DEKO,
    categoryLabel: 'Linea Deko',
    woodType: 'Linea Deko con Acabado Poro Abierto',
    dimensions: '110 cm largo x 65 cm ancho x 42 cm alto',
    priceReference: 'Desde $980.000 COP',
    warranty: 'Garantia de 5 anos en estructura y acabado protector',
    description: 'Juego de mesa de centro con bandeja inferior para revistas y bordes redondeados seguros para el hogar, ideal para salas bogotanas con iluminacion calida.',
    badge: 'Entrega Rapida',
    hasAR: false,
    arModelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
    arModelTitle: 'Mobiliario Auxiliar Linea Deko (Vista 3D / AR)',
    visualTheme: 'deko-warm',
    swatches: Object.freeze(['#D7BA89', '#8B5A3C'])
  }),
  Object.freeze({
    id: 'rs-comedor-herencia-familiar',
    title: 'Restauracion Integral de Comedores de Herencia (Antes y Despues)',
    category: CATEGORIES.RESTAURACION,
    categoryLabel: 'Restauracion y Refaccion',
    woodType: 'Recuperacion de Cedro, Flor Morado, Caoba y Roble',
    dimensions: 'Servicio a medida con recoleccion en domicilio en Bogota',
    priceReference: 'Diagnostico tecnico sin costo por WhatsApp',
    warranty: 'Garantia de 3 anos en pintura catalizada y reensamble',
    description: 'Antes: madera reseca, rayones profundos, uniones desajustadas y espuma vencida. Despues: decapado manual en taller de pintura, ajuste de espigos, inmunizado y retapizado premium.',
    badge: 'Antes / Despues',
    beforeState: 'Barniz cuartado por sol, bisagras cedidas y holgura estructural.',
    afterState: 'Poros sellados en poliuretano mate, uniones prensadas y tapiceria nueva.',
    hasAR: false,
    arModelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
    arModelTitle: 'Muestra de Acabado Restaurado (Vista 3D / AR)',
    visualTheme: 'restore-craft',
    swatches: Object.freeze(['#5C4033', '#C49A6C', '#3D523D'])
  }),
  Object.freeze({
    id: 'rs-poltronas-bifamiliar-clasica',
    title: 'Refaccion de Salas Clasicas y Cambio de Tono (Antes y Despues)',
    category: CATEGORIES.RESTAURACION,
    categoryLabel: 'Restauracion y Refaccion',
    woodType: 'Talleres Satelites de Ebanisteria, Talla y Pintura',
    dimensions: 'Tiempo estimado de intervencion: 12 a 18 dias habiles',
    priceReference: 'Cotizacion personalizada segun estado de la pieza',
    warranty: 'Garantia de 3 anos en estructura, resortado y costura',
    description: 'Antes: salas oscurecidas de anos 80 y 90 con resortes cedidos. Despues: renovacion a tonos miel o caramelo natural que resaltan la veta original y espumas de alta densidad.',
    badge: 'Taller Satelite',
    beforeState: 'Laca oscura envejecida, golpes en patas talladas y tela desgastada.',
    afterState: 'Tono miel calido Flor Morado, talla reconstruida y lino antifluidos.',
    hasAR: true,
    arModelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
    arModelTitle: 'Poltrona Clasica Restaurada (Vista 3D / AR)',
    visualTheme: 'restore-warm',
    swatches: Object.freeze(['#C49A6C', '#8B5A3C', '#E8E3DA'])
  })
]);

/**
 * Construye un indice precomputado en un Map agrupando los productos por categoria.
 * Se ejecuta una sola vez durante la inicializacion del modulo.
 *
 * Complejidad Temporal: O(n), donde n es el total de elementos en `items` (recorrido lineal simple sin bucles anidados).
 * Complejidad Espacial: O(n), almacenando referencias a los n elementos en listas agrupadas por clave de categoria.
 *
 * @param {ReadonlyArray<Object>} items - Lista de productos y servicios del catalogo.
 * @returns {Map<string, ReadonlyArray<Object>>} Indice Map para consultas en O(1).
 */
export function buildCategoryIndex(items = CATALOG_ITEMS) {
  const index = new Map();
  const allItems = [];

  for (let i = 0; i < items.length; i += 1) {
    const item = items[i];
    allItems.push(item);

    const categoryKey = item.category;
    let bucket = index.get(categoryKey);
    if (!bucket) {
      bucket = [];
      index.set(categoryKey, bucket);
    }
    bucket.push(item);
  }

  index.set(CATEGORIES.ALL, Object.freeze(allItems));

  for (const [key, bucket] of index.entries()) {
    if (key !== CATEGORIES.ALL) {
      index.set(key, Object.freeze(bucket));
    }
  }

  return index;
}

/**
 * Indice global precomputado al cargar el modulo (O(n) una sola vez).
 */
export const CATEGORY_INDEX = buildCategoryIndex(CATALOG_ITEMS);

/**
 * Recupera los productos asociados a una categoria desde el Map precomputado.
 * Maneja categorias nulas o inexistentes devolviendo un arreglo vacio inmutable sin lanzar excepciones.
 *
 * Complejidad Temporal: O(1) promedio y peor caso amortizado gracias a la busqueda directa de clave en `Map.prototype.get`.
 * Complejidad Espacial: O(1) auxiliar, ya que retorna la referencia al arreglo precomputado sin clonar elementos.
 *
 * @param {string} category - Identificador de la categoria ('todos', 'flor-morado', 'deko', 'restauracion').
 * @param {Map<string, ReadonlyArray<Object>>} [index=CATEGORY_INDEX] - Indice precomputado.
 * @returns {ReadonlyArray<Object>} Arreglo de productos de la categoria o arreglo vacio si no existe.
 */
export function getProductsByCategory(category, index = CATEGORY_INDEX) {
  if (typeof category !== 'string' || !(index instanceof Map)) {
    return Object.freeze([]);
  }
  const normalized = category.trim().toLowerCase();
  const results = index.get(normalized);
  return results || Object.freeze([]);
}

/**
 * Sanitiza un texto de entrada eliminando caracteres de control y acotando su longitud en O(m) sobre un limite fijo.
 *
 * Complejidad Temporal: O(1) acotado a un maximo de 120 caracteres de entrada.
 * Complejidad Espacial: O(1) acotado a la cadena limpia resultante.
 *
 * @param {unknown} rawValue - Valor ingresado por el usuario.
 * @param {string} fallback - Texto por defecto si el valor esta vacio.
 * @returns {string} Cadena limpia y segura.
 */
export function sanitizeInputText(rawValue, fallback = 'Por definir en visita o asesoria') {
  if (typeof rawValue !== 'string') {
    return fallback;
  }
  const trimmed = rawValue.slice(0, 120).replace(/[\u0000-\u001F\u007F<>]/g, '').trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

/**
 * Genera en tiempo constante O(1) el enlace estructurado y sanitizado hacia WhatsApp Business
 * para el pre-triaje de cotizacion rapida (Mueble Nuevo o Restauracion).
 *
 * Complejidad Temporal: O(1), ya que realiza interpolacion de plantillas de longitud fija y codificacion URI acotada.
 * Complejidad Espacial: O(1), reservando memoria unicamente para la URL resultante y el resumen de pre-triaje.
 *
 * @param {Object} params - Parametros del asistente de cotizacion rapida.
 * @param {string} [params.serviceType='Mueble Nuevo'] - Tipo de solicitud ('Mueble Nuevo' o 'Restauracion').
 * @param {string} [params.woodType='Flor Morado'] - Linea de madera ('Flor Morado' o 'Deko').
 * @param {string} [params.dimensions=''] - Medidas aproximadas o detalle breve.
 * @param {string} [params.productTitle=''] - Nombre de un producto especifico si se invoca desde una tarjeta.
 * @returns {{ url: string, messagePreview: string, triageCode: string }} Objeto con URL lista para WhatsApp y vista previa.
 */
export function buildWhatsAppTriageUrl(params = {}) {
  const rawService = typeof params.serviceType === 'string' ? params.serviceType.trim() : '';
  const isRestoration = rawService.toLowerCase().includes('restaura');

  const normalizedService = isRestoration ? 'Restauracion / Refaccion' : 'Mueble Nuevo a Medida';
  const normalizedWood =
    typeof params.woodType === 'string' && params.woodType.trim().toLowerCase().includes('deko')
      ? 'Linea Deko'
      : 'Flor Morado Macizo';

  const cleanDimensions = sanitizeInputText(
    params.dimensions,
    'Medidas estandar (solicito orientacion de taller)'
  );
  const cleanProduct = sanitizeInputText(params.productTitle, 'Consulta general del catalogo web');

  const triageCode = isRestoration ? 'TRIAJE-IA-RESTAURACION' : 'TRIAJE-IA-NUEVO';

  const messagePreview =
    `Hola ${BUSINESS_CONTACT.brandName}, vengo del Asistente Web [${triageCode}]. ` +
    `Deseo cotizar: ${normalizedService}. ` +
    `Referencia: ${cleanProduct}. ` +
    `Madera o linea de interes: ${normalizedWood}. ` +
    `Medidas o estado aproximado: ${cleanDimensions}. ` +
    `Quedo atento a disponibilidad y tiempos de taller en Bogota.`;

  const encodedText = encodeURIComponent(messagePreview);
  const url = `https://wa.me/${BUSINESS_CONTACT.phoneRaw}?text=${encodedText}`;

  return {
    url,
    messagePreview,
    triageCode
  };
}

