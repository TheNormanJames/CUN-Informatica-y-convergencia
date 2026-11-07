/**
 * @fileoverview Modulo de renderizado y filtrado eficiente del catalogo en el DOM.
 * Utiliza consultas O(1) sobre CATEGORY_INDEX y una unica insercion con DocumentFragment en O(k).
 */

import {
  CATEGORIES,
  CATEGORY_INDEX,
  CATEGORY_LABELS,
  getProductsByCategory,
  buildWhatsAppTriageUrl
} from '../data/catalog.js';

/**
 * Crea el contenedor fotografico de cada pieza del catalogo con carga diferida (lazy loading)
 * y microinteraccion de zoom suave al pasar el cursor.
 *
 * Complejidad Temporal: O(1), construccion determinista de un nodo de imagen estatico.
 * Complejidad Espacial: O(1), tamano constante del subarbol DOM.
 *
 * @param {Object} item - Objeto de producto o servicio del catalogo.
 * @returns {HTMLElement} Contenedor visual de la tarjeta.
 */
export function createProductVisualBanner(item) {
  const visualTheme = item && item.visualTheme ? item.visualTheme : 'wood-warm';
  const banner = document.createElement('div');
  banner.className = `product-card__visual product-card__visual--${visualTheme}`;

  const img = document.createElement('img');
  img.className = 'product-card__img';
  img.src = (item && item.imageUrl) || './assets/images/comedor-flormorado.jpg';
  img.alt = (item && item.title) || 'Pieza de mobiliario artesanal Clasicos Morandua';
  img.loading = 'lazy';
  img.decoding = 'async';

  banner.appendChild(img);
  return banner;
}

/**
 * Construye el nodo DOM `<article>` para un item del catalogo sin usar innerHTML en bucles.
 * Diseno inspirado en galerias editoriales de mobiliario (ORLA y STILL / HOME).
 *
 * Complejidad Temporal: O(1), cantidad fija de nodos DOM creados por tarjeta.
 * Complejidad Espacial: O(1), estructura DOM acotada por tarjeta.
 *
 * @param {Object} item - Objeto de producto o servicio del catalogo.
 * @returns {HTMLElement} Elemento `<article>` listo para agregarse al DocumentFragment.
 */
export function createProductCardNode(item) {
  const article = document.createElement('article');
  article.className = 'product-card reveal-on-scroll is-visible';
  article.dataset.productId = item.id;
  article.dataset.category = item.category;

  const visualBanner = createProductVisualBanner(item);

  const badgeSpan = document.createElement('span');
  badgeSpan.className = 'product-card__badge';
  badgeSpan.textContent = item.badge;
  visualBanner.appendChild(badgeSpan);

  const body = document.createElement('div');
  body.className = 'product-card__body';

  const categoryTag = document.createElement('p');
  categoryTag.className = 'product-card__category';
  categoryTag.textContent = item.categoryLabel;

  const headerRow = document.createElement('div');
  headerRow.className = 'product-card__header-row';

  const title = document.createElement('h3');
  title.className = 'product-card__title';
  title.textContent = item.title;

  const arrowIcon = document.createElement('span');
  arrowIcon.className = 'product-card__arrow';
  arrowIcon.setAttribute('aria-hidden', 'true');
  arrowIcon.textContent = '\u2192';

  headerRow.append(title, arrowIcon);

  const metaRow = document.createElement('div');
  metaRow.className = 'product-card__meta-row';

  const priceTag = document.createElement('p');
  priceTag.className = 'product-card__price';
  priceTag.textContent = item.priceReference;
  metaRow.appendChild(priceTag);

  if (Array.isArray(item.swatches) && item.swatches.length > 0) {
    const swatchesContainer = document.createElement('div');
    swatchesContainer.className = 'product-card__swatches';
    swatchesContainer.setAttribute('aria-label', 'Muestras de acabado y material');

    for (let s = 0; s < item.swatches.length; s += 1) {
      const dot = document.createElement('span');
      dot.className = 'product-card__swatch-dot';
      dot.style.backgroundColor = item.swatches[s];
      swatchesContainer.appendChild(dot);
    }
    metaRow.appendChild(swatchesContainer);
  }

  const description = document.createElement('p');
  description.className = 'product-card__description';
  description.textContent = item.description;

  body.append(categoryTag, headerRow, metaRow, description);

  if (item.beforeState && item.afterState) {
    const comparisonBox = document.createElement('div');
    comparisonBox.className = 'product-card__comparison';

    const beforeP = document.createElement('p');
    beforeP.className = 'product-card__comparison-item';
    const beforeStrong = document.createElement('strong');
    beforeStrong.textContent = 'Antes: ';
    beforeP.append(beforeStrong, document.createTextNode(item.beforeState));

    const afterP = document.createElement('p');
    afterP.className = 'product-card__comparison-item product-card__comparison-item--after';
    const afterStrong = document.createElement('strong');
    afterStrong.textContent = 'Despues: ';
    afterP.append(afterStrong, document.createTextNode(item.afterState));

    comparisonBox.append(beforeP, afterP);
    body.appendChild(comparisonBox);
  }

  const specsList = document.createElement('dl');
  specsList.className = 'product-card__specs';

  const specs = [
    ['Madera / Proceso', item.woodType],
    ['Dimensiones', item.dimensions],
    ['Garantia', item.warranty]
  ];

  for (let i = 0; i < specs.length; i += 1) {
    const [label, value] = specs[i];
    const row = document.createElement('div');
    row.className = 'product-card__spec-row';

    const dt = document.createElement('dt');
    dt.textContent = label;

    const dd = document.createElement('dd');
    dd.textContent = value;

    row.append(dt, dd);
    specsList.appendChild(row);
  }

  const actions = document.createElement('div');
  actions.className = 'product-card__actions';

  const serviceTypeForItem =
    item.category === CATEGORIES.RESTAURACION ? 'Restauracion' : 'Mueble Nuevo';
  const woodTypeForItem = item.category === CATEGORIES.DEKO ? 'Deko' : 'Flor Morado';

  const { url: waUrl } = buildWhatsAppTriageUrl({
    serviceType: serviceTypeForItem,
    woodType: woodTypeForItem,
    dimensions: item.dimensions,
    productTitle: item.title
  });

  const quoteLink = document.createElement('a');
  quoteLink.className = 'btn btn--primary product-card__btn';
  quoteLink.href = waUrl;
  quoteLink.target = '_blank';
  quoteLink.rel = 'noopener noreferrer';
  quoteLink.setAttribute('aria-label', `Cotizar por WhatsApp: ${item.title}`);
  quoteLink.textContent = 'Cotizar en WhatsApp \u2192';

  const arButton = document.createElement('button');
  arButton.type = 'button';
  arButton.className = 'btn btn--secondary product-card__btn js-open-ar-modal';
  arButton.dataset.modelUrl = item.arModelUrl;
  arButton.dataset.modelTitle = item.arModelTitle || item.title;
  arButton.dataset.productTitle = item.title;
  arButton.setAttribute('aria-label', `Ver en tu espacio en Realidad Aumentada 3D: ${item.title}`);
  arButton.textContent = 'Ver en AR 3D \u2192';

  actions.append(quoteLink, arButton);
  body.append(specsList, actions);
  article.append(visualBanner, body);

  return article;
}

/**
 * Filtra el catalogo usando el indice Map precomputado en O(1) y renderiza los k elementos
 * resultantes en el contenedor DOM empleando un unico DocumentFragment en O(k).
 *
 * Complejidad Temporal: O(1) para la busqueda de clave en Map + O(k) para construir e insertar los k nodos.
 * Complejidad Espacial: O(k), donde k es el numero de tarjetas renderizadas en el DocumentFragment.
 *
 * @param {HTMLElement} container - Contenedor DOM de la grilla de productos.
 * @param {string} category - Clave de categoria a mostrar.
 * @param {HTMLElement|null} [statusNode=null] - Nodo con `aria-live="polite"` para anunciar el conteo.
 * @param {Map<string, ReadonlyArray<Object>>} [index=CATEGORY_INDEX] - Indice Map precomputado.
 * @returns {number} Cantidad k de elementos renderizados.
 */
export function renderFilteredCatalog(
  container,
  category = CATEGORIES.ALL,
  statusNode = null,
  index = CATEGORY_INDEX
) {
  if (!container) {
    return 0;
  }

  const normalizedCategory =
    typeof category === 'string' && CATEGORY_INDEX.has(category.trim().toLowerCase())
      ? category.trim().toLowerCase()
      : CATEGORIES.ALL;

  const matchedItems = getProductsByCategory(normalizedCategory, index);
  const fragment = document.createDocumentFragment();

  for (let i = 0; i < matchedItems.length; i += 1) {
    const cardNode = createProductCardNode(matchedItems[i]);
    fragment.appendChild(cardNode);
  }

  container.replaceChildren(fragment);

  if (statusNode) {
    const label = CATEGORY_LABELS[normalizedCategory] || CATEGORY_LABELS[CATEGORIES.ALL];
    statusNode.textContent = `Mostrando ${matchedItems.length} piezas en "${label}"`;
  }

  return matchedItems.length;
}
