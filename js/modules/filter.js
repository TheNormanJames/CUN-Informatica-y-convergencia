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
 * Crea un nodo SVG decorativo accesible que ilustra la pieza o taller sin depender de imagenes externas rotas.
 *
 * Complejidad Temporal: O(1), construccion determinista de un nodo vectorial estatico.
 * Complejidad Espacial: O(1), tamano constante del subarbol SVG.
 *
 * @param {string} visualTheme - Identificador visual del producto.
 * @returns {HTMLElement} Contenedor visual de la tarjeta.
 */
export function createProductVisualBanner(visualTheme = 'wood-warm') {
  const banner = document.createElement('div');
  banner.className = `product-card__visual product-card__visual--${visualTheme}`;
  banner.setAttribute('aria-hidden', 'true');

  const svgNS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(svgNS, 'svg');
  svg.setAttribute('viewBox', '0 0 360 180');
  svg.setAttribute('class', 'product-card__svg');
  svg.setAttribute('focusable', 'false');

  const bgRect = document.createElementNS(svgNS, 'rect');
  bgRect.setAttribute('x', '24');
  bgRect.setAttribute('y', '28');
  bgRect.setAttribute('width', '312');
  bgRect.setAttribute('height', '124');
  bgRect.setAttribute('rx', '12');
  bgRect.setAttribute('fill', 'currentColor');
  bgRect.setAttribute('fill-opacity', '0.12');

  const framePath = document.createElementNS(svgNS, 'path');
  framePath.setAttribute(
    'd',
    'M68 128 L68 78 L292 78 L292 128 M92 78 L92 52 L268 52 L268 78 M88 128 L88 148 M272 128 L272 148'
  );
  framePath.setAttribute('stroke', 'currentColor');
  framePath.setAttribute('stroke-width', '5');
  framePath.setAttribute('stroke-linecap', 'round');
  framePath.setAttribute('stroke-linejoin', 'round');
  framePath.setAttribute('fill', 'none');

  const grainLine = document.createElementNS(svgNS, 'path');
  grainLine.setAttribute('d', 'M100 102 Q180 88 260 102');
  grainLine.setAttribute('stroke', 'currentColor');
  grainLine.setAttribute('stroke-width', '3');
  grainLine.setAttribute('stroke-linecap', 'round');
  grainLine.setAttribute('fill', 'none');

  svg.append(bgRect, framePath, grainLine);
  banner.appendChild(svg);
  return banner;
}

/**
 * Construye el nodo DOM `<article>` para un item del catalogo sin usar innerHTML en bucles.
 *
 * Complejidad Temporal: O(1), cantidad fija de nodos DOM creados por tarjeta.
 * Complejidad Espacial: O(1), estructura DOM acotada por tarjeta.
 *
 * @param {Object} item - Objeto de producto o servicio del catalogo.
 * @returns {HTMLElement} Elemento `<article>` listo para agregarse al DocumentFragment.
 */
export function createProductCardNode(item) {
  const article = document.createElement('article');
  article.className = 'product-card';
  article.dataset.productId = item.id;
  article.dataset.category = item.category;

  const visualBanner = createProductVisualBanner(item.visualTheme);

  const badgeSpan = document.createElement('span');
  badgeSpan.className = 'product-card__badge';
  badgeSpan.textContent = item.badge;
  visualBanner.appendChild(badgeSpan);

  const body = document.createElement('div');
  body.className = 'product-card__body';

  const categoryTag = document.createElement('p');
  categoryTag.className = 'product-card__category';
  categoryTag.textContent = item.categoryLabel;

  const title = document.createElement('h3');
  title.className = 'product-card__title';
  title.textContent = item.title;

  const description = document.createElement('p');
  description.className = 'product-card__description';
  description.textContent = item.description;

  body.append(categoryTag, title, description);

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
    ['Medidas / Tiempo', item.dimensions],
    ['Respaldo', item.warranty]
  ];

  for (let i = 0; i < specs.length; i += 1) {
    const [label, value] = specs[i];
    const row = document.createElement('div');
    row.className = 'product-card__spec-row';

    const dt = document.createElement('dt');
    dt.textContent = `${label}:`;

    const dd = document.createElement('dd');
    dd.textContent = value;

    row.append(dt, dd);
    specsList.appendChild(row);
  }

  const priceTag = document.createElement('p');
  priceTag.className = 'product-card__price';
  priceTag.textContent = item.priceReference;

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
  quoteLink.className = 'btn btn--whatsapp product-card__btn';
  quoteLink.href = waUrl;
  quoteLink.target = '_blank';
  quoteLink.rel = 'noopener noreferrer';
  quoteLink.setAttribute('aria-label', `Cotizar por WhatsApp: ${item.title}`);
  quoteLink.textContent = 'Cotizar en WhatsApp';

  const arButton = document.createElement('button');
  arButton.type = 'button';
  arButton.className = 'btn btn--secondary product-card__btn js-open-ar-modal';
  arButton.dataset.modelUrl = item.arModelUrl;
  arButton.dataset.modelTitle = item.arModelTitle || item.title;
  arButton.dataset.productTitle = item.title;
  arButton.setAttribute('aria-label', `Ver en tu espacio en Realidad Aumentada 3D: ${item.title}`);
  arButton.textContent = 'Ver en tu espacio (AR 3D)';

  actions.append(quoteLink, arButton);
  body.append(specsList, priceTag, actions);
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
    statusNode.textContent = `Mostrando ${matchedItems.length} opciones en "${label}" (consulta O(1) en indice Map).`;
  }

  return matchedItems.length;
}
