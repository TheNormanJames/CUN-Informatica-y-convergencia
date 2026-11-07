/**
 * @fileoverview Suite de pruebas unitarias automatizadas para Clasicos Morandua.
 * Ejecuta con Node Native Test Runner: `node --test tests/catalog.test.mjs`
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  BUSINESS_CONTACT,
  CATEGORIES,
  CATALOG_ITEMS,
  CATEGORY_INDEX,
  buildCategoryIndex,
  getProductsByCategory,
  sanitizeInputText,
  buildWhatsAppTriageUrl
} from '../js/data/catalog.js';

test('Fase 1 y 2: Existencia y campos obligatorios de todos los productos del catalogo (>= 6 items)', () => {
  assert.ok(Array.isArray(CATALOG_ITEMS), 'CATALOG_ITEMS debe ser un arreglo');
  assert.ok(CATALOG_ITEMS.length >= 6, 'El catalogo debe contener al menos 6 items representativos');

  const requiredFields = [
    'id',
    'title',
    'category',
    'categoryLabel',
    'woodType',
    'dimensions',
    'priceReference',
    'warranty',
    'description',
    'badge',
    'arModelUrl'
  ];

  const seenIds = new Set();

  for (const item of CATALOG_ITEMS) {
    for (const field of requiredFields) {
      assert.ok(
        typeof item[field] === 'string' && item[field].trim().length > 0,
        `El campo obligatorio "${field}" debe ser una cadena no vacia en el item ${item.id}`
      );
    }

    assert.equal(seenIds.has(item.id), false, `El ID "${item.id}" no debe estar duplicado`);
    seenIds.add(item.id);
    assert.ok(
      item.arModelUrl.startsWith('https://') && item.arModelUrl.endsWith('.glb'),
      `El item ${item.id} debe incluir una URL segura hacia un modelo 3D .glb`
    );
  }
});

test('Fase 2: CATEGORY_INDEX es un Map precomputado y devuelve exactamente los elementos esperados en O(1)', () => {
  assert.ok(CATEGORY_INDEX instanceof Map, 'CATEGORY_INDEX debe ser una instancia de Map');

  const allProducts = getProductsByCategory(CATEGORIES.ALL);
  assert.equal(allProducts.length, CATALOG_ITEMS.length, 'La clave "todos" debe retornar el total del catalogo');

  const florMoradoProducts = getProductsByCategory(CATEGORIES.FLOR_MORADO);
  assert.equal(florMoradoProducts.length, 2, 'Debe haber exactamente 2 productos en Linea Flor Morado');
  for (const item of florMoradoProducts) {
    assert.equal(item.category, CATEGORIES.FLOR_MORADO);
  }

  const dekoProducts = getProductsByCategory(CATEGORIES.DEKO);
  assert.equal(dekoProducts.length, 2, 'Debe haber exactamente 2 productos en Linea Deko');
  for (const item of dekoProducts) {
    assert.equal(item.category, CATEGORIES.DEKO);
  }

  const restauracionProducts = getProductsByCategory(CATEGORIES.RESTAURACION);
  assert.equal(restauracionProducts.length, 2, 'Debe haber exactamente 2 servicios en Restauracion');
  for (const item of restauracionProducts) {
    assert.equal(item.category, CATEGORIES.RESTAURACION);
    assert.ok(item.beforeState && item.afterState, 'Los casos de restauracion deben documentar Antes y Despues');
  }
});

test('Fase 2 (Caso Limite): getProductsByCategory maneja categorias inexistentes o entradas invalidas sin lanzar excepciones', () => {
  assert.doesNotThrow(() => {
    const unknown = getProductsByCategory('categoria-inexistente-xyz');
    assert.ok(Array.isArray(unknown));
    assert.equal(unknown.length, 0);
  });

  assert.deepEqual(getProductsByCategory(''), []);
  assert.deepEqual(getProductsByCategory(null), []);
  assert.deepEqual(getProductsByCategory(undefined), []);
  assert.deepEqual(getProductsByCategory(12345), []);
  assert.deepEqual(getProductsByCategory('  FLOR-MORADO  ').length, 2);

  const emptyIndex = buildCategoryIndex([]);
  assert.equal(getProductsByCategory(CATEGORIES.ALL, emptyIndex).length, 0);
  assert.equal(getProductsByCategory(CATEGORIES.FLOR_MORADO, emptyIndex).length, 0);
});

test('Fase 1 y 3: Generacion correcta, estructurada y sanitizada de enlaces de contacto y WhatsApp Business', () => {
  assert.equal(BUSINESS_CONTACT.phoneDisplay, '+57 300 000 0000');
  assert.equal(BUSINESS_CONTACT.phoneRaw, '573000000000');
  assert.equal(BUSINESS_CONTACT.email, 'contacto@clasicosmorandua.ejemplo.co');

  const newFurnitureTriage = buildWhatsAppTriageUrl({
    serviceType: 'Mueble Nuevo',
    woodType: 'Flor Morado',
    dimensions: '200cm x 110cm <script>alert(1)</script>',
    productTitle: 'Comedor Imperial 6 Puestos'
  });

  assert.equal(newFurnitureTriage.triageCode, 'TRIAJE-IA-NUEVO');
  assert.ok(
    newFurnitureTriage.url.startsWith('https://wa.me/573000000000?text='),
    'El enlace debe apuntar al numero sanitizado +57 300 000 0000 en wa.me'
  );
  assert.equal(
    newFurnitureTriage.messagePreview.includes('<script>'),
    false,
    'El generador debe sanitizar etiquetas angulares maliciosas'
  );
  assert.ok(
    newFurnitureTriage.messagePreview.includes('Flor Morado Macizo'),
    'El mensaje debe incluir la madera seleccionada'
  );

  const restorationTriage = buildWhatsAppTriageUrl({
    serviceType: 'Restauracion',
    woodType: 'Deko',
    dimensions: '',
    productTitle: ''
  });

  assert.equal(restorationTriage.triageCode, 'TRIAJE-IA-RESTAURACION');
  assert.ok(restorationTriage.messagePreview.includes('Restauración / Refacción'));
  assert.ok(restorationTriage.messagePreview.includes('Línea Deko'));

  const sanitizedFallback = sanitizeInputText('   <>   ', 'Valor por defecto');
  assert.equal(sanitizedFallback, 'Valor por defecto');
});

