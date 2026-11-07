/**
 * @fileoverview Punto de entrada principal de Clasicos Morandua.
 * Orquesta navegacion accesible, filtros de catalogo O(1), modal WebAR 3D,
 * asistente de pre-triaje IA y microinteracciones fisicas de aparicion suave.
 */

import {
  CATEGORIES,
  buildWhatsAppTriageUrl
} from './data/catalog.js';
import { renderFilteredCatalog } from './modules/filter.js';

/**
 * Inicializa el menu de navegacion responsivo con soporte para teclado y cierre con Escape.
 */
function initMobileNavigation() {
  const navToggle = document.getElementById('nav-toggle');
  const navMenu = document.getElementById('primary-nav');

  if (!navToggle || !navMenu) {
    return;
  }

  const closeMenu = () => {
    navToggle.setAttribute('aria-expanded', 'false');
    navMenu.classList.remove('is-open');
  };

  navToggle.addEventListener('click', () => {
    const isExpanded = navToggle.getAttribute('aria-expanded') === 'true';
    const nextState = !isExpanded;
    navToggle.setAttribute('aria-expanded', String(nextState));
    navMenu.classList.toggle('is-open', nextState);
  });

  navMenu.addEventListener('click', (event) => {
    const link = event.target.closest('a');
    if (link) {
      const allLinks = navMenu.querySelectorAll('.primary-nav__link');
      for (let i = 0; i < allLinks.length; i += 1) {
        allLinks[i].classList.toggle('is-active', allLinks[i] === link);
      }
      closeMenu();
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && navMenu.classList.contains('is-open')) {
      closeMenu();
      navToggle.focus();
    }
  });
}

/**
 * Inicializa la barra de filtros interactiva conectada al indice Map O(1).
 */
function initCatalogFiltering() {
  const catalogGrid = document.getElementById('catalog-grid');
  const filterControls = document.getElementById('catalog-filters');
  const catalogStatus = document.getElementById('catalog-status');

  if (!catalogGrid || !filterControls) {
    return;
  }

  renderFilteredCatalog(catalogGrid, CATEGORIES.ALL, catalogStatus);

  filterControls.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-filter]');
    if (!button) {
      return;
    }

    const selectedCategory = button.dataset.filter || CATEGORIES.ALL;
    const allButtons = filterControls.querySelectorAll('button[data-filter]');

    for (let i = 0; i < allButtons.length; i += 1) {
      const btn = allButtons[i];
      const isCurrent = btn === button;
      btn.classList.toggle('is-active', isCurrent);
      btn.setAttribute('aria-pressed', String(isCurrent));
    }

    renderFilteredCatalog(catalogGrid, selectedCategory, catalogStatus);
  });
}

/**
 * Inicializa el modal accesible de Realidad Aumentada 3D con Google `<model-viewer>`.
 */
function initWebARModal() {
  const modal = document.getElementById('ar-modal');
  const closeBtn = document.getElementById('ar-modal-close');
  const backdrop = document.getElementById('ar-modal-backdrop');
  const modelViewer = document.getElementById('ar-model-viewer');
  const modalTitle = document.getElementById('ar-modal-title');
  const modalProductLabel = document.getElementById('ar-modal-product-name');
  const modalWaLink = document.getElementById('ar-modal-wa-link');
  const openHeroArBtn = document.getElementById('open-hero-ar-btn');

  if (!modal) {
    return;
  }

  let lastFocusedElement = null;

  const openModal = ({ modelUrl, modelTitle, productTitle }) => {
    lastFocusedElement = document.activeElement;

    if (modelViewer && modelUrl) {
      if (modelViewer.getAttribute('src') !== modelUrl) {
        modelViewer.setAttribute('src', modelUrl);
      }
      modelViewer.setAttribute('alt', modelTitle || 'Modelo 3D en madera maciza de Clasicos Morandua');
    }

    if (modalTitle) {
      modalTitle.textContent = 'Ver en tu espacio (Realidad Aumentada).';
    }

    if (modalProductLabel) {
      modalProductLabel.textContent =
        productTitle || 'Muestra Demostrativa en Flor Morado y Linea Deko';
    }

    if (modalWaLink) {
      const { url } = buildWhatsAppTriageUrl({
        serviceType: 'Mueble Nuevo',
        woodType: 'Flor Morado',
        dimensions: 'Validado en visor 3D / Realidad Aumentada',
        productTitle: productTitle || 'Pieza vista en Realidad Aumentada 3D'
      });
      modalWaLink.href = url;
    }

    modal.hidden = false;
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');

    if (closeBtn) {
      closeBtn.focus();
    }
  };

  const closeModal = () => {
    modal.hidden = true;
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('modal-open');

    if (lastFocusedElement && typeof lastFocusedElement.focus === 'function') {
      lastFocusedElement.focus();
    }
  };

  document.addEventListener('click', (event) => {
    const trigger = event.target.closest('.js-open-ar-modal');
    if (trigger) {
      event.preventDefault();
      openModal({
        modelUrl: trigger.dataset.modelUrl || 'https://modelviewer.dev/shared-assets/models/Chair.glb',
        modelTitle: trigger.dataset.modelTitle,
        productTitle: trigger.dataset.productTitle
      });
    }
  });

  if (openHeroArBtn) {
    openHeroArBtn.addEventListener('click', () => {
      openModal({
        modelUrl: 'https://modelviewer.dev/shared-assets/models/Chair.glb',
        modelTitle: 'Silla Maestra en Flor Morado (Demostracion WebAR 3D)',
        productTitle: 'Silla Maestra Comedor Imperial - Flor Morado'
      });
    });
  }

  if (closeBtn) {
    closeBtn.addEventListener('click', closeModal);
  }

  if (backdrop) {
    backdrop.addEventListener('click', closeModal);
  }

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !modal.hidden) {
      closeModal();
    }
  });
}

/**
 * Inicializa el widget "Asistente de Cotizacion Rapida (Demo IA)" que genera el enlace
 * de pre-triaje hacia WhatsApp Business en tiempo constante O(1).
 */
function initAITriageAssistant() {
  const form = document.getElementById('ai-triage-form');
  const previewBox = document.getElementById('triage-preview-text');
  const badgeNode = document.getElementById('triage-code-badge');
  const waButton = document.getElementById('triage-whatsapp-link');

  if (!form || !previewBox || !waButton) {
    return;
  }

  const updateTriageOutput = () => {
    const formData = new FormData(form);
    const serviceType = String(formData.get('serviceType') || 'Mueble Nuevo');
    const woodType = String(formData.get('woodType') || 'Flor Morado');
    const dimensions = String(formData.get('dimensions') || '');
    const productTitle = String(formData.get('productTitle') || 'Cotizacion personalizada desde Asistente IA');

    const result = buildWhatsAppTriageUrl({
      serviceType,
      woodType,
      dimensions,
      productTitle
    });

    previewBox.textContent = result.messagePreview;
    waButton.href = result.url;

    if (badgeNode) {
      badgeNode.textContent = result.triageCode;
    }
  };

  form.addEventListener('input', updateTriageOutput);
  form.addEventListener('change', updateTriageOutput);
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    updateTriageOutput();
    window.open(waButton.href, '_blank', 'noopener,noreferrer');
  });

  updateTriageOutput();
}

/**
 * Inicializa microinteracciones fisicas de aparicion al hacer scroll usando IntersectionObserver.
 */
function initScrollReveal() {
  const elements = document.querySelectorAll('.reveal-on-scroll');
  if (!('IntersectionObserver' in window)) {
    for (let i = 0; i < elements.length; i += 1) {
      elements[i].classList.add('is-visible');
    }
    return;
  }

  const observer = new IntersectionObserver(
    (entries) => {
      for (let i = 0; i < entries.length; i += 1) {
        if (entries[i].isIntersecting) {
          entries[i].target.classList.add('is-visible');
          observer.unobserve(entries[i].target);
        }
      }
    },
    { threshold: 0.08, rootMargin: '0px 0px -24px 0px' }
  );

  for (let i = 0; i < elements.length; i += 1) {
    observer.observe(elements[i]);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  initMobileNavigation();
  initCatalogFiltering();
  initWebARModal();
  initAITriageAssistant();
  initScrollReveal();
});
