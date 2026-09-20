/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./assets/resource/PLANTILLA.docx":
/*!****************************************!*\
  !*** ./assets/resource/PLANTILLA.docx ***!
  \****************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "templates/PLANTILLA.docx";

/***/ }),

/***/ "./src/taskpane/taskpane.js?4727":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "80d428829b3f7a8599fb.js";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "2ab4d83191e031cc238f.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		if (!(moduleId in __webpack_modules__)) {
/******/ 			delete __webpack_module_cache__[moduleId];
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = (typeof document !== 'undefined' && document.baseURI) || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _assets_resource_PLANTILLA_docx__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../assets/resource/PLANTILLA.docx */ "./assets/resource/PLANTILLA.docx");
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */



// Variables globales para el formulario
var evaluationData = {
  name: '',
  objectives: '',
  course: '',
  date: '',
  unit: '',
  topic: '',
  version: 1
};

/**
 * Inicialización del panel cuando Office está listo
 */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    // Ocultar mensaje de sideload y mostrar formulario
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // Establecer fecha por defecto (hoy)
    var today = new Date().toISOString().split('T')[0];
    document.getElementById("eval-date").value = today;

    // Configurar event listeners
    setupEventListeners();
    console.log("Panel de evaluación CSJ inicializado correctamente");
  }
});

/**
 * Configura los event listeners para los botones y campos del formulario
 */
function setupEventListeners() {
  // Botón Crear Versión 1
  var btnCreate = document.getElementById("btn-create-eval");
  if (btnCreate) {
    btnCreate.addEventListener("click", handleCreateEvaluation);
  } else {
    console.error("No se encontró el botón btn-create-eval");
  }

  // Botón Cancelar
  var btnCancel = document.getElementById("btn-cancel");
  if (btnCancel) {
    btnCancel.addEventListener("click", handleCancel);
  } else {
    console.error("No se encontró el botón btn-cancel");
  }

  // Listeners para validación en tiempo real
  var formFields = ['eval-name', 'eval-objectives', 'eval-course', 'eval-date', 'eval-unit', 'eval-topic'];
  formFields.forEach(function (fieldId) {
    var field = document.getElementById(fieldId);
    if (field) {
      field.addEventListener('input', function () {
        // Limpiar mensajes de error cuando el usuario escribe
        hideMessage();
      });
      field.addEventListener('blur', function () {
        // Validar campo al perder foco
        validateField(fieldId);
      });
    } else {
      console.error("No se encontr\xF3 el campo ".concat(fieldId));
    }
  });
}

/**
 * Valida un campo individual del formulario
 * @param {string} fieldId - ID del campo a validar
 * @returns {boolean} - true si es válido, false si hay error
 */
function validateField(fieldId) {
  var field = document.getElementById(fieldId);
  if (!field) return true;
  var value = field.value.trim();

  // Todos los campos son requeridos
  if (value === '') {
    field.style.borderColor = '#a80000';
    return false;
  } else {
    field.style.borderColor = '#8a8886';
    return true;
  }
}

/**
 * Valida todos los campos del formulario
 * @returns {boolean} - true si todos los campos son válidos
 */
function validateForm() {
  var fields = ['eval-name', 'eval-objectives', 'eval-course', 'eval-date', 'eval-unit', 'eval-topic'];
  var isValid = true;
  fields.forEach(function (fieldId) {
    if (!validateField(fieldId)) {
      isValid = false;
    }
  });
  return isValid;
}

/**
 * Recopila los datos del formulario
 * @returns {Object} - Objeto con los datos de la evaluación
 */
function getFormData() {
  return {
    name: document.getElementById("eval-name").value.trim(),
    objectives: document.getElementById("eval-objectives").value.trim(),
    course: document.getElementById("eval-course").value.trim(),
    date: document.getElementById("eval-date").value,
    unit: document.getElementById("eval-unit").value.trim(),
    topic: document.getElementById("eval-topic").value.trim(),
    version: Number(document.getElementById("eval-version").value) || 1
  };
}
function loadTemplateBase64() {
  return _loadTemplateBase.apply(this, arguments);
}
function _loadTemplateBase() {
  _loadTemplateBase = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    var response, bytes, binary, index, _t, _t2;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          _context.n = 1;
          return fetch(_assets_resource_PLANTILLA_docx__WEBPACK_IMPORTED_MODULE_0__);
        case 1:
          response = _context.v;
          if (response.ok) {
            _context.n = 2;
            break;
          }
          throw new Error('No se pudo cargar la plantilla de evaluación.');
        case 2:
          _t = Uint8Array;
          _context.n = 3;
          return response.arrayBuffer();
        case 3:
          _t2 = _context.v;
          bytes = new _t(_t2);
          binary = '';
          for (index = 0; index < bytes.length; index += 1) binary += String.fromCharCode(bytes[index]);
          return _context.a(2, btoa(binary));
      }
    }, _callee);
  }));
  return _loadTemplateBase.apply(this, arguments);
}
function replaceText(_x, _x2, _x3, _x4) {
  return _replaceText.apply(this, arguments);
}
/**
 * Maneja el evento de crear evaluación
 */
function _replaceText() {
  _replaceText = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(context, body, searchText, replacement) {
    var ranges;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          ranges = body.search(searchText, {
            matchCase: false,
            matchWholeWord: false
          });
          ranges.load('items');
          _context2.n = 1;
          return context.sync();
        case 1:
          ranges.items.forEach(function (range) {
            return range.insertText(replacement, Word.InsertLocation.replace);
          });
        case 2:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _replaceText.apply(this, arguments);
}
function handleCreateEvaluation() {
  return _handleCreateEvaluation.apply(this, arguments);
}
/**
 * Crea una nueva evaluación basada en la plantilla PLANTILLA.docx
 * @param {Object} data - Datos de la evaluación
 */
function _handleCreateEvaluation() {
  _handleCreateEvaluation = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var btnCreate, originalText, _t3;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          console.log("Iniciando creación de evaluación...");

          // Limpiar mensajes previos
          hideMessage();

          // Validar formulario
          if (validateForm()) {
            _context3.n = 1;
            break;
          }
          console.log("Validación fallida");
          showMessage("Por favor complete todos los campos requeridos.", "error");
          return _context3.a(2);
        case 1:
          console.log("Validación exitosa, obteniendo datos...");

          // Obtener datos del formulario
          evaluationData = getFormData();
          console.log("Datos obtenidos:", evaluationData);

          // Mostrar estado de carga
          btnCreate = document.getElementById("btn-create-eval");
          originalText = btnCreate.innerHTML;
          btnCreate.disabled = true;
          btnCreate.innerHTML = '<span class="ms-Button-label"><i class="ms-Icon ms-Icon--Refresh"></i> Procesando...</span>';
          _context3.p = 2;
          console.log("Llamando a createEvaluationFromTemplate...");
          // Llamar a la función que crea el documento desde la plantilla
          _context3.n = 3;
          return createEvaluationFromTemplate(evaluationData);
        case 3:
          console.log("Evaluación creada exitosamente");

          // Mostrar mensaje de éxito
          showMessage("Evaluaci\xF3n \"".concat(evaluationData.name, "\" creada exitosamente."), "success");
          _context3.n = 5;
          break;
        case 4:
          _context3.p = 4;
          _t3 = _context3.v;
          console.error("Error al crear evaluación:", _t3);
          console.error("Detalle del error:", _t3.message);
          console.error("Stack trace:", _t3.stack);
          showMessage("Error al crear la evaluación: " + _t3.message, "error");
        case 5:
          _context3.p = 5;
          // Restaurar botón
          btnCreate.disabled = false;
          btnCreate.innerHTML = originalText;
          return _context3.f(5);
        case 6:
          return _context3.a(2);
      }
    }, _callee3, null, [[2, 4, 5, 6]]);
  }));
  return _handleCreateEvaluation.apply(this, arguments);
}
function createEvaluationFromTemplate(_x5) {
  return _createEvaluationFromTemplate.apply(this, arguments);
}
/**
 * Formatea una fecha de YYYY-MM-DD a DD/MM/YYYY
 * @param {string} dateString - Fecha en formato YYYY-MM-DD
 * @returns {string} - Fecha formateada
 */
function _createEvaluationFromTemplate() {
  _createEvaluationFromTemplate = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(data) {
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.n) {
        case 0:
          console.log("Iniciando creación de documento Word...");
          return _context5.a(2, Word.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(context) {
              var body, _t4, _t5;
              return _regenerator().w(function (_context4) {
                while (1) switch (_context4.p = _context4.n) {
                  case 0:
                    _context4.p = 0;
                    console.log("Dentro de Word.run, insertando contenido...");
                    body = context.document.body;
                    _t4 = body;
                    _context4.n = 1;
                    return loadTemplateBase64();
                  case 1:
                    _t4.insertFileFromBase64.call(_t4, _context4.v, Word.InsertLocation.replace);
                    _context4.n = 2;
                    return context.sync();
                  case 2:
                    _context4.n = 3;
                    return replaceText(context, body, 'PRUEBA', data.name);
                  case 3:
                    _context4.n = 4;
                    return replaceText(context, body, '[UNIDAD]', data.unit);
                  case 4:
                    _context4.n = 5;
                    return replaceText(context, body, '[Describe los Objetivos]', data.objectives);
                  case 5:
                    _context4.n = 6;
                    return replaceText(context, body, '__ /___ / 2026', formatDate(data.date));
                  case 6:
                    _context4.n = 7;
                    return replaceText(context, body, '[Curso]', data.course);
                  case 7:
                    _context4.n = 8;
                    return replaceText(context, body, '[Tema]', data.topic);
                  case 8:
                    _context4.n = 9;
                    return replaceText(context, body, '[Versión]', data.version);
                  case 9:
                    _context4.n = 10;
                    return context.sync();
                  case 10:
                    console.log("Documento Word creado exitosamente");
                    _context4.n = 12;
                    break;
                  case 11:
                    _context4.p = 11;
                    _t5 = _context4.v;
                    console.error("Error dentro de Word.run:", _t5);
                    console.error("Nombre del error:", _t5.name);
                    console.error("Mensaje del error:", _t5.message);
                    console.error("Stack trace:", _t5.stack);
                    throw _t5;
                  case 12:
                    return _context4.a(2);
                }
              }, _callee4, null, [[0, 11]]);
            }));
            return function (_x6) {
              return _ref.apply(this, arguments);
            };
          }()));
      }
    }, _callee5);
  }));
  return _createEvaluationFromTemplate.apply(this, arguments);
}
function formatDate(dateString) {
  if (!dateString) return '';
  var _dateString$split = dateString.split('-'),
    _dateString$split2 = _slicedToArray(_dateString$split, 3),
    year = _dateString$split2[0],
    month = _dateString$split2[1],
    day = _dateString$split2[2];
  return "".concat(day, "/").concat(month, "/").concat(year);
}

/**
 * Maneja el evento de cancelar
 */
function handleCancel() {
  // Resetear formulario
  document.getElementById("eval-name").value = '';
  document.getElementById("eval-objectives").value = '';
  document.getElementById("eval-course").value = '';
  document.getElementById("eval-unit").value = '';
  document.getElementById("eval-topic").value = '';
  document.getElementById("eval-version").value = '1';

  // Restablecer fecha a hoy
  var today = new Date().toISOString().split('T')[0];
  document.getElementById("eval-date").value = today;

  // Limpiar estilos de validación
  var formFields = document.querySelectorAll('.ms-TextField-field');
  formFields.forEach(function (field) {
    field.style.borderColor = '#8a8886';
  });

  // Ocultar mensajes
  hideMessage();
  console.log("Formulario reseteado");
}

/**
 * Muestra un mensaje en el área de mensajes
 * @param {string} text - Texto del mensaje
 * @param {string} type - Tipo de mensaje: 'success', 'error', o 'info'
 */
function showMessage(text) {
  var type = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 'info';
  var messageArea = document.getElementById("message-area");
  var messageContent = document.getElementById("message-content");
  messageContent.textContent = text;
  messageContent.className = 'ms-Callout-main';
  messageContent.classList.add("message-".concat(type));
  messageArea.style.display = 'block';
}

/**
 * Oculta el área de mensajes
 */
function hideMessage() {
  var messageArea = document.getElementById("message-area");
  messageArea.style.display = 'none';
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.js */ "./src/taskpane/taskpane.js?4727"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- Panel para crear evaluaciones basadas en PLANTILLA.docx -->\r\n\r\n<!DOCTYPE html>\r\n<html lang=\"es-CL\">\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Crear Evaluación</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- Fluent UI -->\r\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-Fabric\">\r\n    <header class=\"panel-header ms-bgColor-themeDark\">\r\n        <!-- <img width=\"60\" height=\"60\" src=\"../../assets/logo-filled.png\" alt=\"CSJ\" title=\"CSJ\" /> -->\r\n        <h1 class=\"ms-font-xl ms-fontColor-white\">Nueva Evaluación</h1>\r\n        <p class=\"ms-font-s ms-fontColor-white\">Basada en PLANTILLA.docx</p>\r\n    </header>\r\n\r\n    <main class=\"panel-content\">\r\n        <!-- Mensaje de sideload para desarrollo -->\r\n        <section id=\"sideload-msg\" class=\"ms-welcome__main\">\r\n            <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\r\n        </section>\r\n\r\n        <!-- Formulario de creación de evaluación -->\r\n        <div id=\"app-body\" class=\"evaluation-form\" style=\"display: none;\">\r\n            \r\n            <div class=\"ms-Grid ms-Grid-row\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <h2 class=\"ms-font-l ms-fontWeight-semibold ms-borderBottom-themePrimary ms-pb-s\">\r\n                        <i class=\"ms-Icon ms-Icon--FileComment\"></i>\r\n                        Datos de la Evaluación\r\n                    </h2>\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Campo: Nombre de la evaluación -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\">\r\n                        Nombre de la Evaluación *\r\n                    </label>\r\n                    <input \r\n                        type=\"text\" \r\n                        id=\"eval-name\" \r\n                        class=\"ms-TextField-field ms-TextField\" \r\n                        placeholder=\"Ej: Prueba de Matemáticas - Unidad 1\"\r\n                        aria-label=\"Nombre de la evaluación\"\r\n                        required\r\n                    />\r\n                </div>\r\n            </div>\r\n\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm6\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\" for=\"eval-version\">Versión</label>\r\n                    <input type=\"number\" id=\"eval-version\" class=\"ms-TextField-field ms-TextField\" min=\"1\" value=\"1\" aria-label=\"Versión de la evaluación\" />\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Campo: Objetivos -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\">\r\n                        Objetivos de Aprendizaje *\r\n                    </label>\r\n                    <textarea \r\n                        id=\"eval-objectives\" \r\n                        class=\"ms-TextField-field ms-TextField\" \r\n                        placeholder=\"Describa los objetivos de aprendizaje que se evaluarán...\"\r\n                        rows=\"4\"\r\n                        aria-label=\"Objetivos de aprendizaje\"\r\n                        required\r\n                    ></textarea>\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Campo: Curso -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\">\r\n                        Curso *\r\n                    </label>\r\n                    <input \r\n                        type=\"text\" \r\n                        id=\"eval-course\" \r\n                        class=\"ms-TextField-field ms-TextField\" \r\n                        placeholder=\"Ej: 5° Básico A\"\r\n                        aria-label=\"Curso\"\r\n                        required\r\n                    />\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Campo: Fecha -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\">\r\n                        Fecha *\r\n                    </label>\r\n                    <input \r\n                        type=\"date\" \r\n                        id=\"eval-date\" \r\n                        class=\"ms-TextField-field ms-TextField\" \r\n                        aria-label=\"Fecha de la evaluación\"\r\n                        required\r\n                    />\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Campo: Unidad -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\">\r\n                        Unidad *\r\n                    </label>\r\n                    <input \r\n                        type=\"text\" \r\n                        id=\"eval-unit\" \r\n                        class=\"ms-TextField-field ms-TextField\" \r\n                        placeholder=\"Ej: Unidad 1 - Números y Operaciones\"\r\n                        aria-label=\"Unidad\"\r\n                        required\r\n                    />\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Campo: Tema a Evaluar -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mb-m\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <label class=\"ms-Label ms-fontWeight-semibold ms-mb-xs\">\r\n                        Tema a Evaluar *\r\n                    </label>\r\n                    <input \r\n                        type=\"text\" \r\n                        id=\"eval-topic\" \r\n                        class=\"ms-TextField-field ms-TextField\" \r\n                        placeholder=\"Ej: Suma y resta de fracciones\"\r\n                        aria-label=\"Tema a evaluar\"\r\n                        required\r\n                    />\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Botones de acción -->\r\n            <div class=\"ms-Grid ms-Grid-row ms-mt-l ms-pb-l\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <div class=\"form-actions\">\r\n                        <button \r\n                            id=\"btn-create-eval\" \r\n                            class=\"ms-Button ms-Button--primary ms-Button--hero\"\r\n                        >\r\n                            <span class=\"ms-Button-label\">\r\n                                <i class=\"ms-Icon ms-Icon--Plus\"></i>\r\n                                Crear versión\r\n                            </span>\r\n                        </button>\r\n                        <button \r\n                            id=\"btn-cancel\" \r\n                            class=\"ms-Button ms-Button--default\"\r\n                        >\r\n                            <span class=\"ms-Button-label\">Cancelar</span>\r\n                        </button>\r\n                    </div>\r\n                </div>\r\n            </div>\r\n\r\n            <!-- Área de mensajes -->\r\n            <div id=\"message-area\" class=\"ms-Grid ms-Grid-row ms-mt-s\" style=\"display: none;\">\r\n                <div class=\"ms-Grid-col ms-sm12\">\r\n                    <div id=\"message-content\" class=\"ms-Callout-main\"></div>\r\n                </div>\r\n            </div>\r\n\r\n        </div>\r\n    </main>\r\n\r\n    <!-- Script principal -->\r\n    <" + "script type=\"module\" src=\"" + ___HTML_LOADER_IMPORT_1___ + "\"><" + "/script>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map