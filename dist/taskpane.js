/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./assets/logo-filled.png":
/*!********************************!*\
  !*** ./assets/logo-filled.png ***!
  \********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "assets/logo-filled.png";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "81d2cb881530cbb00e6e.css";

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
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
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
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
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
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   run: function() { return /* binding */ run; }
/* harmony export */ });
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Afficher Hello world à l'ouverture
    var hello = document.getElementById("hello-text");
    if (hello) {
      hello.textContent = "Hello world";
    }
    document.getElementById("run").onclick = run;
    var apiKeyInput = document.getElementById("api-key");
    var promptInput = document.getElementById("prompt-input");
    var sendBtn = document.getElementById("send-prompt");
    var addMailBtn = document.getElementById("add-mail-content");
    var addCategoriesBtn = document.getElementById("add-mail-categories");
    var output = document.getElementById("response-output");
    if (apiKeyInput && promptInput && sendBtn && output) {
      // charger la clé si déjà sauvegardée en localStorage (dev uniquement)
      try {
        var saved = localStorage.getItem("openai_api_key");
        if (saved) apiKeyInput.value = saved;
      } catch (_unused) {}
      sendBtn.addEventListener("click", /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
        var apiKey, prompt, _data$choices, resp, errText, data, answer, _t;
        return _regenerator().w(function (_context) {
          while (1) switch (_context.p = _context.n) {
            case 0:
              apiKey = apiKeyInput.value.trim();
              prompt = promptInput.value.trim();
              if (!(!apiKey || !prompt)) {
                _context.n = 1;
                break;
              }
              output.textContent = "Veuillez saisir la clé API et un prompt.";
              return _context.a(2);
            case 1:
              _context.p = 1;
              // sauvegarde locale (dev)
              try {
                localStorage.setItem("openai_api_key", apiKey);
              } catch (_unused2) {}
              output.textContent = "Envoi en cours...";

              // Appel API OpenAI (Chat Completions)
              _context.n = 2;
              return fetch("https://api.openai.com/v1/chat/completions", {
                method: "POST",
                headers: {
                  "Content-Type": "application/json",
                  "Authorization": "Bearer ".concat(apiKey)
                },
                body: JSON.stringify({
                  model: "gpt-3.5-turbo",
                  messages: [{
                    role: "system",
                    content: "You are a helpful assistant."
                  }, {
                    role: "user",
                    content: prompt
                  }],
                  temperature: 0.7
                })
              });
            case 2:
              resp = _context.v;
              if (resp.ok) {
                _context.n = 4;
                break;
              }
              _context.n = 3;
              return resp.text();
            case 3:
              errText = _context.v;
              output.textContent = "Erreur API (".concat(resp.status, "): ").concat(errText);
              return _context.a(2);
            case 4:
              _context.n = 5;
              return resp.json();
            case 5:
              data = _context.v;
              answer = (data === null || data === void 0 || (_data$choices = data.choices) === null || _data$choices === void 0 || (_data$choices = _data$choices[0]) === null || _data$choices === void 0 || (_data$choices = _data$choices.message) === null || _data$choices === void 0 ? void 0 : _data$choices.content) || "(Réponse vide)";
              output.textContent = answer;
              _context.n = 7;
              break;
            case 6:
              _context.p = 6;
              _t = _context.v;
              output.textContent = "Erreur: ".concat((_t === null || _t === void 0 ? void 0 : _t.message) || _t);
            case 7:
              return _context.a(2);
          }
        }, _callee, null, [[1, 6]]);
      })));

      // Insérer le contenu du mail en cours dans le prompt
      if (addMailBtn) {
        addMailBtn.addEventListener("click", /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
          var item, subject, bodyText, separator, _t2;
          return _regenerator().w(function (_context2) {
            while (1) switch (_context2.p = _context2.n) {
              case 0:
                _context2.p = 0;
                item = Office.context.mailbox.item;
                subject = (item === null || item === void 0 ? void 0 : item.subject) || "";
                bodyText = ""; // Récupère le corps au format texte
                _context2.n = 1;
                return new Promise(function (resolve, reject) {
                  item.body.getAsync(Office.CoercionType.Text, function (res) {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                      bodyText = res.value || "";
                      resolve();
                    } else {
                      reject(res.error);
                    }
                  });
                });
              case 1:
                separator = promptInput.value ? "\n\n" : "";
                promptInput.value = "".concat(promptInput.value).concat(separator, "Sujet: ").concat(subject, "\n\nCorps:\n").concat(bodyText).trim();
                _context2.n = 3;
                break;
              case 2:
                _context2.p = 2;
                _t2 = _context2.v;
                output.textContent = "Erreur lors de l'extraction du mail: ".concat((_t2 === null || _t2 === void 0 ? void 0 : _t2.message) || _t2);
              case 3:
                return _context2.a(2);
            }
          }, _callee2, null, [[0, 2]]);
        })));
      }

      // Ajouter les catégories du mail (Outlook) au prompt
      if (addCategoriesBtn) {
        addCategoriesBtn.addEventListener("click", /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
          var item, categories, userCats, _Office$context$mailb, setFromItem, all, toAppend, _t3, _t4;
          return _regenerator().w(function (_context3) {
            while (1) switch (_context3.p = _context3.n) {
              case 0:
                _context3.p = 0;
                item = Office.context.mailbox.item; // Récupère les catégories liées à l'élément (lecture/composition)
                categories = (item === null || item === void 0 ? void 0 : item.categories) || []; // Si vide côté item, tenter de récupérer les catégories définies par l'utilisateur (API mailbox.masterCategories)
                userCats = [];
                _context3.p = 1;
                if (!((_Office$context$mailb = Office.context.mailbox) !== null && _Office$context$mailb !== void 0 && (_Office$context$mailb = _Office$context$mailb.masterCategories) !== null && _Office$context$mailb !== void 0 && _Office$context$mailb.getAsync)) {
                  _context3.n = 2;
                  break;
                }
                _context3.n = 2;
                return new Promise(function (resolve) {
                  Office.context.mailbox.masterCategories.getAsync(function (res) {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                      userCats = (res.value || []).map(function (c) {
                        return c.displayName;
                      }).filter(Boolean);
                    }
                    resolve();
                  });
                });
              case 2:
                _context3.n = 4;
                break;
              case 3:
                _context3.p = 3;
                _t3 = _context3.v;
              case 4:
                setFromItem = Array.isArray(categories) ? categories : [];
                all = _toConsumableArray(new Set([].concat(_toConsumableArray(setFromItem || []), _toConsumableArray(userCats || []))));
                if (!(all.length === 0)) {
                  _context3.n = 5;
                  break;
                }
                output.textContent = "Aucune catégorie trouvée.";
                return _context3.a(2);
              case 5:
                toAppend = "\n\nCat\xE9gories Outlook: ".concat(all.join(", "));
                promptInput.value = "".concat(promptInput.value || "").concat(toAppend).trim();
                _context3.n = 7;
                break;
              case 6:
                _context3.p = 6;
                _t4 = _context3.v;
                output.textContent = "Erreur lors de la r\xE9cup\xE9ration des cat\xE9gories: ".concat((_t4 === null || _t4 === void 0 ? void 0 : _t4.message) || _t4);
              case 7:
                return _context3.a(2);
            }
          }, _callee3, null, [[1, 3], [0, 6]]);
        })));
      }
    }
  }
});
function run() {
  return _run.apply(this, arguments);
}
function _run() {
  _run = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
    var item, insertAt, label;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          /**
           * Insert your Outlook code here
           */
          item = Office.context.mailbox.item;
          insertAt = document.getElementById("item-subject");
          label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
          insertAt.appendChild(label);
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createTextNode(item.subject));
          insertAt.appendChild(document.createElement("br"));
        case 1:
          return _context4.a(2);
      }
    }, _callee4);
  }));
  return _run.apply(this, arguments);
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
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../../assets/logo-filled.png */ "./assets/logo-filled.png"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Contoso Task Pane Add-in</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\r\n        <img width=\"90\" height=\"90\" src=\"" + ___HTML_LOADER_IMPORT_1___ + "\" alt=\"Contoso\" title=\"Contoso\" />\r\n        <h1 class=\"ms-font-su\">Page d'accueil - Hello world</h1>\r\n    </header>\r\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\r\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\r\n    </section>\r\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\r\n        <h2 class=\"ms-font-xl\">Bienvenue</h2>\r\n        <p id=\"hello-text\" class=\"ms-font-l\">Hello world</p>\r\n\r\n        <section class=\"card\">\r\n            <h3 class=\"ms-font-l\">Tester l'API OpenAI (ChatGPT)</h3>\r\n            <div class=\"stack\">\r\n                <label class=\"ms-font-m\" for=\"api-key\">Clé API OpenAI (stockée localement, usage dev uniquement):</label>\r\n                <input id=\"api-key\" type=\"password\" placeholder=\"sk-...\" class=\"input\"/>\r\n\r\n                <label class=\"ms-font-m\" for=\"prompt-input\">Prompt</label>\r\n                <textarea id=\"prompt-input\" rows=\"5\" placeholder=\"Votre question...\" class=\"input\"></textarea>\r\n                <div style=\"display:flex; gap:8px; flex-wrap:wrap;\">\r\n                    <button id=\"send-prompt\" class=\"btn-primary\">Envoyer à ChatGPT</button>\r\n                    <button id=\"add-mail-content\" class=\"btn-secondary\">Ajouter le contenu du mail au prompt</button>\r\n                    <button id=\"add-mail-categories\" class=\"btn-secondary\">Ajouter les catégories du mail</button>\r\n                </div>\r\n            </div>\r\n\r\n            <div class=\"section\">\r\n                <h4 class=\"ms-font-m\">Réponse</h4>\r\n                <pre id=\"response-output\" class=\"output\"></pre>\r\n            </div>\r\n        </section>\r\n\r\n        <section class=\"card\">\r\n            <h3 class=\"ms-font-l\">Actions Outlook</h3>\r\n            <button id=\"run\" class=\"btn-primary\">Afficher le sujet</button>\r\n            <p id=\"item-subject\" class=\"ms-font-m\" style=\"margin-top:8px;\"></p>\r\n        </section>\r\n    </main>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map