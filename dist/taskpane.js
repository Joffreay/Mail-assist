/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "81e289566de3ba43da90.css";

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
    document.getElementById("run").onclick = run;
    var apiKeyInput = document.getElementById("api-key");
    var promptInput = document.getElementById("prompt-input");
    var sendBtn = document.getElementById("send-prompt");
    var addMailBtn = document.getElementById("add-mail-content");
    var addCategoriesBtn = document.getElementById("add-mail-categories");
    var output = document.getElementById("response-output");
    var autoReplyBtn = document.getElementById("auto-reply");
    var listEventsBtn = document.getElementById("list-events");
    var msalClientIdInput = document.getElementById("msal-client-id");
    var eventsOutput = document.getElementById("events-output");
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

      // Générer une réponse et ouvrir une réponse préremplie
      if (autoReplyBtn) {
        autoReplyBtn.addEventListener("click", /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
          var _data$choices2, apiKey, item, subject, bodyText, prompt, resp, errText, data, draft, _t2;
          return _regenerator().w(function (_context2) {
            while (1) switch (_context2.p = _context2.n) {
              case 0:
                _context2.p = 0;
                apiKey = apiKeyInput.value.trim();
                if (apiKey) {
                  _context2.n = 1;
                  break;
                }
                output.textContent = "Veuillez saisir la clé API avant de générer la réponse.";
                return _context2.a(2);
              case 1:
                item = Office.context.mailbox.item;
                subject = (item === null || item === void 0 ? void 0 : item.subject) || "";
                bodyText = ""; // Récupère le corps au format texte
                _context2.n = 2;
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
              case 2:
                output.textContent = "Génération de la réponse...";
                prompt = "Tu es un assistant qui r\xE9dige des r\xE9ponses d'email courtes, polies et en fran\xE7ais.\n\nSujet: ".concat(subject, "\n\nContenu:\n").concat(bodyText, "\n\nR\xE9dige une r\xE9ponse appropri\xE9e en gardant un ton professionnel, commence par un salut et termine par une formule de politesse.");
                _context2.n = 3;
                return fetch("https://api.openai.com/v1/chat/completions", {
                  method: "POST",
                  headers: {
                    "Content-Type": "application/json",
                    Authorization: "Bearer ".concat(apiKey)
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
                    temperature: 0.6
                  })
                });
              case 3:
                resp = _context2.v;
                if (resp.ok) {
                  _context2.n = 5;
                  break;
                }
                _context2.n = 4;
                return resp.text();
              case 4:
                errText = _context2.v;
                output.textContent = "Erreur API (".concat(resp.status, "): ").concat(errText);
                return _context2.a(2);
              case 5:
                _context2.n = 6;
                return resp.json();
              case 6:
                data = _context2.v;
                draft = (data === null || data === void 0 || (_data$choices2 = data.choices) === null || _data$choices2 === void 0 || (_data$choices2 = _data$choices2[0]) === null || _data$choices2 === void 0 || (_data$choices2 = _data$choices2.message) === null || _data$choices2 === void 0 ? void 0 : _data$choices2.content) || ""; // Ouvrir une réponse pré-remplie
                _context2.n = 7;
                return new Promise(function (resolve, reject) {
                  item.displayReplyForm(draft, function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                      resolve();
                    } else {
                      reject(asyncResult.error);
                    }
                  });
                });
              case 7:
                output.textContent = "Réponse générée et placée dans un brouillon de réponse.";
                _context2.n = 9;
                break;
              case 8:
                _context2.p = 8;
                _t2 = _context2.v;
                output.textContent = "Erreur lors de la r\xE9ponse automatique: ".concat((_t2 === null || _t2 === void 0 ? void 0 : _t2.message) || _t2);
              case 9:
                return _context2.a(2);
            }
          }, _callee2, null, [[0, 8]]);
        })));
      }

      // Insérer le contenu du mail en cours dans le prompt
      if (addMailBtn) {
        addMailBtn.addEventListener("click", /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
          var item, subject, bodyText, separator, _t3;
          return _regenerator().w(function (_context3) {
            while (1) switch (_context3.p = _context3.n) {
              case 0:
                _context3.p = 0;
                item = Office.context.mailbox.item;
                subject = (item === null || item === void 0 ? void 0 : item.subject) || "";
                bodyText = ""; // Récupère le corps au format texte
                _context3.n = 1;
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
                _context3.n = 3;
                break;
              case 2:
                _context3.p = 2;
                _t3 = _context3.v;
                output.textContent = "Erreur lors de l'extraction du mail: ".concat((_t3 === null || _t3 === void 0 ? void 0 : _t3.message) || _t3);
              case 3:
                return _context3.a(2);
            }
          }, _callee3, null, [[0, 2]]);
        })));
      }

      // Ajouter les catégories du mail (Outlook) au prompt
      if (addCategoriesBtn) {
        addCategoriesBtn.addEventListener("click", /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
          var item, categories, userCats, _Office$context$mailb, req, hasApi, setFromItem, all, toAppend, _t4, _t5;
          return _regenerator().w(function (_context4) {
            while (1) switch (_context4.p = _context4.n) {
              case 0:
                _context4.p = 0;
                item = Office.context.mailbox.item; // Récupère les catégories liées à l'élément (lecture/composition)
                categories = (item === null || item === void 0 ? void 0 : item.categories) || []; // Si vide côté item, tenter de récupérer les catégories définies par l'utilisateur (API mailbox.masterCategories)
                userCats = [];
                _context4.p = 1;
                req = Office.context.requirements;
                hasApi = req && req.isSetSupported && req.isSetSupported("Mailbox", "1.8");
                if (!(hasApi && (_Office$context$mailb = Office.context.mailbox) !== null && _Office$context$mailb !== void 0 && (_Office$context$mailb = _Office$context$mailb.masterCategories) !== null && _Office$context$mailb !== void 0 && _Office$context$mailb.getAsync)) {
                  _context4.n = 2;
                  break;
                }
                _context4.n = 2;
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
                _context4.n = 4;
                break;
              case 3:
                _context4.p = 3;
                _t4 = _context4.v;
              case 4:
                setFromItem = Array.isArray(categories) ? categories : [];
                all = _toConsumableArray(new Set([].concat(_toConsumableArray(setFromItem || []), _toConsumableArray(userCats || []))));
                if (!(all.length === 0)) {
                  _context4.n = 5;
                  break;
                }
                output.textContent = "Aucune catégorie trouvée (l’API Catégories nécessite Outlook supportant Mailbox 1.8).";
                return _context4.a(2);
              case 5:
                toAppend = "\n\nCat\xE9gories Outlook: ".concat(all.join(", "));
                promptInput.value = "".concat(promptInput.value || "").concat(toAppend).trim();
                _context4.n = 7;
                break;
              case 6:
                _context4.p = 6;
                _t5 = _context4.v;
                output.textContent = "Erreur lors de la r\xE9cup\xE9ration des cat\xE9gories: ".concat((_t5 === null || _t5 === void 0 ? void 0 : _t5.message) || _t5);
              case 7:
                return _context4.a(2);
            }
          }, _callee4, null, [[1, 3], [0, 6]]);
        })));
      }
    }
  }
});
function run() {
  return _run.apply(this, arguments);
}

// Ouvre un Dialog pour authentifier l'utilisateur et renvoyer un token MS Graph
function _run() {
  _run = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6() {
    var item, insertAt, label;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.n) {
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
          return _context6.a(2);
      }
    }, _callee6);
  }));
  return _run.apply(this, arguments);
}
function openAuthDialog(clientId, onResult) {
  var url = "".concat(location.origin, "/auth2.html?clientId=").concat(encodeURIComponent(clientId));
  Office.context.ui.displayDialogAsync(url, {
    height: 60,
    width: 40,
    displayInIframe: false
  }, function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      onResult(null, new Error("Impossible d’ouvrir la fenêtre d’authentification."));
      return;
    }
    var dialog = res.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
      try {
        var msg = JSON.parse(arg.message || "{}");
        if (msg.type === "token" && msg.token) {
          dialog.close();
          onResult(msg.token, null);
        } else if (msg.type === "error") {
          dialog.close();
          onResult(null, new Error(msg.error || "Erreur d’authentification."));
        }
      } catch (e) {
        dialog.close();
        onResult(null, e);
      }
    });
  });
}

// Attache l’action du bouton "Lister mes prochains RDV"
if (typeof document !== "undefined") {
  document.addEventListener("DOMContentLoaded", function () {
    // Bouton catégories utilisateur (masterCategories)
    var catBtn = document.getElementById("list-categories");
    var catOut = document.getElementById("categories-output");
    if (catBtn && catOut) {
      catBtn.addEventListener("click", function () {
        try {
          var _Office$context$mailb2;
          var req = Office.context.requirements;
          var hasApi = req && req.isSetSupported && req.isSetSupported("Mailbox", "1.8");
          if (!(hasApi && (_Office$context$mailb2 = Office.context.mailbox) !== null && _Office$context$mailb2 !== void 0 && (_Office$context$mailb2 = _Office$context$mailb2.masterCategories) !== null && _Office$context$mailb2 !== void 0 && _Office$context$mailb2.getAsync)) {
            catOut.textContent = "API non supportée (Mailbox 1.8 requis).";
            return;
          }
          catOut.textContent = "Chargement des catégories...";
          Office.context.mailbox.masterCategories.getAsync(function (res) {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              var names = (res.value || []).map(function (c) {
                return c.displayName;
              }).filter(Boolean);
              catOut.textContent = names.length ? names.join("\n") : "Aucune catégorie définie.";
            } else {
              var _res$error;
              catOut.textContent = "Erreur: ".concat(((_res$error = res.error) === null || _res$error === void 0 ? void 0 : _res$error.message) || res.error || "inconnue");
            }
          });
        } catch (e) {
          catOut.textContent = "Erreur: ".concat((e === null || e === void 0 ? void 0 : e.message) || e);
        }
      });
    }
    var btn = document.getElementById("list-events");
    var clientIdInput = document.getElementById("msal-client-id");
    var out = document.getElementById("events-output");
    // Préremplir depuis localStorage si disponible
    try {
      var savedClientId = localStorage.getItem("msal_client_id");
      if (savedClientId && clientIdInput && !clientIdInput.value) {
        clientIdInput.value = savedClientId;
      }
    } catch (_unused4) {}
    if (btn && clientIdInput && out) {
      btn.addEventListener("click", function () {
        var clientId = clientIdInput.value.trim();
        if (!clientId) {
          out.textContent = "Veuillez saisir un Client ID (Azure AD) pour appeler Microsoft Graph.";
          return;
        }
        // Mémoriser le Client ID pour l’auth dialog
        try {
          localStorage.setItem("msal_client_id", clientId);
        } catch (_unused5) {}
        out.textContent = "Authentification...";
        openAuthDialog(clientId, /*#__PURE__*/function () {
          var _ref5 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(accessToken, err) {
            var now, end, startIso, endIso, url, resp, t, data, items, lines, _t6;
            return _regenerator().w(function (_context5) {
              while (1) switch (_context5.p = _context5.n) {
                case 0:
                  if (!err) {
                    _context5.n = 1;
                    break;
                  }
                  out.textContent = "Erreur: ".concat(err.message || err);
                  return _context5.a(2);
                case 1:
                  _context5.p = 1;
                  out.textContent = "Chargement des rendez-vous...";
                  now = new Date();
                  end = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
                  startIso = now.toISOString();
                  endIso = end.toISOString();
                  url = "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=".concat(encodeURIComponent(startIso), "&endDateTime=").concat(encodeURIComponent(endIso), "&$orderby=start/dateTime&$top=10");
                  _context5.n = 2;
                  return fetch(url, {
                    headers: {
                      Authorization: "Bearer ".concat(accessToken)
                    }
                  });
                case 2:
                  resp = _context5.v;
                  if (resp.ok) {
                    _context5.n = 4;
                    break;
                  }
                  _context5.n = 3;
                  return resp.text();
                case 3:
                  t = _context5.v;
                  out.textContent = "Erreur Graph (".concat(resp.status, "): ").concat(t);
                  return _context5.a(2);
                case 4:
                  _context5.n = 5;
                  return resp.json();
                case 5:
                  data = _context5.v;
                  items = (data === null || data === void 0 ? void 0 : data.value) || [];
                  if (!(items.length === 0)) {
                    _context5.n = 6;
                    break;
                  }
                  out.textContent = "Aucun rendez-vous à venir dans les 7 prochains jours.";
                  return _context5.a(2);
                case 6:
                  lines = items.map(function (it) {
                    var _it$start, _it$end, _it$location;
                    var subject = it.subject || "(Sans objet)";
                    var start = ((_it$start = it.start) === null || _it$start === void 0 ? void 0 : _it$start.dateTime) || "";
                    var end = ((_it$end = it.end) === null || _it$end === void 0 ? void 0 : _it$end.dateTime) || "";
                    var location = (_it$location = it.location) !== null && _it$location !== void 0 && _it$location.displayName ? " @ ".concat(it.location.displayName) : "";
                    return "- ".concat(subject, " (").concat(start, " -> ").concat(end, ")").concat(location);
                  });
                  out.textContent = lines.join("\n");
                  _context5.n = 8;
                  break;
                case 7:
                  _context5.p = 7;
                  _t6 = _context5.v;
                  out.textContent = "Erreur lors de la r\xE9cup\xE9ration des \xE9v\xE9nements: ".concat((_t6 === null || _t6 === void 0 ? void 0 : _t6.message) || _t6);
                case 8:
                  return _context5.a(2);
              }
            }, _callee5, null, [[1, 7]]);
          }));
          return function (_x, _x2) {
            return _ref5.apply(this, arguments);
          };
        }());
      });
    }
  });
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
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Contoso Task Pane Add-in</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\r\n        <h1 class=\"ms-font-su\" style=\"margin:0;\">Bienvenue</h1>\r\n    </header>\r\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\r\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\r\n    </section>\r\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\r\n\r\n        <section class=\"card\">\r\n            <h3 class=\"ms-font-l\">Tester l'API OpenAI (ChatGPT)</h3>\r\n            <div class=\"stack\">\r\n                <label class=\"ms-font-m\" for=\"api-key\">Clé API OpenAI (stockée localement, usage dev uniquement):</label>\r\n                <input id=\"api-key\" type=\"password\" placeholder=\"sk-...\" class=\"input\"/>\r\n\r\n                <label class=\"ms-font-m\" for=\"prompt-input\">Prompt</label>\r\n                <textarea id=\"prompt-input\" rows=\"5\" placeholder=\"Votre question...\" class=\"input\"></textarea>\r\n                <div style=\"display:flex; gap:8px; flex-wrap:wrap;\">\r\n                    <button id=\"send-prompt\" class=\"btn-primary\">Envoyer à ChatGPT</button>\r\n                    <button id=\"add-mail-content\" class=\"btn-secondary\">Ajouter le contenu du mail au prompt</button>\r\n                    <button id=\"add-mail-categories\" class=\"btn-secondary\">Ajouter les catégories du mail</button>\r\n                </div>\r\n            </div>\r\n\r\n            <div class=\"section\">\r\n                <h4 class=\"ms-font-m\">Réponse</h4>\r\n                <pre id=\"response-output\" class=\"output\"></pre>\r\n            </div>\r\n        </section>\r\n\r\n        <section class=\"card\">\r\n            <h3 class=\"ms-font-l\">Actions Outlook</h3>\r\n            <div style=\"display:flex; gap:8px; flex-wrap:wrap;\">\r\n                <button id=\"run\" class=\"btn-primary\">Afficher le sujet</button>\r\n                <button id=\"auto-reply\" class=\"btn-secondary\">Réponse automatique</button>\r\n            </div>\r\n            <p id=\"item-subject\" class=\"ms-font-m\" style=\"margin-top:8px;\"></p>\r\n        </section>\r\n\r\n        <section class=\"card\">\r\n            <h3 class=\"ms-font-l\">Mes prochains rendez-vous</h3>\r\n            <div class=\"stack\">\r\n                <div class=\"stack\" style=\"gap:6px;\">\r\n                    <label class=\"ms-font-m\" for=\"msal-client-id\">Client ID (Azure AD) — requis pour Microsoft Graph</label>\r\n                    <input id=\"msal-client-id\" class=\"input\" placeholder=\"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\" />\r\n                </div>\r\n                <button id=\"list-events\" class=\"btn-primary\">Lister mes prochains RDV</button>\r\n                <div id=\"events-output\" class=\"output\" style=\"min-height: 40px;\">Aucune donnée.</div>\r\n            </div>\r\n        </section>\r\n\r\n        <section class=\"card\">\r\n            <h3 class=\"ms-font-l\">Catégories de l’utilisateur</h3>\r\n            <div class=\"stack\">\r\n                <button id=\"list-categories\" class=\"btn-primary\">Récupérer les catégories</button>\r\n                <div id=\"categories-output\" class=\"output\" style=\"min-height: 40px;\">Aucune donnée.</div>\r\n            </div>\r\n        </section>\r\n    </main>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map