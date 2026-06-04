"use strict";
self["webpackHotUpdateb018cc9a_f2bc_4ba4_90c1_1ebfd12e0aca_0_0_1"]("desk-booking-web-part",{

/***/ 82223:
/*!********************************************************!*\
  !*** ./lib/webparts/deskBooking/DeskBookingWebPart.js ***!
  \********************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 85959);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react-dom */ 48398);
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react_dom__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-core-library */ 89676);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-property-pane */ 39877);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 56642);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__);
/* harmony import */ var DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! DeskBookingWebPartStrings */ 45295);
/* harmony import */ var DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var _components_DeskBooking__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./components/DeskBooking */ 37027);








var DeskBookingWebPart = /** @class */ (function (_super) {
    (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__extends)(DeskBookingWebPart, _super);
    function DeskBookingWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    DeskBookingWebPart.prototype.render = function () {
        var element = react__WEBPACK_IMPORTED_MODULE_0__.createElement(_components_DeskBooking__WEBPACK_IMPORTED_MODULE_6__["default"], {
            context: this.context,
            deskMasterListTitle: this.properties.deskMasterListTitle || 'Desk Master',
            deskBookingListTitle: this.properties.deskBookingListTitle || 'Desk Booking',
            adminEmails: this.properties.adminEmails || ''
        });
        react_dom__WEBPACK_IMPORTED_MODULE_1__.render(element, this.domElement);
    };
    DeskBookingWebPart.prototype.onDispose = function () {
        react_dom__WEBPACK_IMPORTED_MODULE_1__.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(DeskBookingWebPart.prototype, "dataVersion", {
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_2__.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    DeskBookingWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.BasicGroupName,
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__.PropertyPaneTextField)('deskMasterListTitle', {
                                    label: DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.DeskMasterListTitleLabel
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__.PropertyPaneTextField)('deskBookingListTitle', {
                                    label: DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.DeskBookingListTitleLabel
                                }),
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_3__.PropertyPaneTextField)('adminEmails', {
                                    label: DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AdminEmailsLabel,
                                    description: DeskBookingWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AdminEmailsDescription
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return DeskBookingWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_4__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (DeskBookingWebPart);


/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ (() => {
/******/ 	__webpack_require__.h = () => ("e56620969c7bb7277324")
/******/ })();
/******/ 
/******/ }
);
//# sourceMappingURL=desk-booking-web-part.ec734790f830340450c7.hot-update.js.map