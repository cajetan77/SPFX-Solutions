"use strict";
self["webpackHotUpdateb018cc9a_f2bc_4ba4_90c1_1ebfd12e0aca_0_0_1"]("desk-booking-web-part",{

/***/ 37027:
/*!************************************************************!*\
  !*** ./lib/webparts/deskBooking/components/DeskBooking.js ***!
  \************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 85959);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! @fluentui/react */ 21314);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! @fluentui/react */ 72674);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! @fluentui/react */ 63208);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! @fluentui/react */ 46643);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! @fluentui/react */ 51330);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! @fluentui/react */ 67102);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! @fluentui/react */ 12042);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! @fluentui/react */ 5613);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! @fluentui/react */ 23166);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! @fluentui/react */ 80954);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! @fluentui/react */ 49885);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! @fluentui/react */ 29425);
/* harmony import */ var _pnp_spfx_controls_react_lib_PeoplePicker__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/spfx-controls-react/lib/PeoplePicker */ 75216);
/* harmony import */ var _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./DeskBooking.module.scss */ 26861);
/* harmony import */ var _services_DeskBookingService__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../services/DeskBookingService */ 55236);
/* harmony import */ var _models_IModels__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../models/IModels */ 36556);
/* harmony import */ var _utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/dateTimeUtils */ 86418);
/* harmony import */ var _utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/bookingRules */ 62488);










var DeskBooking = function (props) {
    var context = props.context, deskMasterListTitle = props.deskMasterListTitle, deskBookingListTitle = props.deskBookingListTitle, adminEmails = props.adminEmails;
    var service = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return new _services_DeskBookingService__WEBPACK_IMPORTED_MODULE_3__.DeskBookingService(context, deskMasterListTitle, deskBookingListTitle); }, [context, deskMasterListTitle, deskBookingListTitle]);
    var peoplePickerContext = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return ({
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    }); }, [context]);
    var timeSlots = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.generateTimeSlots)(); }, []);
    var timeOptions = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return timeSlots.map(function (slot) { return ({ key: slot, text: slot }); }); }, [timeSlots]);
    var _a = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(''), bookerName = _a[0], setBookerName = _a[1];
    var _b = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(''), bookerEmail = _b[0], setBookerEmail = _b[1];
    var _c = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)('internal'), bookingUserType = _c[0], setBookingUserType = _c[1];
    var bookingDate = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.getTodayDate)(); }, []);
    var isBookingDay = (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.isWeekday)(bookingDate);
    var _d = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)('09:00'), startTime = _d[0], setStartTime = _d[1];
    var _e = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)('17:00'), endTime = _e[0], setEndTime = _e[1];
    var _f = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)([]), desks = _f[0], setDesks = _f[1];
    var _g = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)([]), yourBookings = _g[0], setYourBookings = _g[1];
    var _h = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(false), bookingsLoaded = _h[0], setBookingsLoaded = _h[1];
    var _j = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(true), loading = _j[0], setLoading = _j[1];
    var _k = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(false), loadingBookings = _k[0], setLoadingBookings = _k[1];
    var _l = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(false), actionInProgress = _l[0], setActionInProgress = _l[1];
    var _m = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(undefined), errorMessage = _m[0], setErrorMessage = _m[1];
    var _o = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(undefined), successMessage = _o[0], setSuccessMessage = _o[1];
    var _p = (0,react__WEBPACK_IMPORTED_MODULE_0__.useState)(undefined), internalPersonId = _p[0], setInternalPersonId = _p[1];
    var currentUserLoginName = context.pageContext.user.loginName || '';
    var currentUserEmail = context.pageContext.user.email || '';
    var bookingAdminEmails = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return (0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.parseAdminEmails)(adminEmails); }, [adminEmails]);
    var isCurrentUserAdmin = (0,react__WEBPACK_IMPORTED_MODULE_0__.useMemo)(function () { return (0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.isBookingAdmin)(currentUserEmail, bookingAdminEmails); }, [bookingAdminEmails, currentUserEmail]);
    var handleInternalUserChange = (0,react__WEBPACK_IMPORTED_MODULE_0__.useCallback)(function (users) { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
        var user, numericId, loginName, id, error_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!(users === null || users === void 0 ? void 0 : users.length)) {
                        setInternalPersonId(undefined);
                        setBookerName('');
                        setBookerEmail('');
                        setBookingsLoaded(false);
                        return [2 /*return*/];
                    }
                    user = users[0];
                    setBookerName(user.text || '');
                    setBookerEmail(user.secondaryText || '');
                    setBookingsLoaded(false);
                    setErrorMessage(undefined);
                    numericId = Number(user.id);
                    if (!isNaN(numericId) && numericId > 0) {
                        setInternalPersonId(numericId);
                        return [2 /*return*/];
                    }
                    loginName = user.loginName || user.id;
                    if (!loginName) {
                        setInternalPersonId(undefined);
                        return [2 /*return*/];
                    }
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, service.resolveUserId(loginName)];
                case 2:
                    id = _a.sent();
                    setInternalPersonId(id);
                    return [3 /*break*/, 4];
                case 3:
                    error_1 = _a.sent();
                    setInternalPersonId(undefined);
                    setErrorMessage(getErrorMessage(error_1));
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); }, [service]);
    (0,react__WEBPACK_IMPORTED_MODULE_0__.useEffect)(function () {
        if (bookingUserType === 'external') {
            setInternalPersonId(undefined);
            setBookerName('');
            setBookerEmail('');
            setBookingsLoaded(false);
            return;
        }
        if (!currentUserLoginName) {
            return;
        }
        var initializeDefaultInternalUser = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
            var id, _a;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, service.resolveUserId(currentUserLoginName)];
                    case 1:
                        id = _b.sent();
                        setInternalPersonId(id);
                        setBookerName(context.pageContext.user.displayName || '');
                        setBookerEmail(currentUserEmail);
                        return [3 /*break*/, 3];
                    case 2:
                        _a = _b.sent();
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        initializeDefaultInternalUser().catch(function () {
            // User can select a person manually in the picker.
        });
    }, [bookingUserType, context.pageContext.user.displayName, currentUserEmail, currentUserLoginName, service]);
    var loadDesks = (0,react__WEBPACK_IMPORTED_MODULE_0__.useCallback)(function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
        var _a, activeDesks, bookingsForDate, activeBookings_1, desksWithStatus, error_2;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_b) {
            switch (_b.label) {
                case 0:
                    setLoading(true);
                    setErrorMessage(undefined);
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, Promise.all([
                            service.getActiveDesks(),
                            service.getBookingsForDate(bookingDate)
                        ])];
                case 2:
                    _a = _b.sent(), activeDesks = _a[0], bookingsForDate = _a[1];
                    activeBookings_1 = bookingsForDate.filter(function (booking) { return booking.bookingStatus === _models_IModels__WEBPACK_IMPORTED_MODULE_4__.BookingStatus.Booked; });
                    desksWithStatus = activeDesks.map(function (desk) {
                        var deskBookings = activeBookings_1.filter(function (booking) { return booking.deskId === desk.id; });
                        var overlappingBooking = deskBookings.find(function (booking) {
                            return (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.timesOverlap)(startTime, endTime, booking.startTime, booking.endTime);
                        });
                        var isBooked = !!overlappingBooking;
                        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)({}, desk), { displayStatus: isBooked ? _models_IModels__WEBPACK_IMPORTED_MODULE_4__.DeskDisplayStatus.Booked : _models_IModels__WEBPACK_IMPORTED_MODULE_4__.DeskDisplayStatus.Available, bookedByName: overlappingBooking === null || overlappingBooking === void 0 ? void 0 : overlappingBooking.bookerName, activeBookingId: overlappingBooking === null || overlappingBooking === void 0 ? void 0 : overlappingBooking.id });
                    });
                    setDesks(desksWithStatus);
                    return [3 /*break*/, 5];
                case 3:
                    error_2 = _b.sent();
                    setErrorMessage(getErrorMessage(error_2));
                    return [3 /*break*/, 5];
                case 4:
                    setLoading(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); }, [bookingDate, endTime, service, startTime]);
    var loadYourBookings = (0,react__WEBPACK_IMPORTED_MODULE_0__.useCallback)(function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
        var validation, todayBookings, bookings, error_3;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!isCurrentUserAdmin) {
                        validation = (0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.validateBookerDetails)({ name: bookerName, email: bookerEmail });
                        if (!validation.isValid) {
                            setErrorMessage(validation.message);
                            return [2 /*return*/];
                        }
                    }
                    setLoadingBookings(true);
                    setErrorMessage(undefined);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    if (!isCurrentUserAdmin) return [3 /*break*/, 3];
                    return [4 /*yield*/, service.getBookingsForDate(bookingDate)];
                case 2:
                    todayBookings = _a.sent();
                    setYourBookings(todayBookings.filter(function (booking) { return booking.bookingStatus === _models_IModels__WEBPACK_IMPORTED_MODULE_4__.BookingStatus.Booked; }));
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, service.getBookingsByEmail(bookerEmail)];
                case 4:
                    bookings = _a.sent();
                    setYourBookings(bookings);
                    _a.label = 5;
                case 5:
                    setBookingsLoaded(true);
                    return [3 /*break*/, 8];
                case 6:
                    error_3 = _a.sent();
                    setErrorMessage(getErrorMessage(error_3));
                    return [3 /*break*/, 8];
                case 7:
                    setLoadingBookings(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); }, [bookerEmail, bookerName, bookingDate, isCurrentUserAdmin, service]);
    (0,react__WEBPACK_IMPORTED_MODULE_0__.useEffect)(function () {
        loadDesks().catch(function () {
            // Error state is handled in loadDesks.
        });
    }, [loadDesks]);
    var handleBookDesk = function (deskId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
        var request, validation, bookingsForDate, error_4;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!isBookingDay) {
                        setErrorMessage('Desk booking is only available on weekdays (Monday through Friday).');
                        return [2 /*return*/];
                    }
                    request = {
                        deskId: deskId,
                        bookingDate: bookingDate,
                        startTime: startTime,
                        endTime: endTime,
                        bookerName: bookerName,
                        bookerEmail: bookerEmail,
                        bookerPersonId: bookingUserType === 'internal' ? internalPersonId : undefined
                    };
                    validation = (0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.validateBookingRequest)(request, {
                        requireInternalPerson: bookingUserType === 'internal'
                    });
                    if (!validation.isValid) {
                        setErrorMessage(validation.message);
                        return [2 /*return*/];
                    }
                    setActionInProgress(true);
                    setErrorMessage(undefined);
                    setSuccessMessage(undefined);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 7, 8, 9]);
                    return [4 /*yield*/, service.getBookingsForDate(bookingDate)];
                case 2:
                    bookingsForDate = _a.sent();
                    if ((0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.hasConflictingBooking)(request, bookingsForDate)) {
                        setErrorMessage('This desk is already booked for the selected date and time.');
                        return [2 /*return*/];
                    }
                    return [4 /*yield*/, service.createBooking(request)];
                case 3:
                    _a.sent();
                    setSuccessMessage('Desk booked successfully.');
                    return [4 /*yield*/, loadDesks()];
                case 4:
                    _a.sent();
                    if (!bookingsLoaded) return [3 /*break*/, 6];
                    return [4 /*yield*/, loadYourBookings()];
                case 5:
                    _a.sent();
                    _a.label = 6;
                case 6: return [3 /*break*/, 9];
                case 7:
                    error_4 = _a.sent();
                    setErrorMessage(getErrorMessage(error_4));
                    return [3 /*break*/, 9];
                case 8:
                    setActionInProgress(false);
                    return [7 /*endfinally*/];
                case 9: return [2 /*return*/];
            }
        });
    }); };
    var handleCancelBooking = function (booking) { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
        var error_5;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!(0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.canCancelBooking)(booking, bookerEmail, currentUserEmail, isCurrentUserAdmin)) {
                        setErrorMessage('You do not have permission to cancel this booking.');
                        return [2 /*return*/];
                    }
                    setActionInProgress(true);
                    setErrorMessage(undefined);
                    setSuccessMessage(undefined);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 5, 6, 7]);
                    return [4 /*yield*/, service.cancelBooking(booking.id)];
                case 2:
                    _a.sent();
                    setSuccessMessage('Booking cancelled successfully.');
                    return [4 /*yield*/, loadDesks()];
                case 3:
                    _a.sent();
                    return [4 /*yield*/, loadYourBookings()];
                case 4:
                    _a.sent();
                    return [3 /*break*/, 7];
                case 5:
                    error_5 = _a.sent();
                    setErrorMessage(getErrorMessage(error_5));
                    return [3 /*break*/, 7];
                case 6:
                    setActionInProgress(false);
                    return [7 /*endfinally*/];
                case 7: return [2 /*return*/];
            }
        });
    }); };
    var renderStatusBadge = function (status) {
        var className = "".concat(_DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].statusBadge, " ").concat(_DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"]["status".concat(status)]);
        return react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { className: className }, status);
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("section", { className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].deskBooking },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { tokens: { childrenGap: 16 } },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { tokens: { childrenGap: 4 } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { variant: "xLarge", block: true }, "Desk Booking"),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { variant: "medium", block: true }, "Select an internal user (People picker) to book for yourself or others, or enter external guest details. Bookings are for today only, on weekdays between 7:00 AM and 6:00 PM.")),
            errorMessage && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_10__.MessageBar, { messageBarType: _fluentui_react__WEBPACK_IMPORTED_MODULE_11__.MessageBarType.error, onDismiss: function () { return setErrorMessage(undefined); } }, errorMessage)),
            successMessage && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_10__.MessageBar, { messageBarType: _fluentui_react__WEBPACK_IMPORTED_MODULE_11__.MessageBarType.success, onDismiss: function () { return setSuccessMessage(undefined); } }, successMessage)),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_12__.ChoiceGroup, { label: "Booking type", selectedKey: bookingUserType, options: [
                    { key: 'internal', text: 'Internal user (People picker)' },
                    { key: 'external', text: 'External user (manual name and email)' }
                ], onChange: function (_, option) {
                    var nextType = (option === null || option === void 0 ? void 0 : option.key) || 'internal';
                    setBookingUserType(nextType);
                    setErrorMessage(undefined);
                    setSuccessMessage(undefined);
                    if (nextType === 'internal') {
                        setBookerName('');
                        setBookerEmail('');
                        setInternalPersonId(undefined);
                        setBookingsLoaded(false);
                    }
                } }),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { horizontal: true, wrap: true, tokens: { childrenGap: 16 }, verticalAlign: "end", className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].filters }, bookingUserType === 'internal' ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 320, maxWidth: 480 } } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_pnp_spfx_controls_react_lib_PeoplePicker__WEBPACK_IMPORTED_MODULE_1__.PeoplePicker, { key: "internal-people-picker", context: peoplePickerContext, ensureUser: true, titleText: "Internal user", personSelectionLimit: 1, showtooltip: true, required: true, disabled: actionInProgress, principalTypes: [_pnp_spfx_controls_react_lib_PeoplePicker__WEBPACK_IMPORTED_MODULE_1__.PrincipalType.User], resolveDelay: 300, defaultSelectedUsers: currentUserLoginName ? [currentUserLoginName] : [], onChange: function (users) {
                        handleInternalUserChange(users).catch(function () {
                            // Error state is handled in handleInternalUserChange.
                        });
                    } }))) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(react__WEBPACK_IMPORTED_MODULE_0__.Fragment, null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 220 } } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_13__.TextField, { label: "External name", value: bookerName, onChange: function (_, value) { return setBookerName(value || ''); }, required: true, placeholder: "Enter external user's full name" })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 220 } } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_13__.TextField, { label: "External email", value: bookerEmail, onChange: function (_, value) {
                            setBookerEmail(value || '');
                            setBookingsLoaded(false);
                        }, required: true, placeholder: "Enter external user's email address" }))))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { horizontal: true, wrap: true, tokens: { childrenGap: 16 }, verticalAlign: "end", className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].filters },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 220 } } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_13__.TextField, { label: "Booking date", value: (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.formatDisplayDate)(bookingDate), readOnly: true })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 180 } } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_14__.Dropdown, { label: "Start time", selectedKey: startTime, options: timeOptions, onChange: function (_, option) { return option && setStartTime(option.key); }, required: true })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 180 } } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_14__.Dropdown, { label: "End time", selectedKey: endTime, options: timeOptions, onChange: function (_, option) { return option && setEndTime(option.key); }, required: true })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, null,
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.DefaultButton, { text: "Refresh", onClick: function () { return loadDesks(); }, disabled: loading || actionInProgress }))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { tokens: { childrenGap: 12 } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_16__.Label, null,
                    "Available desks for ",
                    (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.formatDisplayDate)(bookingDate)),
                !isBookingDay ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_10__.MessageBar, { messageBarType: _fluentui_react__WEBPACK_IMPORTED_MODULE_11__.MessageBarType.warning }, "Desk booking is only available on weekdays. Today is not a booking day.")) : loading ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_17__.Spinner, { size: _fluentui_react__WEBPACK_IMPORTED_MODULE_18__.SpinnerSize.large, label: "Loading desks..." })) : desks.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_10__.MessageBar, { messageBarType: _fluentui_react__WEBPACK_IMPORTED_MODULE_11__.MessageBarType.info },
                    "No active desks found. Add desks to the \"",
                    deskMasterListTitle,
                    "\" list or check your list configuration.")) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].deskGrid }, desks.map(function (desk) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { key: desk.id, className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].deskCard },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { tokens: { childrenGap: 8 } },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { variant: "large", block: true }, desk.deskName),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { variant: "small", block: true }, desk.location || 'No location specified'),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { horizontal: true, verticalAlign: "center", tokens: { childrenGap: 8 } }, renderStatusBadge(desk.displayStatus)),
                        desk.displayStatus === _models_IModels__WEBPACK_IMPORTED_MODULE_4__.DeskDisplayStatus.Booked && desk.bookedByName && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { variant: "small", block: true },
                            "Booked by ",
                            desk.bookedByName)),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { tokens: { childrenGap: 8 } },
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.PrimaryButton, { text: "Book", disabled: !isBookingDay
                                    || desk.displayStatus !== _models_IModels__WEBPACK_IMPORTED_MODULE_4__.DeskDisplayStatus.Available
                                    || actionInProgress, onClick: function () { return handleBookDesk(desk.id); } }),
                            isCurrentUserAdmin && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.DefaultButton, { text: "Cancel", disabled: desk.displayStatus !== _models_IModels__WEBPACK_IMPORTED_MODULE_4__.DeskDisplayStatus.Booked
                                    || !desk.activeBookingId
                                    || actionInProgress, onClick: function () { return handleCancelBooking({
                                    id: desk.activeBookingId,
                                    deskId: desk.id,
                                    deskName: desk.deskName,
                                    bookingDate: bookingDate,
                                    startTime: startTime,
                                    endTime: endTime,
                                    bookerName: desk.bookedByName || '',
                                    bookerEmail: '',
                                    createdByEmail: '',
                                    bookingStatus: _models_IModels__WEBPACK_IMPORTED_MODULE_4__.BookingStatus.Booked
                                }); } })))))); })))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { tokens: { childrenGap: 12 } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { horizontal: true, wrap: true, verticalAlign: "center", tokens: { childrenGap: 12 } },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_16__.Label, null, isCurrentUserAdmin ? 'Today\'s bookings' : 'Your bookings'),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.DefaultButton, { text: isCurrentUserAdmin ? 'Load today\'s bookings' : 'Find my bookings', onClick: function () { return loadYourBookings(); }, disabled: loadingBookings || actionInProgress })),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { variant: "small", block: true }, isCurrentUserAdmin
                    ? 'As a booking admin, you can view and cancel any active booking for today.'
                    : 'Use the selected booking type details above, then click Find my bookings to view or cancel reservations.'),
                loadingBookings ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_17__.Spinner, { size: _fluentui_react__WEBPACK_IMPORTED_MODULE_18__.SpinnerSize.medium, label: "Loading your bookings..." })) : bookingsLoaded && yourBookings.length === 0 ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_10__.MessageBar, { messageBarType: _fluentui_react__WEBPACK_IMPORTED_MODULE_11__.MessageBarType.info }, "No bookings found for this email address.")) : bookingsLoaded ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].bookingList }, yourBookings.map(function (booking) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { key: booking.id, className: _DeskBooking_module_scss__WEBPACK_IMPORTED_MODULE_2__["default"].bookingRow },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack, { horizontal: true, wrap: true, verticalAlign: "center", tokens: { childrenGap: 12 } },
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, { grow: 1, styles: { root: { minWidth: 240 } } },
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { block: true, variant: "mediumPlus" }, booking.deskName),
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { block: true, variant: "small" },
                                (0,_utils_dateTimeUtils__WEBPACK_IMPORTED_MODULE_6__.formatDisplayDate)(booking.bookingDate),
                                " | ",
                                booking.startTime,
                                " - ",
                                booking.endTime),
                            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_9__.Text, { block: true, variant: "small" },
                                booking.bookerName,
                                " (",
                                booking.bookerEmail,
                                ")")),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, null, renderStatusBadge(booking.bookingStatus)),
                        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_8__.Stack.Item, null, (0,_utils_bookingRules__WEBPACK_IMPORTED_MODULE_5__.canCancelBooking)(booking, bookerEmail, currentUserEmail, isCurrentUserAdmin) && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.DefaultButton, { text: "Cancel", disabled: actionInProgress, onClick: function () { return handleCancelBooking(booking); } })))))); }))) : null))));
};
function getErrorMessage(error) {
    if (error instanceof Error) {
        return error.message;
    }
    return 'An unexpected error occurred. Please try again.';
}
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (DeskBooking);


/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ (() => {
/******/ 	__webpack_require__.h = () => ("b222ffabb13bceaeabb0")
/******/ })();
/******/ 
/******/ }
);
//# sourceMappingURL=desk-booking-web-part.3c1f8dcac3faf2a9d72d.hot-update.js.map