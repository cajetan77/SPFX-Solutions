"use strict";
self["webpackHotUpdateb018cc9a_f2bc_4ba4_90c1_1ebfd12e0aca_0_0_1"]("desk-booking-web-part",{

/***/ 86418:
/*!*********************************************************!*\
  !*** ./lib/webparts/deskBooking/utils/dateTimeUtils.js ***!
  \*********************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   FIXED_BOOKING_END_TIME: () => (/* binding */ FIXED_BOOKING_END_TIME),
/* harmony export */   FIXED_BOOKING_START_TIME: () => (/* binding */ FIXED_BOOKING_START_TIME),
/* harmony export */   formatDisplayDate: () => (/* binding */ formatDisplayDate),
/* harmony export */   formatMinutesToTime: () => (/* binding */ formatMinutesToTime),
/* harmony export */   generateTimeSlots: () => (/* binding */ generateTimeSlots),
/* harmony export */   getTodayDate: () => (/* binding */ getTodayDate),
/* harmony export */   isSameDate: () => (/* binding */ isSameDate),
/* harmony export */   isToday: () => (/* binding */ isToday),
/* harmony export */   isWeekday: () => (/* binding */ isWeekday),
/* harmony export */   isWithinBusinessHours: () => (/* binding */ isWithinBusinessHours),
/* harmony export */   parseSharePointDate: () => (/* binding */ parseSharePointDate),
/* harmony export */   parseTimeToMinutes: () => (/* binding */ parseTimeToMinutes),
/* harmony export */   timesOverlap: () => (/* binding */ timesOverlap),
/* harmony export */   toDateOnlyIsoString: () => (/* binding */ toDateOnlyIsoString),
/* harmony export */   toLocalDateOnly: () => (/* binding */ toLocalDateOnly),
/* harmony export */   toSharePointBookingDate: () => (/* binding */ toSharePointBookingDate)
/* harmony export */ });
var BUSINESS_START_MINUTES = 7 * 60;
var BUSINESS_END_MINUTES = 18 * 60;
var FIXED_BOOKING_START_TIME = '09:00';
var FIXED_BOOKING_END_TIME = '18:00';
var TIME_SLOT_INTERVAL_MINUTES = 30;
function parseTimeToMinutes(time) {
    var parts = time.split(':');
    if (parts.length !== 2) {
        return NaN;
    }
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);
    if (isNaN(hours) || isNaN(minutes)) {
        return NaN;
    }
    return hours * 60 + minutes;
}
function padTwoDigits(value) {
    return value < 10 ? "0".concat(value) : "".concat(value);
}
function formatMinutesToTime(totalMinutes) {
    var hours = Math.floor(totalMinutes / 60);
    var minutes = totalMinutes % 60;
    return "".concat(padTwoDigits(hours), ":").concat(padTwoDigits(minutes));
}
function generateTimeSlots() {
    var slots = [];
    for (var minutes = BUSINESS_START_MINUTES; minutes <= BUSINESS_END_MINUTES; minutes += TIME_SLOT_INTERVAL_MINUTES) {
        slots.push(formatMinutesToTime(minutes));
    }
    return slots;
}
function isWeekday(date) {
    var day = date.getDay();
    return day >= 1 && day <= 5;
}
function isWithinBusinessHours(startTime, endTime) {
    var startMinutes = parseTimeToMinutes(startTime);
    var endMinutes = parseTimeToMinutes(endTime);
    if (isNaN(startMinutes) || isNaN(endMinutes)) {
        return false;
    }
    return startMinutes >= BUSINESS_START_MINUTES
        && endMinutes <= BUSINESS_END_MINUTES
        && startMinutes < endMinutes;
}
function timesOverlap(startA, endA, startB, endB) {
    var startAMinutes = parseTimeToMinutes(startA);
    var endAMinutes = parseTimeToMinutes(endA);
    var startBMinutes = parseTimeToMinutes(startB);
    var endBMinutes = parseTimeToMinutes(endB);
    return startAMinutes < endBMinutes && startBMinutes < endAMinutes;
}
function toDateOnlyIsoString(date) {
    var year = date.getFullYear();
    var month = padTwoDigits(date.getMonth() + 1);
    var day = padTwoDigits(date.getDate());
    return "".concat(year, "-").concat(month, "-").concat(day);
}
function toLocalDateOnly(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0, 0);
}
function isSameDate(a, b) {
    return toDateOnlyIsoString(a) === toDateOnlyIsoString(b);
}
function getTodayDate() {
    return toLocalDateOnly(new Date());
}
function isToday(date) {
    return isSameDate(date, getTodayDate());
}
function toSharePointBookingDate(date) {
    return "".concat(toDateOnlyIsoString(date), "T00:00:00");
}
function parseSharePointDate(value) {
    var parsed = new Date(value);
    if (isNaN(parsed.getTime())) {
        return toLocalDateOnly(new Date());
    }
    return toLocalDateOnly(parsed);
}
function formatDisplayDate(date) {
    return date.toLocaleDateString(undefined, {
        weekday: 'short',
        year: 'numeric',
        month: 'short',
        day: 'numeric'
    });
}


/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ (() => {
/******/ 	__webpack_require__.h = () => ("87963eec734f1bb942de")
/******/ })();
/******/ 
/******/ }
);
//# sourceMappingURL=desk-booking-web-part.c4f8579c3b8197c16757.hot-update.js.map