/**
 * index.js
 * Combines all modular checkers into a single global OfficeCheckers namespace.
 */

window.OfficeCheckers = {
  ...window.CheckerHelpers,
  ...window.ExcelCheckers,
  ...window.WordCheckers,
  ...window.PowerPointCheckers
};
