/**
 * public/js/utils/api.js
 * ───────────────────────
 * Format respons standar untuk interaksi API/database di frontend.
 */
(function () {
  'use strict';

  const ApiHelper = {
    /**
     * Membungkus fungsi asinkron ke dalam format standard response.
     * @param {Function} asyncFunc
     * @returns {Promise<{success: boolean, data: any, error: string|null}>}
     */
    async handle(asyncFunc) {
      try {
        const data = await asyncFunc();
        return {
          success: true,
          data: data,
          error: null
        };
      } catch (err) {
        console.error('[API Helper Error]', err);
        return {
          success: false,
          data: null,
          error: err.message || 'Terjadi kesalahan sistem'
        };
      }
    }
  };

  window.ApiHelper = ApiHelper;
})();
