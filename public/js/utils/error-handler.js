/**
 * public/js/utils/error-handler.js
 * ───────────────────────────────
 * Error handler global untuk menangkap error yang tidak tertangani di frontend.
 */
(function () {
  'use strict';

  const ErrorHandler = {
    init() {
      // Tangani unhandled javascript errors
      window.onerror = function (message, source, lineno, colno, error) {
        console.error('[Global Error]', { message, source, lineno, colno, error });
        ErrorHandler.showToast(message || 'Terjadi kesalahan internal javascript.');
        return false; // Tetap tampilkan di console
      };

      // Tangani unhandled promise rejections
      window.onunhandledrejection = function (event) {
        console.error('[Unhandled Promise Rejection]', event.reason);
        const reason = event.reason;
        const msg = reason ? (reason.message || reason) : 'Unhandled promise rejection';
        ErrorHandler.showToast(msg);
      };
    },

    showToast(message) {
      // Buat container toast jika belum ada
      let container = document.getElementById('toast-container');
      if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        container.style.position = 'fixed';
        container.style.bottom = '20px';
        container.style.right = '20px';
        container.style.zIndex = '999999';
        container.style.display = 'flex';
        container.style.flexDirection = 'column';
        container.style.gap = '10px';
        if (document.body) {
          document.body.appendChild(container);
        } else {
          document.addEventListener('DOMContentLoaded', () => {
            document.body.appendChild(container);
          });
        }
      }

      // Buat element toast baru
      const toast = document.createElement('div');
      toast.className = 'error-toast';
      toast.style.background = '#f44336';
      toast.style.color = '#ffffff';
      toast.style.padding = '12px 20px';
      toast.style.borderRadius = '6px';
      toast.style.boxShadow = '0 4px 12px rgba(0, 0, 0, 0.15)';
      toast.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
      toast.style.fontSize = '14px';
      toast.style.fontWeight = '500';
      toast.style.display = 'flex';
      toast.style.alignItems = 'center';
      toast.style.justifyContent = 'space-between';
      toast.style.minWidth = '250px';
      toast.style.animation = 'slideIn 0.3s ease-out';

      toast.innerHTML = `
        <span>⚠️ ${message}</span>
        <button style="background:none; border:none; color:#fff; cursor:pointer; font-weight:bold; margin-left:15px; font-size:16px;">&times;</button>
      `;

      // Event listener close button
      toast.querySelector('button').onclick = function () {
        toast.remove();
      };

      container.appendChild(toast);

      // Hilangkan otomatis setelah 5 detik
      setTimeout(() => {
        toast.style.animation = 'fadeOut 0.5s ease-out forwards';
        setTimeout(() => toast.remove(), 500);
      }, 5000);
    }
  };

  // Tambahkan CSS keyframes untuk slideIn dan fadeOut jika belum ada
  const style = document.createElement('style');
  style.innerHTML = `
    @keyframes slideIn {
      from { transform: translateX(100%); opacity: 0; }
      to { transform: translateX(0); opacity: 1; }
    }
    @keyframes fadeOut {
      from { opacity: 1; }
      to { opacity: 0; }
    }
  `;
  document.head.appendChild(style);

  // Inisialisasi error handler
  ErrorHandler.init();

  window.ErrorHandler = ErrorHandler;
})();
