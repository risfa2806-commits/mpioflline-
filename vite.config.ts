import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import { defineConfig, PluginOption } from 'vite';

const suppressHmrPlugin = (): PluginOption => ({
  name: 'suppress-hmr-head',
  enforce: 'pre',
  transformIndexHtml: {
    order: 'pre',
    handler() {
      return [
        {
          tag: 'script',
          attrs: { type: 'text/javascript' },
          children: `
            (function() {
              try {
                function MockWS() {
                  this.readyState = 1;
                  this.url = '';
                  this.protocol = '';
                  this.extensions = '';
                  this.binaryType = 'blob';
                  this.bufferedAmount = 0;
                  this.onopen = null;
                  this.onclose = null;
                  this.onerror = null;
                  this.onmessage = null;
                  this.close = function() { this.readyState = 3; };
                  this.send = function() {};
                  this.addEventListener = function() {};
                  this.removeEventListener = function() {};
                  this.dispatchEvent = function() { return false; };
                }
                MockWS.CONNECTING = 0;
                MockWS.OPEN = 1;
                MockWS.CLOSING = 2;
                MockWS.CLOSED = 3;
                window.WebSocket = MockWS;
              } catch (e) {}

              var suppress = function(e) {
                var s = String((e && (e.message || (e.reason && (e.reason.message || e.reason)) || e.reason)) || '').toLowerCase();
                if (
                  s.indexOf('websocket') !== -1 ||
                  s.indexOf('ditutup tanpa dibuka') !== -1 ||
                  s.indexOf('closed without being opened') !== -1 ||
                  s.indexOf('network error') !== -1 ||
                  s.indexOf('[vite]') !== -1
                ) {
                  if (e.preventDefault) e.preventDefault();
                  if (e.stopImmediatePropagation) e.stopImmediatePropagation();
                  return true;
                }
              };
              window.addEventListener('error', suppress, true);
              window.addEventListener('unhandledrejection', suppress, true);
            })();
          `,
          injectTo: 'head-prepend',
        },
      ];
    },
  },
});

export default defineConfig(() => {
  return {
    plugins: [suppressHmrPlugin(), react(), tailwindcss()],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
        'decimal.js-light': path.resolve(__dirname, 'node_modules/decimal.js-light/decimal.js'),
        'eventemitter3': path.resolve(__dirname, 'node_modules/eventemitter3/dist/eventemitter3.esm.js'),
      },
    },
    build: {
      chunkSizeWarningLimit: 5000,
      rollupOptions: {
        output: {
          manualChunks(id) {
            if (id.includes('node_modules')) {
              if (id.includes('lucide-react')) return 'vendor-lucide';
              if (id.includes('react') || id.includes('react-dom')) return 'vendor-react';
            }
          },
        },
      },
    },
    server: {
      hmr: false,
      ws: false as const,
      watch: null,
    },
  };
});
