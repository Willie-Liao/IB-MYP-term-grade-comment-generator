import path from 'path';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    const minimaxBaseUrl = (env.MINIMAX_BASE_URL || 'https://api.minimaxi.com/v1').replace(/\/$/, '');
    const minimaxOrigin = minimaxBaseUrl.replace(/\/v1$/, '');

    return {
      server: {
        port: 3000,
        host: '0.0.0.0',
        proxy: {
          '/api/minimax': {
            target: minimaxOrigin,
            changeOrigin: true,
            rewrite: (path) => path.replace(/^\/api\/minimax/, '/v1'),
            configure: (proxy) => {
              proxy.on('proxyReq', (proxyReq) => {
                if (env.MINIMAX_API_KEY) {
                  proxyReq.setHeader('Authorization', `Bearer ${env.MINIMAX_API_KEY}`);
                }
              });
            },
          },
        },
      },
      plugins: [react()],
      define: {
        'process.env.MINIMAX_API_KEY': JSON.stringify(env.MINIMAX_API_KEY),
        'process.env.MINIMAX_BASE_URL': JSON.stringify(minimaxBaseUrl),
      },
      resolve: {
        alias: {
          '@': path.resolve(__dirname, '.'),
        }
      }
    };
});
