/**
 * 库产物构建配置。
 *
 * 输出目录必须和 demo 站（vite.config.ts，走 dist/）分开：两者都写 dist 的话，
 * 先 build:lib 再 build 会把库产物清掉。库走 dist-lib/。
 *
 * docx / jszip 留作外部依赖 —— 它们体积不小，应该由使用方决定怎么加载。
 */

import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'

export default defineConfig({
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  build: {
    outDir: 'dist-lib',
    emptyOutDir: true,
    target: 'es2022',
    lib: {
      entry: fileURLToPath(new URL('./src/lib/index.ts', import.meta.url)),
      name: 'WordToHtml',
      formats: ['es'],
      fileName: () => 'wordtohtml.mjs',
    },
    rollupOptions: {
      external: ['docx', 'jszip'],
    },
  },
})
