/**
 * 库产物构建配置。
 *
 * 输出目录必须和 demo 站（vite.config.ts，走 dist/）分开：两者都写 dist 的话，
 * 先 build:lib 再 build 会把库产物清掉。库走 dist-lib/。
 *
 * docx / jszip / vue 留作外部依赖 —— 它们体积不小，应该由使用方决定怎么加载
 * （vue 不 external 的话，产物里会打进一份自己的运行时，使用方的 app 里就有两个 Vue）。
 *
 * 入口现在带着 `.vue`（对外组件 WtpEditor），所以必须有 @vitejs/plugin-vue；
 * 组件里的 `<style scoped>` 会被抽成一个独立 CSS 产物（文件名见 verify-lib 的实测结论）。
 */

import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
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
      external: ['vue', 'docx', 'jszip'],
    },
  },
})
