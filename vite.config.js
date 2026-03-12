import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    open: true
  }
})


// import { defineConfig } from 'vite'
// import react from '@vitejs/plugin-react'
// import { resolve } from 'path' // Ye line zaroori hai

// export default defineConfig({
//   plugins: [react()],
//   server: {
//     port: 3000,
//     open: true
//   },
//   build: {
//     rollupOptions: {
//       input: {
//         main: resolve(__dirname, 'index.html'),
//         login: resolve(__dirname, 'login.html')
//       }
//     }
//   }
// })