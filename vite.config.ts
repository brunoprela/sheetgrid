import { univerPlugin } from '@univerjs/vite-plugin'
import react from '@vitejs/plugin-react'
import { defineConfig, loadEnv } from 'vite'
import packageJson from './package.json'
import path from 'path'

export default ({ mode }: { mode: string }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return defineConfig({
    plugins: [
      react(),
      univerPlugin(),
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, './'),
      },
    },
    define: {
      'process.env.UNIVER_CLIENT_LICENSE': `"${env.UNIVER_CLIENT_LICENSE}"` || '"%%UNIVER_CLIENT_LICENSE_PLACEHOLDER%%"',
      'process.env.UNIVER_VERSION': `"${packageJson.dependencies['@univerjs/presets']}"`,
    },
  })
}
