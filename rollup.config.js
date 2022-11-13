import { nodeResolve } from '@rollup/plugin-node-resolve'
import commonjs from '@rollup/plugin-commonjs'
import { babel } from '@rollup/plugin-babel'
import eslint from '@rollup/plugin-eslint'
import terser from '@rollup/plugin-terser'
import globals from 'rollup-plugin-node-globals'
import builtins from 'rollup-plugin-node-builtins'

const onwarn = warning => {
  if (warning.code === 'CIRCULAR_DEPENDENCY') return

  console.warn(`(!) ${warning.message}`) // eslint-disable-line
}

export default {
  input: 'src/pptxtojson.js',
  onwarn,
  output: [
    {
      file: 'dist/index.js',
      format: 'umd',
      name: 'pptxtojson',
    },
    {
      file: 'dist/index.esm.js',
      format: 'es',
    },
  ],
  plugins: [
    nodeResolve({
      preferBuiltins: false,
    }),
    commonjs(),
    eslint(),
    babel({
      babelHelpers: 'runtime',
      exclude: ['node_modules/**'],
    }),
    terser(),
    globals(),
    builtins(),
  ]
}