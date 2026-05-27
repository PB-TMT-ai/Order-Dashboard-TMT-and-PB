const isDev = process.env.NODE_ENV === 'development'

export const logger = {
  info: (msg: string, data?: unknown) => console.log(`[INFO] ${msg}`, data || ''),
  warn: (msg: string, data?: unknown) => console.warn(`[WARN] ${msg}`, data || ''),
  error: (msg: string, err?: unknown) => console.error(`[ERROR] ${msg}`, err || ''),
  debug: (msg: string, data?: unknown) => isDev && console.debug(`[DEBUG] ${msg}`, data || '')
}
