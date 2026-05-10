import logger from './src/utils/logger.js';

/**
 * 必须在其他业务模块之前加载（server.js 首行 import），确保任何早期抛错也能被记录。
 */
process.on('uncaughtException', (err) => {
  logger.error('【致命错误】Uncaught Exception:', err);
  if (err && typeof err === 'object' && err.stack) {
    logger.error(err.stack);
  }
});

process.on('unhandledRejection', (reason, promise) => {
  logger.error('【致命错误】Unhandled Rejection at:', promise, 'reason:', reason);
  if (reason && typeof reason === 'object' && reason.stack) {
    logger.error(reason.stack);
  }
});
