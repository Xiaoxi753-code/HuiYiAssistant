import fs from 'fs';
import path from 'path';
import winston from 'winston';
import DailyRotateFile from 'winston-daily-rotate-file';

const logsDir = String(process.env.OPENCLAW_LOG_DIR || '').trim() || path.resolve(process.cwd(), 'logs');
const logLevel = String(process.env.LOG_LEVEL || '').trim() || 'info';

function ensureWritableDir(dirPath) {
  try {
    fs.mkdirSync(dirPath, { recursive: true });
    fs.accessSync(dirPath, fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

const canWriteFiles = ensureWritableDir(logsDir);

const consoleFormat = winston.format.combine(
  winston.format.colorize(),
  winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
  winston.format.printf(({ level, message, timestamp, stack, ...rest }) => {
    const extra = Object.keys(rest).length ? ` ${JSON.stringify(rest)}` : '';
    const detail = stack ? `\n${stack}` : '';
    return `${timestamp} ${level}: ${message}${extra}${detail}`;
  }),
);

const fileFormat = winston.format.combine(
  winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
  winston.format.errors({ stack: true }),
  winston.format.json(),
);

const transports = [
  new winston.transports.Console({
    level: logLevel,
    format: consoleFormat,
  }),
];

if (canWriteFiles) {
  transports.push(
    new winston.transports.File({
      filename: path.join(logsDir, 'app.log'),
      level: logLevel,
      format: fileFormat,
    }),
    new winston.transports.File({
      filename: path.join(logsDir, 'error.log'),
      level: 'error',
      format: fileFormat,
    }),
    new DailyRotateFile({
      filename: path.join(logsDir, 'app-%DATE%.log'),
      datePattern: 'YYYY-MM-DD',
      level: logLevel,
      format: fileFormat,
      maxFiles: '14d',
    }),
    new DailyRotateFile({
      filename: path.join(logsDir, 'error-%DATE%.log'),
      datePattern: 'YYYY-MM-DD',
      level: 'error',
      format: fileFormat,
      maxFiles: '30d',
    }),
  );
} else {
  process.stderr.write(`[logger] logs dir not writable: ${logsDir}; falling back to console only.\n`);
}

const logger = winston.createLogger({
  level: logLevel,
  transports,
  exitOnError: false,
});

export default logger;
export { logger, logsDir };
