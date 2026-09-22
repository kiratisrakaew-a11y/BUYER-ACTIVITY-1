/**
 * Errors.js — standard error envelope for every api_* function.
 *
 * SPEC §9: every API returns { ok:true, data } or { ok:false, error:{ code, message, details } }.
 * Nothing outside this file should invent an error shape.
 */

/** The only error codes the client knows how to react to (SPEC §9). */
var ERROR_CODES = {
  UNAUTHORIZED: 'UNAUTHORIZED',
  FORBIDDEN: 'FORBIDDEN',
  VALIDATION: 'VALIDATION',
  NOT_FOUND: 'NOT_FOUND',
  CONFLICT: 'CONFLICT',
  RULE_VIOLATION: 'RULE_VIOLATION',
  DUPLICATE: 'DUPLICATE',
  INTERNAL: 'INTERNAL'
};

/**
 * Application error carrying a code from ERROR_CODES.
 * `details` is safe to show the user; stack traces never leave the server.
 */
function AppError(code, message, details) {
  this.name = 'AppError';
  this.code = code;
  this.message = message || code;
  this.details = details || null;
  this.stack = new Error(this.message).stack;
}
AppError.prototype = Object.create(Error.prototype);
AppError.prototype.constructor = AppError;

function isAppError(e) {
  return !!e && e.name === 'AppError' && !!e.code;
}

/** Shorthand constructors, so call sites read as prose. */
var Err = {
  unauthorized: function (m, d) { return new AppError(ERROR_CODES.UNAUTHORIZED, m || 'ไม่มีสิทธิ์เข้าใช้งานระบบ', d); },
  forbidden: function (m, d) { return new AppError(ERROR_CODES.FORBIDDEN, m || 'คุณไม่มีสิทธิ์ทำรายการนี้', d); },
  validation: function (m, d) { return new AppError(ERROR_CODES.VALIDATION, m || 'ข้อมูลไม่ถูกต้อง', d); },
  notFound: function (m, d) { return new AppError(ERROR_CODES.NOT_FOUND, m || 'ไม่พบข้อมูล', d); },
  conflict: function (m, d) { return new AppError(ERROR_CODES.CONFLICT, m || 'ข้อมูลถูกแก้ไขโดยผู้อื่น กรุณาโหลดใหม่', d); },
  ruleViolation: function (m, d) { return new AppError(ERROR_CODES.RULE_VIOLATION, m || 'ไม่ผ่านเงื่อนไขของระบบ', d); },
  duplicate: function (m, d) { return new AppError(ERROR_CODES.DUPLICATE, m || 'ข้อมูลซ้ำกับที่มีอยู่แล้ว', d); },
  internal: function (m, d) { return new AppError(ERROR_CODES.INTERNAL, m || 'เกิดข้อผิดพลาดภายในระบบ', d); }
};
