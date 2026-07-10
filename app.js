const express = require('express');
const cors = require('cors');
const path = require('path');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const mysql = require('mysql2/promise');
const { Sequelize, DataTypes, Op } = require('sequelize');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
});
const PORT = process.env.PORT || 3200;
const JWT_SECRET =  'topi';
const DB_HOST = process.env.MYSQL_HOST || '127.0.0.1';
const DB_PORT = Number(process.env.MYSQL_PORT || 3306);
const DB_USER = process.env.MYSQL_USER || 'root';
const DB_PASSWORD = process.env.MYSQL_PASSWORD || 'dita';
const DB_NAME = 'topi';

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const pad = (value) => String(value).padStart(2, '0');
const toInt = (value, fallback = 0) => {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
};
const toFloat = (value, fallback = 0) => {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};
const toText = (value) => (value === undefined || value === null ? '' : String(value).trim());
const jakartaDate = () => new Intl.DateTimeFormat('sv-SE', {
  timeZone: 'Asia/Jakarta',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
}).format(new Date());
const jakartaTime = () => new Intl.DateTimeFormat('sv-SE', {
  timeZone: 'Asia/Jakarta',
  hour: '2-digit',
  minute: '2-digit',
  hour12: false,
}).format(new Date());
const TIME_RE = /^([01]\d|2[0-3]):([0-5]\d)$/;
const PAYDAY_WEEKDAYS = ['senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu', 'minggu'];
const ROLES = ['owner', 'leader', 'admin', 'karyawan'];
const DEFAULT_OPERATIONAL_ORDER = [
  'Beban Gaji, Upah & Honorer',
  'Beban Bensin, Parkir, Tol Kendaraan',
  'Beban Katering & Makan Karyawan',
  'Beban Listrik',
  'Beban Telekomunikasi',
  'Beban Peralatan Packing',
  'Beban Perlengkapan Toko & Produksi',
  'Beban Retribusi & Sumbangan',
  'Beban Operasional Lainnya',
  'Beban Pembelian Sparepart Mesin',
  'Refund Uang',
  'Biaya Service / Pemeliharaan Kendaraan',
  'Biaya Lainnya',
];
const DEFAULT_MARKETING_ORDER = [
  'Beban Iklan Shopee + Pajak',
  'Beban Iklan Tiktok + Pajak',
  'Beban Iklan Meta Ads + Pajak',
];
const DEFAULT_ADMIN_ORDER = [
  'Biaya Admin POS',
  'Biaya Admin Shopee',
  'Biaya Admin Tiktok',
];
function haversineMeters(lat1, lng1, lat2, lng2) {
  const toRad = (deg) => (deg * Math.PI) / 180;
  const R = 6371000;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2
    + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
function isCheckInLate(window) {
  const [dh, dm] = window.checkInDeadline.split(':').map(Number);
  const [nh, nm] = jakartaTime().split(':').map(Number);
  return (nh * 60 + nm) > (dh * 60 + dm + toInt(window.lateToleranceMinutes, 0));
}
function resolveShiftWindow(user, settings) {
  if (user.Shift) {
    return {
      checkInDeadline: user.Shift.checkInDeadline,
      lateToleranceMinutes: user.Shift.lateToleranceMinutes,
      name: user.Shift.name,
    };
  }
  return {
    checkInDeadline: settings.checkInDeadline,
    lateToleranceMinutes: settings.lateToleranceMinutes,
    name: null,
  };
}
function sundayOfWeek(dateStr) {
  const [y, m, d] = String(dateStr || jakartaDate()).split('-').map(Number);
  const date = new Date(y, m - 1, d);
  const day = date.getDay();
  date.setDate(date.getDate() - day);
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}
function weekBounds(weekStart) {
  const sunday = sundayOfWeek(weekStart);
  const [y, m, d] = sunday.split('-').map(Number);
  const end = new Date(y, m - 1, d);
  end.setDate(end.getDate() + 6);
  return { start: sunday, end: `${end.getFullYear()}-${pad(end.getMonth() + 1)}-${pad(end.getDate())}` };
}
async function isOvertimeDay(workDate) {
  const [y, m, d] = String(workDate).split('-').map(Number);
  const date = new Date(y, m - 1, d);
  if (date.getDay() === 0) {
    return true;
  }
  const holiday = await Holiday.findOne({ where: { date: workDate } });
  return Boolean(holiday);
}
const money = (value) => toInt(value, 0);
const now = () => new Date();

function getPagination(query) {
  const page = Math.max(toInt(query.page, 1), 1);
  const limit = Math.min(Math.max(toInt(query.limit, 10), 5), 100);
  return {
    page,
    limit,
    offset: (page - 1) * limit,
  };
}

function paginationMeta(page, limit, total) {
  return {
    page,
    limit,
    total,
    totalPages: Math.max(Math.ceil(total / limit), 1),
  };
}

function searchWhere(search, fields) {
  const keyword = toText(search);
  if (!keyword) return {};
  return {
    [Op.or]: fields.map((field) => ({
      [field]: {
        [Op.like]: `%${keyword}%`,
      },
    })),
  };
}

async function ensureDatabase() {
  const conn = await mysql.createConnection({
    host: DB_HOST,
    port: DB_PORT,
    user: DB_USER,
    password: DB_PASSWORD,
  });
  await conn.query(`CREATE DATABASE IF NOT EXISTS \`${DB_NAME}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci`);
  await conn.end();
}

const sequelize = new Sequelize(DB_NAME, DB_USER, DB_PASSWORD, {
  host: DB_HOST,
  port: DB_PORT,
  dialect: 'mysql',
  logging: false,
  timezone: '+07:00',
  define: {
    underscored: true,
    timestamps: true,
  },
});

const User = sequelize.define('User', {
  username: {
    type: DataTypes.STRING(100),
    allowNull: false,
  },
  passwordHash: {
    type: DataTypes.STRING(255),
    allowNull: false,
  },
  fullName: {
    type: DataTypes.STRING(150),
    allowNull: false,
  },
  role: {
    type: DataTypes.STRING(20),
    allowNull: false,
    defaultValue: 'karyawan',
    validate: { isIn: [ROLES] },
  },
  jobTitle: {
    type: DataTypes.STRING(120),
    allowNull: true,
  },
  workShift: {
    type: DataTypes.STRING(60),
    allowNull: true,
  },
  phone: {
    type: DataTypes.STRING(40),
    allowNull: true,
  },
  dailyWage: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  mealAllowance: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  overtimeRatePerHour: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  shiftId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
  faceDescriptor: {
    type: DataTypes.JSON,
    allowNull: true,
  },
  faceImageFilename: {
    type: DataTypes.STRING(255),
    allowNull: true,
  },
  faceImageUrl: {
    type: DataTypes.STRING(500),
    allowNull: true,
  },
  faceImageSize: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
  faceRegisteredAt: {
    type: DataTypes.DATE,
    allowNull: true,
  },
  lastLoginAt: {
    type: DataTypes.DATE,
    allowNull: true,
  },
  active: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: true,
  },
}, {
  tableName: 'users',
});

const Platform = sequelize.define('Platform', {
  name: {
    type: DataTypes.STRING(80),
    allowNull: false,
  },
  description: {
    type: DataTypes.STRING(255),
    allowNull: true,
  },
  active: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: true,
  },
}, {
  tableName: 'platforms',
});

const Store = sequelize.define('Store', {
  platform: {
    type: DataTypes.STRING(80),
    allowNull: false,
  },
  name: {
    type: DataTypes.STRING(120),
    allowNull: false,
  },
  platformStore: {
    type: DataTypes.STRING(200),
    allowNull: false,
  },
  taxRate: {
    type: DataTypes.FLOAT,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'stores',
});

const Product = sequelize.define('Product', {
  sku: {
    type: DataTypes.STRING(120),
    allowNull: false,
  },
  name: {
    type: DataTypes.STRING(200),
    allowNull: false,
  },
  variant: {
    type: DataTypes.STRING(160),
    allowNull: false,
  },
  hpp: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'products',
});

const Sale = sequelize.define('Sale', {
  saleDate: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  channel: {
    type: DataTypes.STRING(60),
    allowNull: false,
  },
  storeName: {
    type: DataTypes.STRING(160),
    allowNull: false,
  },
  orderNumber: {
    type: DataTypes.STRING(120),
    allowNull: true,
  },
  customer: {
    type: DataTypes.STRING(160),
    allowNull: true,
  },
  sku: {
    type: DataTypes.STRING(120),
    allowNull: false,
  },
  qty: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 1,
  },
  price: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  subtotal: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  adminFee: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  hpp: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  totalHpp: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'sales',
});

const OperationalExpense = sequelize.define('OperationalExpense', {
  expenseDate: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  category: {
    type: DataTypes.STRING(160),
    allowNull: false,
  },
  description: {
    type: DataTypes.STRING(255),
    allowNull: false,
  },
  nominal: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  paymentMethod: {
    type: DataTypes.STRING(80),
    allowNull: false,
    defaultValue: 'Transfer',
  },
}, {
  tableName: 'operational_expenses',
});

const MarketingExpense = sequelize.define('MarketingExpense', {
  expenseDate: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  category: {
    type: DataTypes.STRING(160),
    allowNull: false,
  },
  platformStore: {
    type: DataTypes.STRING(200),
    allowNull: false,
  },
  description: {
    type: DataTypes.STRING(255),
    allowNull: false,
  },
  nominal: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  taxRate: {
    type: DataTypes.FLOAT,
    allowNull: false,
    defaultValue: 0,
  },
  totalTax: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'marketing_expenses',
});

const Attendance = sequelize.define('Attendance', {
  workDate: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  checkInAt: {
    type: DataTypes.DATE,
    allowNull: true,
  },
  checkOutAt: {
    type: DataTypes.DATE,
    allowNull: true,
  },
  status: {
    type: DataTypes.ENUM('hadir', 'izin', 'sakit', 'cuti'),
    allowNull: false,
    defaultValue: 'hadir',
  },
  method: {
    type: DataTypes.STRING(60),
    allowNull: false,
    defaultValue: 'wajah',
  },
  confidence: {
    type: DataTypes.FLOAT,
    allowNull: true,
  },
  note: {
    type: DataTypes.STRING(255),
    allowNull: true,
  },
  isLate: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: false,
  },
  checkInDistanceMeters: {
    type: DataTypes.FLOAT,
    allowNull: true,
  },
  checkOutDistanceMeters: {
    type: DataTypes.FLOAT,
    allowNull: true,
  },
  checkInPhotoUrl: {
    type: DataTypes.STRING(500),
    allowNull: true,
  },
  checkOutPhotoUrl: {
    type: DataTypes.STRING(500),
    allowNull: true,
  },
  checkInLocationName: {
    type: DataTypes.STRING(255),
    allowNull: true,
  },
  checkOutLocationName: {
    type: DataTypes.STRING(255),
    allowNull: true,
  },
  shiftName: {
    type: DataTypes.STRING(60),
    allowNull: true,
  },
  overtimeHours: {
    type: DataTypes.FLOAT,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'attendance',
});

const AttendanceSetting = sequelize.define('AttendanceSetting', {
  checkInStart: {
    type: DataTypes.STRING(5),
    allowNull: false,
    defaultValue: '07:00',
  },
  checkInDeadline: {
    type: DataTypes.STRING(5),
    allowNull: false,
    defaultValue: '08:00',
  },
  checkOutStart: {
    type: DataTypes.STRING(5),
    allowNull: false,
    defaultValue: '17:00',
  },
  lateToleranceMinutes: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  radiusEnabled: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: false,
  },
  radiusMeters: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 100,
  },
  locationLabel: {
    type: DataTypes.STRING(160),
    allowNull: true,
  },
  latitude: {
    type: DataTypes.DOUBLE,
    allowNull: true,
  },
  longitude: {
    type: DataTypes.DOUBLE,
    allowNull: true,
  },
  paydayWeekday: {
    type: DataTypes.STRING(10),
    allowNull: true,
  },
}, {
  tableName: 'attendance_settings',
});

const Shift = sequelize.define('Shift', {
  name: {
    type: DataTypes.STRING(60),
    allowNull: false,
  },
  checkInStart: {
    type: DataTypes.STRING(5),
    allowNull: false,
    defaultValue: '07:00',
  },
  checkInDeadline: {
    type: DataTypes.STRING(5),
    allowNull: false,
    defaultValue: '08:00',
  },
  checkOutStart: {
    type: DataTypes.STRING(5),
    allowNull: false,
    defaultValue: '17:00',
  },
  lateToleranceMinutes: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  active: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: true,
  },
}, {
  tableName: 'shifts',
  hooks: {
    beforeDestroy: async (instance) => {
      await User.update({ shiftId: null }, { where: { shiftId: instance.id } });
    },
  },
});

const SalaryPayment = sequelize.define('SalaryPayment', {
  weekStart: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  paid: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: false,
  },
  paidAt: {
    type: DataTypes.DATE,
    allowNull: true,
  },
  bonus: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  thr: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  deduction: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'salary_payments',
  indexes: [
    { unique: true, fields: ['user_id', 'week_start'] },
  ],
});

const Holiday = sequelize.define('Holiday', {
  date: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  label: {
    type: DataTypes.STRING(160),
    allowNull: true,
  },
}, {
  tableName: 'holidays',
  indexes: [
    { unique: true, fields: ['date'] },
  ],
});

const ExpenseCategory = sequelize.define('ExpenseCategory', {
  name: {
    type: DataTypes.STRING(160),
    allowNull: false,
  },
  kind: {
    type: DataTypes.ENUM('operasional', 'marketing'),
    allowNull: false,
  },
}, {
  tableName: 'expense_categories',
  indexes: [
    { unique: true, fields: ['name', 'kind'] },
  ],
});

Platform.hasMany(Store, { foreignKey: 'platformId' });
Store.belongsTo(Platform, { foreignKey: 'platformId' });
Platform.hasMany(Sale, { foreignKey: 'platformId' });
Sale.belongsTo(Platform, { foreignKey: 'platformId' });
Platform.hasMany(MarketingExpense, { foreignKey: 'platformId' });
MarketingExpense.belongsTo(Platform, { foreignKey: 'platformId' });
User.hasMany(Attendance, { foreignKey: 'userId' });
Attendance.belongsTo(User, { foreignKey: 'userId' });
Shift.hasMany(User, { foreignKey: 'shiftId' });
User.belongsTo(Shift, { foreignKey: 'shiftId' });
User.hasMany(SalaryPayment, { foreignKey: 'userId' });
SalaryPayment.belongsTo(User, { foreignKey: 'userId' });

async function getAttendanceSettings() {
  let row = await AttendanceSetting.findOne({ order: [['id', 'ASC']] });
  if (!row) {
    row = await AttendanceSetting.create({});
  }
  return row;
}

const serializeUser = (user) => ({
  id: user.id,
  username: user.username,
  fullName: user.fullName,
  role: user.role,
  jobTitle: user.jobTitle || '',
  workShift: user.workShift || '',
  phone: user.phone || '',
  dailyWage: user.dailyWage || 0,
  mealAllowance: user.mealAllowance || 0,
  overtimeRatePerHour: user.overtimeRatePerHour || 0,
  shiftId: user.shiftId || null,
  shiftName: user.Shift ? user.Shift.name : null,
  shiftCheckInStart: user.Shift ? user.Shift.checkInStart : null,
  shiftCheckInDeadline: user.Shift ? user.Shift.checkInDeadline : null,
  shiftCheckOutStart: user.Shift ? user.Shift.checkOutStart : null,
  active: Boolean(user.active),
  faceRegistered: Boolean(user.faceDescriptor || user.faceImageUrl),
  faceImageFilename: user.faceImageFilename || '',
  faceImageUrl: user.faceImageUrl || '',
  faceImageSize: user.faceImageSize || 0,
  faceRegisteredAt: user.faceRegisteredAt,
  lastLoginAt: user.lastLoginAt,
  createdAt: user.createdAt,
  updatedAt: user.updatedAt,
});

const serializePlatform = (row) => ({
  id: row.id,
  name: row.name,
  description: row.description || '',
  active: Boolean(row.active),
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeStore = (row) => ({
  id: row.id,
  platformId: row.platformId || '',
  platform: row.platform,
  platformData: row.Platform ? serializePlatform(row.Platform) : null,
  name: row.name,
  platformStore: row.platformStore,
  taxRate: row.taxRate,
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeShift = (row) => ({
  id: row.id,
  name: row.name,
  checkInStart: row.checkInStart,
  checkInDeadline: row.checkInDeadline,
  checkOutStart: row.checkOutStart,
  lateToleranceMinutes: row.lateToleranceMinutes,
  active: Boolean(row.active),
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeProduct = (row) => ({
  id: row.id,
  sku: row.sku,
  name: row.name,
  variant: row.variant,
  hpp: row.hpp,
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeSale = (row, { includeHpp = true } = {}) => ({
  id: row.id,
  saleDate: row.saleDate,
  platformId: row.platformId || '',
  platformData: row.Platform ? serializePlatform(row.Platform) : null,
  channel: row.channel,
  storeName: row.storeName,
  orderNumber: row.orderNumber || '',
  customer: row.customer || '',
  sku: row.sku,
  qty: row.qty,
  price: row.price,
  subtotal: row.subtotal,
  adminFee: row.adminFee,
  ...(includeHpp ? { hpp: row.hpp, totalHpp: row.totalHpp } : {}),
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeOperational = (row) => ({
  id: row.id,
  expenseDate: row.expenseDate,
  category: row.category,
  description: row.description,
  nominal: row.nominal,
  paymentMethod: row.paymentMethod,
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeMarketing = (row) => ({
  id: row.id,
  expenseDate: row.expenseDate,
  category: row.category,
  platformId: row.platformId || '',
  platformStore: row.platformStore,
  platformData: row.Platform ? serializePlatform(row.Platform) : null,
  description: row.description,
  nominal: row.nominal,
  taxRate: row.taxRate || 0,
  totalTax: row.totalTax || 0,
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeAttendance = (row) => ({
  id: row.id,
  userId: row.userId,
  workDate: row.workDate,
  checkInAt: row.checkInAt,
  checkOutAt: row.checkOutAt,
  status: row.status,
  method: row.method,
  confidence: row.confidence,
  note: row.note || '',
  isLate: Boolean(row.isLate),
  checkInDistanceMeters: row.checkInDistanceMeters === undefined ? null : row.checkInDistanceMeters,
  checkOutDistanceMeters: row.checkOutDistanceMeters === undefined ? null : row.checkOutDistanceMeters,
  checkInPhotoUrl: row.checkInPhotoUrl || '',
  checkOutPhotoUrl: row.checkOutPhotoUrl || '',
  checkInLocationName: row.checkInLocationName || '',
  checkOutLocationName: row.checkOutLocationName || '',
  shiftName: row.shiftName || '',
  overtimeHours: row.overtimeHours || 0,
  user: row.User ? serializeUser(row.User) : undefined,
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeAttendanceSettings = (row) => ({
  id: row.id,
  checkInStart: row.checkInStart,
  checkInDeadline: row.checkInDeadline,
  checkOutStart: row.checkOutStart,
  lateToleranceMinutes: row.lateToleranceMinutes,
  radiusEnabled: Boolean(row.radiusEnabled),
  radiusMeters: row.radiusMeters,
  locationLabel: row.locationLabel || '',
  latitude: row.latitude === undefined ? null : row.latitude,
  longitude: row.longitude === undefined ? null : row.longitude,
  paydayWeekday: row.paydayWeekday || null,
});

const serializeHoliday = (row) => ({
  id: row.id,
  date: row.date,
  label: row.label || '',
});

const serializeExpenseCategory = (row) => ({
  id: row.id,
  name: row.name,
  kind: row.kind,
});

function auth(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (!token) {
    return res.status(401).json({ ok: false, message: 'Sesi belum masuk.' });
  }
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    req.auth = payload;
    return next();
  } catch (error) {
    return res.status(401).json({ ok: false, message: 'Sesi sudah kedaluwarsa. Masuk ulang.' });
  }
}

const ownerOnly = (req, res, next) => {
  if (req.user?.role !== 'owner') {
    return res.status(403).json({ ok: false, message: 'Akses ini hanya untuk owner.' });
  }
  return next();
};

const ownerOrLeader = (req, res, next) => {
  if (!['owner', 'leader'].includes(req.user?.role)) {
    return res.status(403).json({ ok: false, message: 'Akses ini hanya untuk owner atau leader.' });
  }
  return next();
};

const salesRecapAccess = (req, res, next) => {
  if (!['owner', 'leader', 'admin'].includes(req.user?.role)) {
    return res.status(403).json({ ok: false, message: 'Akses ini tidak tersedia untuk role Anda.' });
  }
  return next();
};

async function loadCurrentUser(req, res, next) {
  try {
    const user = await User.findByPk(req.auth.id, { include: [Shift] });
    if (!user || !user.active) {
      return res.status(401).json({ ok: false, message: 'Akun tidak aktif.' });
    }
    req.user = user;
    return next();
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat sesi.' });
  }
}

const authRequired = [auth, loadCurrentUser];

function buildSalaryRow(user, presentDays, attendanceRecords, paymentRow) {
  const dailyWage = user.dailyWage || 0;
  const mealAllowance = user.mealAllowance || 0;
  const overtimeRatePerHour = user.overtimeRatePerHour || 0;
  const overtimeHours = (attendanceRecords || []).reduce((sum, record) => sum + (record.overtimeHours || 0), 0);
  const overtimePay = Math.round(overtimeHours * overtimeRatePerHour);
  const bonus = paymentRow?.bonus || 0;
  const thr = paymentRow?.thr || 0;
  const deduction = paymentRow?.deduction || 0;
  const salaryBase = presentDays * dailyWage;
  const mealTotal = presentDays * mealAllowance;
  const gross = salaryBase + mealTotal + overtimePay + bonus + thr;
  const net = gross - deduction;
  return {
    user: serializeUser(user),
    presentDays,
    dailyWage,
    mealAllowance,
    overtimeHours,
    overtimePay,
    bonus,
    thr,
    allowanceTotal: mealTotal,
    grossSalary: gross,
    deduction,
    netSalary: net,
    paid: Boolean(paymentRow?.paid),
    paidAt: paymentRow?.paidAt || null,
  };
}

async function buildProfitSummary(from, to) {
  const saleWhere = dateWhere('saleDate', from, to);
  const expenseWhere = dateWhere('expenseDate', from, to);
  const reportDate = new Intl.DateTimeFormat('id-ID', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });

  const salesRows = await Sale.findAll({
    where: saleWhere,
    raw: true,
    order: [['saleDate', 'ASC'], ['createdAt', 'ASC']],
  });
  const operationalRows = await OperationalExpense.findAll({
    where: expenseWhere,
    raw: true,
    order: [['expenseDate', 'ASC'], ['createdAt', 'ASC']],
  });
  const marketingRows = await MarketingExpense.findAll({
    where: expenseWhere,
    raw: true,
    order: [['expenseDate', 'ASC'], ['createdAt', 'ASC']],
  });

  const DEFAULT_REVENUE_ORDER = ['Penjualan Offline', 'Penjualan Shopee', 'Penjualan Tiktok', 'Penjualan Lainnya'];
  const DEFAULT_NON_OPERATIONAL_INCOME_ORDER = ['Ongkos Kirim'];
  const DEFAULT_NON_OPERATIONAL_EXPENSE_ORDER = [
    'Beban Adm. Bank & Buku Cek/Giro',
    'Beban diluar usaha lainnya',
  ];

  const normalize = (value) => toText(value).toLowerCase().replace(/\s+/g, ' ');
  const readAmount = (row, ...keys) => {
    for (const key of keys) {
      if (row[key] !== undefined && row[key] !== null) {
        return toInt(row[key], 0);
      }
    }
    return 0;
  };
  const bucket = (map, label, value) => {
    const amount = Math.max(0, toInt(value, 0));
    map.set(label, (map.get(label) || 0) + amount);
  };
  const pickByOrder = (map, order) => {
    const used = new Set();
    const rows = [];
    for (const label of order) {
      rows.push({ label, amount: map.get(label) || 0 });
      used.add(label);
    }
    const extras = [...map.keys()]
      .filter((label) => !used.has(label))
      .sort((a, b) => a.localeCompare(b, 'id-ID'));
    for (const label of extras) {
      rows.push({ label, amount: map.get(label) || 0 });
    }
    return rows;
  };
  const formatPeriodLabel = (start, end) => {
    if (!start && !end) return 'Semua data';
    if (start && end && start === end) {
      return reportDate.format(new Date(`${start}T00:00:00`));
    }
    if (!start) return `Sampai ${reportDate.format(new Date(`${end}T00:00:00`))}`;
    if (!end) return `Mulai ${reportDate.format(new Date(`${start}T00:00:00`))}`;
    return `${reportDate.format(new Date(`${start}T00:00:00`))} - ${reportDate.format(new Date(`${end}T00:00:00`))}`;
  };

  const revenueMap = new Map();
  const adminMap = new Map();
  const operationalMap = new Map();
  const marketingMap = new Map();
  const nonOperationalIncomeMap = new Map();
  const nonOperationalExpenseMap = new Map();

  const resolveRevenueLabel = (channel, storeName) => {
    const value = `${normalize(channel)} ${normalize(storeName)}`.trim();
    if (!value || value.includes('offline') || value.includes('toko') || value.includes('store') || value.includes('pos')) {
      return 'Penjualan Offline';
    }
    if (value.includes('shopee')) return 'Penjualan Shopee';
    if (value.includes('tiktok') || value.includes('tik tok') || value.includes('tt')) return 'Penjualan Tiktok';
    return 'Penjualan Lainnya';
  };

  const resolveAdminLabel = (channel, storeName) => {
    const value = `${normalize(channel)} ${normalize(storeName)}`.trim();
    if (value.includes('shopee')) return 'Biaya Admin Shopee';
    if (value.includes('tiktok') || value.includes('tik tok') || value.includes('tt')) return 'Biaya Admin Tiktok';
    if (value.includes('offline') || value.includes('toko') || value.includes('pos') || value.includes('store')) return 'Biaya Admin POS';
    return 'Biaya Admin POS';
  };

  const resolveMarketingLabel = (category, platformStore) => {
    const value = `${normalize(category)} ${normalize(platformStore)}`.trim();
    if (value.includes('shopee')) return 'Beban Iklan Shopee + Pajak';
    if (value.includes('tiktok') || value.includes('tik tok') || value.includes('tt')) return 'Beban Iklan Tiktok + Pajak';
    if (value.includes('meta') || value.includes('facebook') || value.includes('instagram')) return 'Beban Iklan Meta Ads + Pajak';
    const cleaned = toText(category).replace(/^beban\s+/i, '').trim() || 'Iklan Lainnya';
    return `Beban ${cleaned} + Pajak`;
  };

  for (const row of salesRows) {
    bucket(revenueMap, resolveRevenueLabel(row.channel, row.storeName), readAmount(row, 'subtotal', 'subTotal'));
    bucket(adminMap, resolveAdminLabel(row.channel, row.storeName), readAmount(row, 'adminFee', 'admin_fee'));
  }

  for (const row of operationalRows) {
    const label = toText(row.category) || 'Beban Operasional Lainnya';
    bucket(operationalMap, label, readAmount(row, 'nominal'));
  }

  for (const row of marketingRows) {
    const label = resolveMarketingLabel(row.category, row.platformStore);
    bucket(marketingMap, label, readAmount(row, 'nominal') + readAmount(row, 'totalTax', 'total_tax'));
  }

  const revenueRows = pickByOrder(revenueMap, DEFAULT_REVENUE_ORDER);
  const adminRows = pickByOrder(adminMap, DEFAULT_ADMIN_ORDER);
  const operationalRowsGrouped = pickByOrder(operationalMap, DEFAULT_OPERATIONAL_ORDER);
  const marketingRowsGrouped = pickByOrder(marketingMap, DEFAULT_MARKETING_ORDER);
  const nonOperationalIncomeRows = pickByOrder(nonOperationalIncomeMap, DEFAULT_NON_OPERATIONAL_INCOME_ORDER);
  const nonOperationalExpenseRows = pickByOrder(nonOperationalExpenseMap, DEFAULT_NON_OPERATIONAL_EXPENSE_ORDER);

  const revenue = revenueRows.reduce((sum, row) => sum + row.amount, 0);
  const hpp = salesRows.reduce((sum, row) => sum + readAmount(row, 'totalHpp', 'total_hpp'), 0);
  const operational = operationalRowsGrouped.reduce((sum, row) => sum + row.amount, 0);
  const marketingNominal = marketingRows.reduce((sum, row) => sum + readAmount(row, 'nominal'), 0);
  const marketingTax = marketingRows.reduce((sum, row) => sum + readAmount(row, 'totalTax', 'total_tax'), 0);
  const marketing = marketingRowsGrouped.reduce((sum, row) => sum + row.amount, 0);
  const adminFee = adminRows.reduce((sum, row) => sum + row.amount, 0);
  const nonOperationalIncome = nonOperationalIncomeRows.reduce((sum, row) => sum + row.amount, 0);
  const nonOperationalExpense = nonOperationalExpenseRows.reduce((sum, row) => sum + row.amount, 0);
  const grossProfit = revenue - hpp;
  const operatingExpense = operational + marketing + adminFee;
  const operatingProfit = grossProfit - operatingExpense;
  const netOther = nonOperationalIncome - nonOperationalExpense;
  const netProfit = operatingProfit + netOther;

  return {
    period: { from, to },
    periodLabel: formatPeriodLabel(from, to),
    counts: {
      sales: salesRows.length,
      operational: operationalRows.length,
      marketing: marketingRows.length,
    },
    sections: {
      revenue: revenueRows,
      costOfGoods: [{ label: 'Beban Pokok Penjualan', amount: hpp }],
      operational: operationalRowsGrouped,
      marketing: marketingRowsGrouped,
      admin: adminRows,
      nonOperationalIncome: nonOperationalIncomeRows,
      nonOperationalExpense: nonOperationalExpenseRows,
    },
    totals: {
      revenue,
      hpp,
      grossProfit,
      operational,
      marketing,
      marketingNominal,
      marketingTax,
      adminFee,
      operatingExpense,
      operatingProfit,
      nonOperationalIncome,
      nonOperationalExpense,
      netOther,
      netProfit,
    },
  };
}

function dateWhere(field, from, to) {
  if (!from && !to) {
    return {};
  }
  const clause = {};
  if (from) clause[Op.gte] = from;
  if (to) clause[Op.lte] = to;
  return { [field]: clause };
}

async function computeGajiTotalForPeriod(from, to) {
  const attendanceWhere = dateWhere('workDate', from, to);
  const paymentWhere = dateWhere('weekStart', from, to);
  const users = await User.findAll({ where: { active: true } });
  const attendanceRows = await Attendance.findAll({ where: attendanceWhere });
  const payments = await SalaryPayment.findAll({ where: paymentWhere });

  const attendanceByUser = new Map();
  for (const row of attendanceRows) {
    if (!attendanceByUser.has(row.userId)) attendanceByUser.set(row.userId, []);
    attendanceByUser.get(row.userId).push(row);
  }
  const paymentsByUser = new Map();
  for (const payment of payments) {
    if (!paymentsByUser.has(payment.userId)) paymentsByUser.set(payment.userId, []);
    paymentsByUser.get(payment.userId).push(payment);
  }

  let total = 0;
  for (const user of users) {
    const records = attendanceByUser.get(user.id) || [];
    const presentDays = records.filter((record) => record.status === 'hadir').length;
    const overtimeHours = records.reduce((sum, record) => sum + (record.overtimeHours || 0), 0);
    const overtimePay = Math.round(overtimeHours * (user.overtimeRatePerHour || 0));
    const salaryBase = presentDays * (user.dailyWage || 0);
    const mealTotal = presentDays * (user.mealAllowance || 0);
    const userPayments = paymentsByUser.get(user.id) || [];
    const bonus = userPayments.reduce((sum, payment) => sum + (payment.bonus || 0), 0);
    const thr = userPayments.reduce((sum, payment) => sum + (payment.thr || 0), 0);
    const deduction = userPayments.reduce((sum, payment) => sum + (payment.deduction || 0), 0);
    total += salaryBase + mealTotal + overtimePay + bonus + thr - deduction;
  }
  return total;
}

async function buildDashboard(from, to, { includeHpp = true } = {}) {
  const profitSummary = await buildProfitSummary(from, to);
  const gajiKaryawan = await computeGajiTotalForPeriod(from, to);
  const totalBebanPeriode = profitSummary.totals.operational + profitSummary.totals.marketing + profitSummary.totals.adminFee + gajiKaryawan;
  const saleWhere = dateWhere('saleDate', from, to);
  const expenseWhere = dateWhere('expenseDate', from, to);
  const [
    totalEmployees,
    totalFaces,
    totalStores,
    totalProducts,
    attendanceToday,
    recentSales,
    recentOperational,
    recentMarketing,
    recentAttendance,
  ] = await Promise.all([
    User.count({ where: { role: 'karyawan', active: true } }),
    User.count({
      where: {
        role: 'karyawan',
        active: true,
        [Op.or]: [
          { faceDescriptor: { [Op.ne]: null } },
          { faceImageUrl: { [Op.ne]: null } },
        ],
      },
    }),
    Store.count(),
    Product.count(),
    Attendance.count({
      where: {
        workDate: jakartaDate(),
      },
    }),
    Sale.findAll({
      where: saleWhere,
      order: [['saleDate', 'DESC'], ['createdAt', 'DESC']],
      limit: 5,
    }),
    OperationalExpense.findAll({
      where: expenseWhere,
      order: [['createdAt', 'DESC']],
      limit: 5,
    }),
    MarketingExpense.findAll({
      where: expenseWhere,
      order: [['createdAt', 'DESC']],
      limit: 5,
    }),
    Attendance.findAll({
      include: [{ model: User }],
      order: [['createdAt', 'DESC']],
      limit: 5,
    }),
  ]);

  return {
    period: { from, to },
    counts: {
      employees: totalEmployees,
      faces: totalFaces,
      stores: totalStores,
      products: totalProducts,
      sales: profitSummary.counts.sales,
      operational: profitSummary.counts.operational,
      marketing: profitSummary.counts.marketing,
      attendanceToday,
    },
    totals: {
      ...profitSummary.totals,
      gajiKaryawan,
      totalBebanPeriode,
    },
    recent: {
      sales: recentSales.map((row) => serializeSale(row, { includeHpp })),
      operational: recentOperational.map(serializeOperational),
      marketing: recentMarketing.map(serializeMarketing),
      attendance: recentAttendance.map(serializeAttendance),
    },
  };
}

async function migrateOwnerRole() {
  await sequelize.query("UPDATE `users` SET `role` = 'owner' WHERE `role` = 'admin'");
}

async function createInitialAdmin() {
  const existing = await User.findOne({ where: { username: 'admin' } });
  if (existing) {
    if (existing.role !== 'owner') {
      existing.role = 'owner';
      await existing.save();
    }
    return;
  }
  const passwordHash = await bcrypt.hash('admin123', 10);
  await User.create({
    username: 'admin',
    passwordHash,
    fullName: 'Administrator SDS Apps',
    role: 'owner',
    jobTitle: 'Administrator',
    workShift: 'Utama',
    active: true,
  });
}

function createSimpleCrudRoutes(prefix, model, serializer, parser, options = {}) {
  const writeMiddleware = options.writeMiddleware || ownerOnly;
  app.get(`/api/${prefix}`, ...authRequired, async (req, res) => {
    try {
      const where = options.searchFields ? searchWhere(req.query.search, options.searchFields) : {};
      if (options.paginated) {
        const { page, limit, offset } = getPagination(req.query);
        const { count, rows } = await model.findAndCountAll({
          where,
          include: options.include || [],
          order: [['createdAt', 'DESC']],
          limit,
          offset,
        });
        return res.json({
          ok: true,
          data: rows.map(serializer),
          pagination: paginationMeta(page, limit, count),
        });
      }
      const rows = await model.findAll({
        where,
        include: options.include || [],
        order: [['createdAt', 'DESC']],
      });
      return res.json({ ok: true, data: rows.map(serializer) });
    } catch (error) {
      return res.status(500).json({ ok: false, message: 'Gagal memuat data.' });
    }
  });

  app.post(`/api/${prefix}`, ...authRequired, writeMiddleware, async (req, res) => {
    try {
      const payload = await parser(req.body, null);
      const row = await model.create(payload);
      const output = options.reload ? await model.findByPk(row.id, { include: options.include || [] }) : row;
      return res.status(201).json({ ok: true, data: serializer(output) });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan data.' });
    }
  });

  app.put(`/api/${prefix}/:id`, ...authRequired, writeMiddleware, async (req, res) => {
    try {
      const row = await model.findByPk(req.params.id);
      if (!row) {
        return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
      }
      const payload = await parser(req.body, row);
      await row.update(payload);
      const output = options.reload ? await model.findByPk(row.id, { include: options.include || [] }) : row;
      return res.json({ ok: true, data: serializer(output) });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error.message || 'Gagal memperbarui data.' });
    }
  });

  app.delete(`/api/${prefix}/:id`, ...authRequired, writeMiddleware, async (req, res) => {
    try {
      const row = await model.findByPk(req.params.id);
      if (!row) {
        return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
      }
      await row.destroy();
      return res.json({ ok: true, message: 'Data berhasil dihapus.' });
    } catch (error) {
      return res.status(500).json({ ok: false, message: 'Gagal menghapus data.' });
    }
  });
}

async function parseShift(body) {
  const name = toText(body.name);
  if (!name) {
    throw new Error('Nama shift wajib diisi.');
  }
  const checkInStart = toText(body.checkInStart);
  const checkInDeadline = toText(body.checkInDeadline);
  const checkOutStart = toText(body.checkOutStart);
  for (const [label, value] of [['Jam mulai masuk', checkInStart], ['Batas telat', checkInDeadline], ['Jam mulai pulang', checkOutStart]]) {
    if (!TIME_RE.test(value)) {
      throw new Error(`${label} harus berformat HH:mm.`);
    }
  }
  return {
    name,
    checkInStart,
    checkInDeadline,
    checkOutStart,
    lateToleranceMinutes: toInt(body.lateToleranceMinutes, 0),
    active: body.active === undefined ? true : Boolean(body.active),
  };
}

async function parseStore(body) {
  const platform = toText(body.platform);
  const name = toText(body.name);
  const platformStore = toText(body.platformStore) || `${platform} ${name}`.trim();
  if (!platform || !name) {
    throw new Error('Platform dan nama toko wajib diisi.');
  }
  return {
    platformId: null,
    platform,
    name,
    platformStore,
    taxRate: toFloat(body.taxRate, 0),
  };
}

function parseProduct(body) {
  const sku = toText(body.sku);
  const name = toText(body.name);
  const variant = toText(body.variant);
  if (!sku || !name || !variant) {
    throw new Error('SKU, nama barang, dan jenis wajib diisi.');
  }
  return {
    sku,
    name,
    variant,
    hpp: money(body.hpp),
  };
}

function parseOperational(body) {
  const expenseDate = toText(body.expenseDate || body.date);
  const category = toText(body.category);
  const description = toText(body.description);
  if (!expenseDate || !category || !description) {
    throw new Error('Tanggal, kategori, dan keterangan wajib diisi.');
  }
  return {
    expenseDate,
    category,
    description,
    nominal: money(body.nominal),
    paymentMethod: toText(body.paymentMethod) || 'Transfer',
  };
}

async function parseMarketing(body) {
  const expenseDate = toText(body.expenseDate || body.date);
  const category = toText(body.category);
  const platformStore = toText(body.platformStore || body.platformAndStore);
  const description = toText(body.description || body.note);
  if (!expenseDate || !category || !platformStore || !description) {
    throw new Error('Tanggal, kategori, platform & toko, dan keterangan wajib diisi.');
  }
  const nominal = money(body.nominal);
  const taxRate = body.taxRate === undefined || body.taxRate === null || body.taxRate === '' ? 0 : Number(body.taxRate);
  const totalTax = body.totalTax === undefined || body.totalTax === null || body.totalTax === ''
    ? Math.round(nominal * (Number.isFinite(taxRate) ? taxRate : 0))
    : money(body.totalTax);
  return {
    expenseDate,
    category,
    platformId: null,
    platformStore,
    description,
    nominal,
    taxRate: Number.isFinite(taxRate) ? taxRate : 0,
    totalTax,
  };
}

async function parseSale(body, existingRow) {
  const saleDate = toText(body.saleDate || body.date);
  const channelInput = toText(body.channel);
  const channel = channelInput;
  const storeName = toText(body.storeName || body.platformStore);
  const orderNumber = toText(body.orderNumber || body.noPesanan);
  const sku = toText(body.sku);
  if (!saleDate || !channel || !storeName || !sku) {
    throw new Error('Tanggal, channel, toko, dan SKU wajib diisi.');
  }
  const product = await Product.findOne({ where: { sku } });
  const qty = Math.max(1, money(body.qty || 1));
  const price = money(body.price);
  const baseHpp = product ? product.hpp : money(body.hpp || 0);
  return {
    saleDate,
    platformId: null,
    channel,
    storeName,
    orderNumber,
    customer: toText(body.customer),
    sku,
    qty,
    price,
    subtotal: qty * price,
    adminFee: money(body.adminFee),
    hpp: baseHpp,
    totalHpp: qty * baseHpp,
    updatedAt: existingRow ? existingRow.updatedAt : now(),
  };
}

const excelText = (value) => {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') return '';
  return String(value).trim();
};

const excelNumber = (value) => {
  if (value === undefined || value === null || value === '') return 0;
  if (typeof value === 'number') return Math.round(value);
  const cleaned = String(value).replace(/[^\d.-]/g, '');
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? Math.round(parsed) : 0;
};

const excelDate = (value) => {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
  }
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${pad(parsed.m)}-${pad(parsed.d)}`;
    }
  }
  const text = excelText(value);
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    return text.slice(0, 10);
  }
  const shortDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+\d{1,2}:\d{2})?/);
  if (shortDate) {
    const year = shortDate[3].length === 2 ? 2000 + toInt(shortDate[3]) : toInt(shortDate[3]);
    return `${year}-${pad(toInt(shortDate[1]))}-${pad(toInt(shortDate[2]))}`;
  }
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }
  return '';
};

const normalizeHeader = (value) => excelText(value)
  .toLowerCase()
  .replace(/\s+/g, ' ')
  .trim();

const headerIndex = (headers) => {
  const map = {};
  headers.forEach((header, index) => {
    const key = normalizeHeader(header);
    if (key) map[key] = index;
  });
  return map;
};

const cellByHeader = (row, map, aliases) => {
  for (const alias of aliases) {
    const index = map[normalizeHeader(alias)];
    if (index !== undefined) {
      return row[index];
    }
  }
  return '';
};

const textByHeader = (row, map, aliases) => excelText(cellByHeader(row, map, aliases));
const numberByHeader = (row, map, aliases) => excelNumber(cellByHeader(row, map, aliases));
const dateByHeader = (row, map, aliases) => excelDate(cellByHeader(row, map, aliases));

const detectStoreNameFromFile = (fileName) => {
  const base = path.basename(fileName || '', path.extname(fileName || ''));
  const match = base.match(/^[^_]+_(.+?)_\d{4}-\d{2}-\d{2}$/);
  return match ? match[1].trim() : '';
};

const detectSaleTemplate = (headers) => {
  const keys = new Set(headers.map(normalizeHeader).filter(Boolean));
  const has = (key) => keys.has(normalizeHeader(key));
  if (has('Tipe Transaksi') && has('No Pesanan') && has('Sub Total') && has('Channel')) {
    return 'rincian-pendapatan-barang';
  }
  if (has('ID Pesanan') && has('Total Penghasilan') && has('Total Biaya')) {
    return 'penghasilan';
  }
  if (has('ID Pesanan') && has('Qty Aktual') && has('Total HPP')) {
    return 'pesanan';
  }
  if (has('Channel') && has('Nama Toko') && has('Tanggal') && has('SKU')) {
    return 'laporan-penjualan';
  }
  return '';
};

async function ensureImportedProduct(record, productHpp, summary) {
  if (!record.sku) return 0;
  const existing = await Product.findOne({ where: { sku: record.sku } });
  if (existing) {
    return existing.hpp || 0;
  }
  const hpp = productHpp.get(record.sku) || record.hpp || 0;
  await Product.create({
    sku: record.sku,
    name: record.productName || record.sku,
    variant: record.variant || '-',
    hpp,
  });
  summary.productsCreated += 1;
  return hpp;
}

async function createImportedSale(record, productHpp, summary, sheetStats, occurrenceCounts) {
  const saleDate = record.saleDate;
  const channel = record.channel;
  const storeName = record.storeName || channel;
  const orderNumber = record.orderNumber;
  const customer = record.customer || '';
  const sku = record.sku;
  const qty = Math.max(0, record.qty || 0);
  const price = record.price || (qty > 0 ? Math.round((record.subtotal || 0) / qty) : 0);
  const subtotal = record.subtotal || qty * price;
  const adminFee = Math.abs(record.adminFee || 0);

  if (!channel || !storeName || !saleDate || !sku || qty <= 0) {
    return;
  }
  if (!summary.firstSaleDate || saleDate < summary.firstSaleDate) {
    summary.firstSaleDate = saleDate;
  }
  if (!summary.lastSaleDate || saleDate > summary.lastSaleDate) {
    summary.lastSaleDate = saleDate;
  }

  const baseProductHpp = await ensureImportedProduct(record, productHpp, summary);
  const totalHpp = record.totalHpp || qty * (record.hpp || baseProductHpp || 0);
  const hpp = record.hpp || (qty > 0 ? Math.round(totalHpp / qty) : 0) || baseProductHpp || 0;
  // A single order number legitimately can repeat the exact same SKU more than once (e.g. an
  // embroidery add-on fee line per item). A plain existence check can't tell "already
  // imported before" apart from "another genuinely identical line in this order", so we
  // track how many times this order+SKU has been seen so far in this import and only skip
  // once that many rows already exist in the database.
  const duplicateWhere = orderNumber
    ? { orderNumber, sku }
    : { saleDate, channel, storeName, customer, sku, qty, price, subtotal };
  const duplicateKey = JSON.stringify(duplicateWhere);
  const occurrence = (occurrenceCounts.get(duplicateKey) || 0) + 1;
  occurrenceCounts.set(duplicateKey, occurrence);
  const existingCount = await Sale.count({ where: duplicateWhere });
  if (existingCount >= occurrence) {
    summary.salesSkipped += 1;
    sheetStats.skipped += 1;
    return;
  }

  await Sale.create({
    saleDate,
    platformId: null,
    channel,
    storeName,
    orderNumber,
    customer,
    sku,
    qty,
    price,
    subtotal,
    adminFee,
    hpp,
    totalHpp: totalHpp || qty * hpp,
  });
  summary.salesCreated += 1;
  sheetStats.created += 1;
}

async function importSalesWorkbook(buffer, fileName = '') {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true, dateNF: 'yyyy-mm-dd' });
  const summary = {
    productsCreated: 0,
    productsUpdated: 0,
    salesCreated: 0,
    salesSkipped: 0,
    firstSaleDate: '',
    lastSaleDate: '',
    sheets: [],
  };
  const occurrenceCounts = new Map();

  const productSheet = workbook.Sheets['Data Barang'];
  const productHpp = new Map();
  if (productSheet) {
    const rows = XLSX.utils.sheet_to_json(productSheet, { header: 1, raw: true, defval: null });
    for (const row of rows.slice(1)) {
      const sku = excelText(row[0]);
      const name = excelText(row[1]);
      const variant = excelText(row[2]);
      const hpp = excelNumber(row[3]);
      if (!sku || !name) continue;
      productHpp.set(sku, hpp);
      const [product, created] = await Product.findOrCreate({
        where: { sku },
        defaults: {
          sku,
          name,
          variant: variant || '-',
          hpp,
        },
      });
      if (created) {
        summary.productsCreated += 1;
      } else {
        await product.update({
          name,
          variant: variant || product.variant || '-',
          hpp,
        });
        summary.productsUpdated += 1;
      }
    }
  }

  const defaultStoreName = detectStoreNameFromFile(fileName);
  const saleSheetNames = workbook.SheetNames.filter((name) => name !== 'Data Barang');
  for (const sheetName of saleSheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    const headerRowIndex = rows.findIndex((row) => detectSaleTemplate(row));
    if (headerRowIndex < 0) {
      summary.sheets.push({
        name: sheetName,
        format: 'tidak dikenali',
        created: 0,
        skipped: 0,
      });
      continue;
    }
    const headers = rows[headerRowIndex];
    const map = headerIndex(headers);
    const format = detectSaleTemplate(headers);
    const sheetStats = { created: 0, skipped: 0 };

    for (const row of rows.slice(headerRowIndex + 1)) {
      let record = null;

      if (format === 'pesanan') {
        const qty = numberByHeader(row, map, ['Qty Aktual', 'Jumlah']);
        const totalHpp = numberByHeader(row, map, ['Total HPP']);
        record = {
          saleDate: dateByHeader(row, map, ['Tanggal']),
          channel: textByHeader(row, map, ['Platform']) || 'Shopee',
          storeName: defaultStoreName || textByHeader(row, map, ['Platform']) || 'Shopee',
          orderNumber: textByHeader(row, map, ['ID Pesanan']),
          customer: textByHeader(row, map, ['Pembeli']),
          productName: textByHeader(row, map, ['Produk']),
          variant: textByHeader(row, map, ['Varian']),
          sku: textByHeader(row, map, ['SKU']),
          qty,
          price: numberByHeader(row, map, ['Harga Satuan']),
          subtotal: numberByHeader(row, map, ['Total Harga Jual', 'Total Harga Produk']),
          adminFee: 0,
          hpp: qty > 0 ? Math.round(totalHpp / qty) : 0,
          totalHpp,
        };
      } else if (format === 'penghasilan') {
        const qty = numberByHeader(row, map, ['Jumlah']);
        const totalFee = numberByHeader(row, map, ['Total Biaya']) || numberByHeader(row, map, ['Biaya Admin']);
        record = {
          saleDate: dateByHeader(row, map, ['Tanggal']),
          channel: 'Shopee',
          storeName: defaultStoreName || 'Shopee',
          orderNumber: textByHeader(row, map, ['ID Pesanan']),
          customer: '',
          productName: textByHeader(row, map, ['Produk']),
          variant: textByHeader(row, map, ['Varian']),
          sku: textByHeader(row, map, ['SKU']),
          qty,
          price: qty > 0 ? Math.round(numberByHeader(row, map, ['Total Harga Produk']) / qty) : 0,
          subtotal: numberByHeader(row, map, ['Total Harga Jual', 'Total Harga Produk']),
          adminFee: totalFee,
          hpp: productHpp.get(textByHeader(row, map, ['SKU'])) || 0,
          totalHpp: 0,
        };
      } else if (format === 'rincian-pendapatan-barang') {
        const store = textByHeader(row, map, ['Nama Toko']);
        const channel = textByHeader(row, map, ['Channel']);
        record = {
          saleDate: dateByHeader(row, map, ['Tanggal']),
          channel,
          storeName: store || channel,
          orderNumber: textByHeader(row, map, ['No Pesanan']),
          customer: textByHeader(row, map, ['Pelanggan']),
          productName: textByHeader(row, map, ['Nama Barang']),
          variant: '',
          sku: textByHeader(row, map, ['SKU']),
          qty: numberByHeader(row, map, ['QTY']),
          price: numberByHeader(row, map, ['Harga']),
          subtotal: numberByHeader(row, map, ['Sub Total']),
          adminFee: numberByHeader(row, map, ['Potongan Biaya', 'Biaya Lainnya']),
          hpp: productHpp.get(textByHeader(row, map, ['SKU'])) || 0,
          totalHpp: 0,
        };
      } else {
        record = {
          saleDate: dateByHeader(row, map, ['Tanggal']),
          channel: textByHeader(row, map, ['Channel']),
          storeName: textByHeader(row, map, ['Nama Toko']),
          orderNumber: textByHeader(row, map, ['No Pesanan']),
          customer: textByHeader(row, map, ['Pelanggan']),
          productName: textByHeader(row, map, ['Nama Barang']),
          variant: '',
          sku: textByHeader(row, map, ['SKU']),
          qty: numberByHeader(row, map, ['QTY', 'Jumlah']),
          price: numberByHeader(row, map, ['Harga']),
          subtotal: numberByHeader(row, map, ['Sub Total', 'Subtotal']),
          adminFee: numberByHeader(row, map, ['Biaya Admin']),
          hpp: numberByHeader(row, map, ['HPP', 'Hpp']) || productHpp.get(textByHeader(row, map, ['SKU'])) || 0,
          totalHpp: numberByHeader(row, map, ['Total Hpp', 'Total HPP']),
        };
      }

      if (!record || !record.sku || !record.saleDate) {
        continue;
      }
      await createImportedSale(record, productHpp, summary, sheetStats, occurrenceCounts);
    }

    summary.sheets.push({
      name: sheetName,
      format,
      created: sheetStats.created,
      skipped: sheetStats.skipped,
    });
  }

  return summary;
}

const SALES_TEMPLATE_NAME = 'template_penjualan.xlsx';

function buildSalesTemplateBuffer() {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['Channel', 'Nama Toko', 'Tanggal', 'No Pesanan', 'Pelanggan', 'SKU', 'Jumlah', 'Harga', 'Subtotal', 'Biaya Admin', 'Hpp', 'Total Hpp'],
    ['POS', 'Contoh Toko', '2026-01-01', 'SO-000001', 'Contoh Pelanggan', 'TPD-CONTOH-001', 1, 65000, 65000, 0, 33000, 33000],
  ];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Penjualan');
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

const PRODUCT_TEMPLATE_NAME = 'template_data_barang.xlsx';

function buildProductTemplateBuffer() {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['SKU', 'Nama Barang', 'Jenis', 'HPP'],
    ['TPD-CONTOH-001', 'Topi Contoh', 'Model A', 50000],
  ];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data Barang');
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

const PRODUCT_SKU_ALIASES = ['SKU', 'ID SKU BARANG', 'ID SKU', 'Kode Barang'];
const PRODUCT_VARIANT_ALIASES = ['Jenis', 'Varian', 'Variasi'];

const detectProductTemplate = (headers) => {
  const keys = new Set(headers.map(normalizeHeader).filter(Boolean));
  const has = (key) => keys.has(normalizeHeader(key));
  const hasAny = (aliases) => aliases.some(has);
  return hasAny(PRODUCT_SKU_ALIASES) && has('Nama Barang') && hasAny(PRODUCT_VARIANT_ALIASES) && has('HPP');
};

const detectMarketingTemplate = (headers) => {
  const keys = new Set(headers.map(normalizeHeader).filter(Boolean));
  const has = (key) => keys.has(normalizeHeader(key));
  return has('Tanggal') && has('Kategori') && has('Platform & Toko') && has('Keterangan') && has('Nominal');
};

const excelPercent = (value) => {
  if (value === undefined || value === null || value === '') return 0;
  if (typeof value === 'number') {
    return value > 1 ? value / 100 : value;
  }
  const text = String(value).trim().replace(',', '.');
  if (!text) return 0;
  if (text.endsWith('%')) {
    const parsed = Number.parseFloat(text.slice(0, -1));
    return Number.isFinite(parsed) ? parsed / 100 : 0;
  }
  const parsed = Number.parseFloat(text);
  if (!Number.isFinite(parsed)) return 0;
  return parsed > 1 ? parsed / 100 : parsed;
};

async function importProductsWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true, dateNF: 'yyyy-mm-dd' });
  const sheetName = workbook.SheetNames.find((name) => normalizeHeader(name) === normalizeHeader('Data Barang')) || workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('Sheet Data Barang tidak ditemukan.');
  }
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const headerRowIndex = rows.findIndex((row) => detectProductTemplate(row));
  if (headerRowIndex < 0) {
    throw new Error('Format Data Barang tidak dikenali. Gunakan template yang tersedia.');
  }

  const map = headerIndex(rows[headerRowIndex]);
  const summary = {
    productsCreated: 0,
    productsUpdated: 0,
    sheets: [],
  };
  let created = 0;
  let updated = 0;

  for (const row of rows.slice(headerRowIndex + 1)) {
    const sku = textByHeader(row, map, PRODUCT_SKU_ALIASES);
    const name = textByHeader(row, map, ['Nama Barang']);
    const variant = textByHeader(row, map, PRODUCT_VARIANT_ALIASES);
    const hpp = numberByHeader(row, map, ['HPP']);
    if (!sku || !name) {
      continue;
    }

    const [product, isCreated] = await Product.findOrCreate({
      where: { sku },
      defaults: {
        sku,
        name,
        variant: variant || '-',
        hpp,
      },
    });
    if (isCreated) {
      created += 1;
      summary.productsCreated += 1;
      continue;
    }
    await product.update({
      name,
      variant: variant || product.variant || '-',
      hpp,
    });
    updated += 1;
    summary.productsUpdated += 1;
  }

  summary.sheets.push({
    name: sheetName,
    created,
    updated,
  });

  return summary;
}

function buildMarketingTemplateBuffer() {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['Tanggal', 'Kategori', 'Platform & Toko', 'Keterangan', 'Nominal', 'Pajak Iklan', 'Total Pajak'],
    ['1 Januari 2026', 'Iklan Shopee', 'Shopee Mimimau Baby & Kids', '', 'Rp128.700', '11,00%', 'Rp14.157'],
  ];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Beban Marketing');
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

async function importMarketingWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true, dateNF: 'yyyy-mm-dd' });
  let sheetName = null;
  let rows = null;
  let headerRowIndex = -1;

  for (const candidate of workbook.SheetNames) {
    const candidateRows = XLSX.utils.sheet_to_json(workbook.Sheets[candidate], { header: 1, raw: true, defval: null });
    const candidateHeaderIndex = candidateRows.findIndex((row) => detectMarketingTemplate(row || []));
    if (candidateHeaderIndex >= 0) {
      sheetName = candidate;
      rows = candidateRows;
      headerRowIndex = candidateHeaderIndex;
      break;
    }
  }

  if (!sheetName || !rows) {
    throw new Error('Format beban marketing tidak dikenali. Gunakan template yang tersedia.');
  }
  if (headerRowIndex < 0) {
    throw new Error('Header beban marketing tidak ditemukan.');
  }
  const map = headerIndex(rows[headerRowIndex]);
  let created = 0;
  let updated = 0;
  const summary = {
    marketingCreated: 0,
    marketingUpdated: 0,
    sheets: [],
  };

  for (const row of rows.slice(headerRowIndex + 1)) {
    const expenseDate = dateByHeader(row, map, ['Tanggal']);
    const category = textByHeader(row, map, ['Kategori']);
    const platformStore = textByHeader(row, map, ['Platform & Toko']);
    const description = textByHeader(row, map, ['Keterangan']);
    const nominal = numberByHeader(row, map, ['Nominal']);
    const taxRate = excelPercent(cellByHeader(row, map, ['Pajak Iklan']));
    const totalTax = numberByHeader(row, map, ['Total Pajak']) || Math.round(nominal * taxRate);

    if (!expenseDate || !category || !platformStore) {
      continue;
    }

    const duplicate = await MarketingExpense.findOne({
      where: {
        expenseDate,
        category,
        platformStore,
        description: description || '',
        nominal,
      },
    });
    if (duplicate) {
      await duplicate.update({
        description: description || duplicate.description,
        nominal,
        taxRate,
        totalTax,
      });
      updated += 1;
      summary.marketingUpdated += 1;
      continue;
    }

    await MarketingExpense.create({
      expenseDate,
      category,
      platformStore,
      description: description || '-',
      nominal,
      taxRate,
      totalTax,
      platformId: null,
    });
    created += 1;
    summary.marketingCreated += 1;
  }

  summary.sheets.push({
    name: sheetName,
    created,
    updated,
  });

  return summary;
}

async function resetDatabase() {
  const tables = [
    'attendance',
    'attendance_settings',
    'salary_payments',
    'shifts',
    'marketing_expenses',
    'operational_expenses',
    'sales',
    'stores',
    'products',
    'users',
    'platforms',
  ];
  await sequelize.query('SET FOREIGN_KEY_CHECKS = 0');
  try {
    for (const tableName of tables) {
      await sequelize.query(`TRUNCATE TABLE \`${tableName}\``);
    }
  } finally {
    await sequelize.query('SET FOREIGN_KEY_CHECKS = 1');
  }
  await createInitialAdmin();
  await getAttendanceSettings();
}

createSimpleCrudRoutes('shifts', Shift, serializeShift, parseShift, {});
createSimpleCrudRoutes('stores', Store, serializeStore, parseStore, { include: [Platform], reload: true });
createSimpleCrudRoutes('products', Product, serializeProduct, parseProduct, {
  paginated: true,
  searchFields: ['sku', 'name', 'variant'],
});
app.get('/api/products/template', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const buffer = buildProductTemplateBuffer();
    res.setHeader('Content-Disposition', `attachment; filename="${PRODUCT_TEMPLATE_NAME}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal membuat template barang.' });
  }
});

app.post('/api/products/import-excel', ...authRequired, ownerOnly, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'File Excel wajib diunggah.' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) {
      return res.status(400).json({ ok: false, message: 'Format file harus .xlsx atau .xls.' });
    }
    const summary = await importProductsWorkbook(req.file.buffer);
    return res.json({
      ok: true,
      message: 'Import barang berhasil diproses.',
      data: summary,
    });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal import data barang.' });
  }
});

app.get('/api/sales/template', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const buffer = buildSalesTemplateBuffer();
    res.setHeader('Content-Disposition', `attachment; filename="${SALES_TEMPLATE_NAME}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal membuat template penjualan.' });
  }
});

const SALE_GROUP_KEY_SQL = "CASE WHEN `order_number` IS NULL OR `order_number` = '' THEN CONCAT('__row_', `id`) ELSE `order_number` END";

async function fetchGroupedSales({ search, storeName, from, to, limit, offset }) {
  const conditions = [];
  const replacements = {};
  if (search) {
    conditions.push('(`order_number` LIKE :search OR `customer` LIKE :search OR `sku` LIKE :search OR `store_name` LIKE :search OR `channel` LIKE :search)');
    replacements.search = `%${search}%`;
  }
  if (storeName) {
    conditions.push('`store_name` = :storeName');
    replacements.storeName = storeName;
  }
  if (from) {
    conditions.push('`sale_date` >= :from');
    replacements.from = from;
  }
  if (to) {
    conditions.push('`sale_date` <= :to');
    replacements.to = to;
  }
  const whereSql = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

  const countRows = await sequelize.query(
    `SELECT COUNT(*) AS total FROM (SELECT 1 FROM \`sales\` ${whereSql} GROUP BY ${SALE_GROUP_KEY_SQL}) t`,
    { replacements, type: sequelize.QueryTypes.SELECT },
  );
  const total = toInt(countRows[0]?.total, 0);

  const rows = await sequelize.query(
    `SELECT
       MIN(\`id\`) AS id,
       MIN(\`sale_date\`) AS saleDate,
       MIN(\`channel\`) AS channel,
       MIN(\`store_name\`) AS storeName,
       CASE WHEN MIN(\`order_number\`) IS NULL OR MIN(\`order_number\`) = '' THEN '' ELSE MIN(\`order_number\`) END AS orderNumber,
       MIN(\`customer\`) AS customer,
       COUNT(*) AS itemCount,
       SUM(\`qty\`) AS totalQty,
       SUM(\`subtotal\`) AS totalSubtotal,
       SUM(\`total_hpp\`) AS totalHpp
     FROM \`sales\`
     ${whereSql}
     GROUP BY ${SALE_GROUP_KEY_SQL}
     ORDER BY saleDate DESC, id DESC
     LIMIT :limit OFFSET :offset`,
    { replacements: { ...replacements, limit, offset }, type: sequelize.QueryTypes.SELECT },
  );

  return { rows, total };
}

const serializeGroupedSale = (row, { includeHpp }) => ({
  id: row.id,
  saleDate: row.saleDate,
  channel: row.channel,
  storeName: row.storeName,
  orderNumber: row.orderNumber || '',
  customer: row.customer || '',
  itemCount: toInt(row.itemCount, 0),
  totalQty: toInt(row.totalQty, 0),
  totalSubtotal: toInt(row.totalSubtotal, 0),
  ...(includeHpp ? { totalHpp: toInt(row.totalHpp, 0) } : {}),
});

app.get('/api/sales', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const includeHpp = req.user.role === 'owner';
    const { page, limit, offset } = getPagination(req.query);
    const searchText = toText(req.query.search);
    const storeName = toText(req.query.storeName);
    const from = toText(req.query.from);
    const to = toText(req.query.to);

    if (toText(req.query.grouped) === 'true') {
      const { rows, total } = await fetchGroupedSales({ search: searchText, storeName, from, to, limit, offset });
      return res.json({
        ok: true,
        data: rows.map((row) => serializeGroupedSale(row, { includeHpp })),
        pagination: paginationMeta(page, limit, total),
      });
    }

    const where = { ...searchWhere(searchText, ['orderNumber', 'customer', 'sku', 'storeName', 'channel']) };
    if (storeName) {
      where.storeName = storeName;
    }
    if (from || to) {
      where.saleDate = {};
      if (from) where.saleDate[Op.gte] = from;
      if (to) where.saleDate[Op.lte] = to;
    }

    const { count, rows } = await Sale.findAndCountAll({
      where,
      include: [Platform],
      order: [['saleDate', 'DESC'], ['createdAt', 'DESC']],
      limit,
      offset,
    });
    return res.json({
      ok: true,
      data: rows.map((row) => serializeSale(row, { includeHpp })),
      pagination: paginationMeta(page, limit, count),
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat penjualan.' });
  }
});

app.get('/api/sales/detail', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const includeHpp = req.user.role === 'owner';
    const orderNumber = toText(req.query.orderNumber);
    const id = toInt(req.query.id, null);
    if (!orderNumber && !id) {
      return res.status(400).json({ ok: false, message: 'Nomor pesanan atau id wajib diisi.' });
    }
    const where = orderNumber ? { orderNumber } : { id };
    const rows = await Sale.findAll({ where, include: [Platform], order: [['id', 'ASC']] });
    return res.json({ ok: true, data: rows.map((row) => serializeSale(row, { includeHpp })) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat detail penjualan.' });
  }
});

app.post('/api/sales', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const payload = await parseSale(req.body, null);
    const row = await Sale.create(payload);
    const output = await Sale.findByPk(row.id, { include: [Platform] });
    return res.status(201).json({ ok: true, data: serializeSale(output, { includeHpp: req.user.role === 'owner' }) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan penjualan.' });
  }
});

app.post('/api/sales/import-excel', ...authRequired, salesRecapAccess, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'File Excel wajib diunggah.' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) {
      return res.status(400).json({ ok: false, message: 'Format file harus .xlsx atau .xls.' });
    }
    const summary = await importSalesWorkbook(req.file.buffer, req.file.originalname);
    return res.json({
      ok: true,
      message: 'Import penjualan berhasil diproses.',
      data: summary,
    });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal import file Excel.' });
  }
});

app.put('/api/sales/:id', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const row = await Sale.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    const payload = await parseSale(req.body, row);
    await row.update(payload);
    const output = await Sale.findByPk(row.id, { include: [Platform] });
    return res.json({ ok: true, data: serializeSale(output, { includeHpp: req.user.role === 'owner' }) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal memperbarui penjualan.' });
  }
});

app.delete('/api/sales/:id', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const row = await Sale.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Penjualan berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus penjualan.' });
  }
});
createSimpleCrudRoutes('expenses/operasional', OperationalExpense, serializeOperational, parseOperational, { writeMiddleware: ownerOrLeader });
createSimpleCrudRoutes('expenses/marketing', MarketingExpense, serializeMarketing, parseMarketing, { include: [Platform], reload: true, writeMiddleware: ownerOrLeader });

app.get('/api/expense-categories', ...authRequired, async (req, res) => {
  try {
    const kind = toText(req.query.kind);
    const where = kind ? { kind } : {};
    const rows = await ExpenseCategory.findAll({ where, order: [['name', 'ASC']] });
    return res.json({ ok: true, data: rows.map(serializeExpenseCategory) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat kategori.' });
  }
});

app.post('/api/expense-categories', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const name = toText(req.body.name);
    const kind = toText(req.body.kind);
    if (!name || !['operasional', 'marketing'].includes(kind)) {
      return res.status(400).json({ ok: false, message: 'Nama dan jenis kategori wajib diisi.' });
    }
    const [row] = await ExpenseCategory.findOrCreate({ where: { name, kind }, defaults: { name, kind } });
    return res.status(201).json({ ok: true, data: serializeExpenseCategory(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah kategori.' });
  }
});

app.get('/api/expenses/marketing/template', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const buffer = buildMarketingTemplateBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="template_beban_marketing.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal membuat template marketing.' });
  }
});

app.post('/api/expenses/marketing/import-excel', ...authRequired, ownerOrLeader, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'File Excel wajib diunggah.' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) {
      return res.status(400).json({ ok: false, message: 'Format file harus .xlsx atau .xls.' });
    }
    const summary = await importMarketingWorkbook(req.file.buffer);
    return res.json({
      ok: true,
      message: 'Import beban marketing berhasil diproses.',
      data: summary,
    });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal import beban marketing.' });
  }
});

app.post('/api/admin/reset-database', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const password = toText(req.body.password);
    if (!password) {
      return res.status(400).json({ ok: false, message: 'Password admin wajib diisi.' });
    }
    const valid = await bcrypt.compare(password, req.user.passwordHash);
    if (!valid) {
      return res.status(401).json({ ok: false, message: 'Password admin tidak sesuai.' });
    }
    await resetDatabase();
    return res.json({
      ok: true,
      message: 'Database berhasil direset. Silakan masuk kembali dengan akun admin awal.',
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal mereset database.' });
  }
});

app.get('/api/auth/me', ...authRequired, async (req, res) => {
  return res.json({ ok: true, data: serializeUser(req.user) });
});

app.post('/api/auth/login', async (req, res) => {
  try {
    const username = toText(req.body.username);
    const password = toText(req.body.password);
    if (!username || !password) {
      return res.status(400).json({ ok: false, message: 'Username dan password wajib diisi.' });
    }
    const user = await User.findOne({ where: { username }, include: [Shift] });
    if (!user || !user.active) {
      return res.status(401).json({ ok: false, message: 'Username atau password salah.' });
    }
    const valid = await bcrypt.compare(password, user.passwordHash);
    if (!valid) {
      return res.status(401).json({ ok: false, message: 'Username atau password salah.' });
    }
    user.lastLoginAt = now();
    await user.save();
    const token = jwt.sign({ id: user.id }, JWT_SECRET, { expiresIn: '7d' });
    return res.json({
      ok: true,
      token,
      data: serializeUser(user),
      message: 'Masuk berhasil.',
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal masuk ke sistem.' });
  }
});

app.get('/api/users', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const rows = await User.findAll({ order: [['createdAt', 'DESC']], include: [Shift] });
    return res.json({ ok: true, data: rows.map(serializeUser) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat karyawan.' });
  }
});

app.post('/api/users', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const username = toText(req.body.username);
    const password = toText(req.body.password);
    const fullName = toText(req.body.fullName);
    if (!username || !fullName) {
      return res.status(400).json({ ok: false, message: 'Username dan nama lengkap wajib diisi.' });
    }
    const exists = await User.findOne({ where: { username } });
    if (exists) {
      return res.status(409).json({ ok: false, message: 'Username sudah dipakai.' });
    }
    const passwordHash = await bcrypt.hash(password || '12345678', 10);
    const row = await User.create({
      username,
      passwordHash,
      fullName,
      role: ROLES.includes(req.body.role) ? req.body.role : 'karyawan',
      jobTitle: toText(req.body.jobTitle),
      workShift: toText(req.body.workShift),
      phone: toText(req.body.phone),
      dailyWage: money(req.body.dailyWage),
      mealAllowance: money(req.body.mealAllowance),
      overtimeRatePerHour: money(req.body.overtimeRatePerHour),
      shiftId: req.body.shiftId ? toInt(req.body.shiftId, null) : null,
      active: req.body.active === undefined ? true : Boolean(req.body.active),
    });
    const output = await User.findByPk(row.id, { include: [Shift] });
    return res.status(201).json({ ok: true, data: serializeUser(output) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah karyawan.' });
  }
});

app.put('/api/users/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await User.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    if (req.body.username) {
      const duplicate = await User.findOne({
        where: {
          username: toText(req.body.username),
          id: { [Op.ne]: row.id },
        },
      });
      if (duplicate) {
        return res.status(409).json({ ok: false, message: 'Username sudah dipakai.' });
      }
      row.username = toText(req.body.username);
    }
    if (req.body.password) {
      row.passwordHash = await bcrypt.hash(toText(req.body.password), 10);
    }
    row.fullName = toText(req.body.fullName || row.fullName);
    row.role = ROLES.includes(req.body.role) ? req.body.role : row.role;
    row.jobTitle = toText(req.body.jobTitle);
    row.workShift = toText(req.body.workShift);
    row.phone = toText(req.body.phone);
    row.dailyWage = money(req.body.dailyWage);
    row.mealAllowance = money(req.body.mealAllowance);
    row.overtimeRatePerHour = money(req.body.overtimeRatePerHour);
    row.shiftId = req.body.shiftId ? toInt(req.body.shiftId, null) : null;
    if (req.body.active !== undefined) {
      row.active = Boolean(req.body.active);
    }
    await row.save();
    const output = await User.findByPk(row.id, { include: [Shift] });
    return res.json({ ok: true, data: serializeUser(output) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal memperbarui karyawan.' });
  }
});

app.delete('/api/users/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    if (Number(req.params.id) === req.user.id) {
      return res.status(400).json({ ok: false, message: 'Akun sendiri tidak bisa dihapus.' });
    }
    const row = await User.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Karyawan berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus karyawan.' });
  }
});

app.get('/api/face/me', ...authRequired, async (req, res) => {
  if (req.user.role !== 'karyawan') {
    return res.status(403).json({ ok: false, message: 'Daftar wajah hanya tersedia untuk karyawan.' });
  }
  return res.json({
    ok: true,
    data: {
      registered: Boolean(req.user.faceDescriptor || req.user.faceImageUrl),
      faceDescriptor: req.user.faceDescriptor || null,
      faceImageFilename: req.user.faceImageFilename || '',
      faceImageUrl: req.user.faceImageUrl || '',
      faceImageSize: req.user.faceImageSize || 0,
      faceRegisteredAt: req.user.faceRegisteredAt,
    },
  });
});

app.post('/api/face/register', ...authRequired, async (req, res) => {
  try {
    if (req.user.role !== 'karyawan') {
      return res.status(403).json({ ok: false, message: 'Daftar wajah hanya untuk karyawan.' });
    }
    const descriptor = req.body.descriptor;
    const faceImageUrl = toText(req.body.faceImageUrl);
    if (!faceImageUrl && (!Array.isArray(descriptor) || descriptor.length < 64)) {
      return res.status(400).json({ ok: false, message: 'Foto wajah wajib diunggah.' });
    }
    req.user.faceDescriptor = Array.isArray(descriptor) && descriptor.length >= 64 ? descriptor : null;
    req.user.faceImageFilename = toText(req.body.faceImageFilename);
    req.user.faceImageUrl = faceImageUrl;
    req.user.faceImageSize = req.body.faceImageSize ? money(req.body.faceImageSize) : null;
    req.user.faceRegisteredAt = now();
    await req.user.save();
    return res.json({
      ok: true,
      message: 'Wajah berhasil didaftarkan.',
      data: {
        registered: true,
        faceImageFilename: req.user.faceImageFilename || '',
        faceImageUrl: req.user.faceImageUrl || '',
        faceImageSize: req.user.faceImageSize || 0,
        faceRegisteredAt: req.user.faceRegisteredAt,
      },
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menyimpan data wajah.' });
  }
});

app.get('/api/attendance/me', ...authRequired, async (req, res) => {
  try {
    const rows = await Attendance.findAll({
      where: { userId: req.user.id },
      include: [{ model: User }],
      order: [['workDate', 'DESC'], ['createdAt', 'DESC']],
      limit: 120,
    });
    return res.json({ ok: true, data: rows.map(serializeAttendance) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat absensi.' });
  }
});

app.get('/api/attendance', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const rows = await Attendance.findAll({
      include: [{ model: User }],
      order: [['workDate', 'DESC'], ['createdAt', 'DESC']],
      limit: 400,
    });
    return res.json({ ok: true, data: rows.map(serializeAttendance) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat absensi.' });
  }
});

app.get('/api/attendance/settings', ...authRequired, async (req, res) => {
  try {
    const settings = await getAttendanceSettings();
    return res.json({ ok: true, data: serializeAttendanceSettings(settings) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat pengaturan absensi.' });
  }
});

app.put('/api/attendance/settings', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const checkInStart = toText(req.body.checkInStart);
    const checkInDeadline = toText(req.body.checkInDeadline);
    const checkOutStart = toText(req.body.checkOutStart);
    for (const [label, value] of [['Jam mulai masuk', checkInStart], ['Batas telat', checkInDeadline], ['Jam mulai pulang', checkOutStart]]) {
      if (!TIME_RE.test(value)) {
        return res.status(400).json({ ok: false, message: `${label} harus berformat HH:mm.` });
      }
    }
    const lateToleranceMinutes = toInt(req.body.lateToleranceMinutes, 0);
    if (lateToleranceMinutes < 0 || lateToleranceMinutes > 120) {
      return res.status(400).json({ ok: false, message: 'Toleransi telat harus antara 0-120 menit.' });
    }
    const radiusEnabled = Boolean(req.body.radiusEnabled);
    const radiusMeters = toInt(req.body.radiusMeters, 100);
    if (radiusMeters < 10 || radiusMeters > 5000) {
      return res.status(400).json({ ok: false, message: 'Radius harus antara 10-5000 meter.' });
    }
    const paydayWeekday = toText(req.body.paydayWeekday);
    if (paydayWeekday && !PAYDAY_WEEKDAYS.includes(paydayWeekday)) {
      return res.status(400).json({ ok: false, message: 'Hari gajian tidak valid.' });
    }
    const hasLatitude = req.body.latitude !== undefined && req.body.latitude !== null && req.body.latitude !== '';
    const hasLongitude = req.body.longitude !== undefined && req.body.longitude !== null && req.body.longitude !== '';
    const latitude = hasLatitude ? toFloat(req.body.latitude, null) : null;
    const longitude = hasLongitude ? toFloat(req.body.longitude, null) : null;
    const settings = await getAttendanceSettings();
    const willHaveCoordinates = hasLatitude && hasLongitude ? true : (settings.latitude !== null && settings.longitude !== null);
    if (radiusEnabled && !willHaveCoordinates) {
      return res.status(400).json({ ok: false, message: 'Set lokasi toko dengan "Gunakan Lokasi Saat Ini" sebelum mengaktifkan radius.' });
    }
    settings.checkInStart = checkInStart;
    settings.checkInDeadline = checkInDeadline;
    settings.checkOutStart = checkOutStart;
    settings.lateToleranceMinutes = lateToleranceMinutes;
    settings.radiusEnabled = radiusEnabled;
    settings.radiusMeters = radiusMeters;
    settings.locationLabel = toText(req.body.locationLabel);
    settings.paydayWeekday = paydayWeekday || null;
    if (hasLatitude) settings.latitude = latitude;
    if (hasLongitude) settings.longitude = longitude;
    await settings.save();
    return res.json({ ok: true, message: 'Pengaturan absensi tersimpan.', data: serializeAttendanceSettings(settings) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menyimpan pengaturan absensi.' });
  }
});

app.get('/api/holidays', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const rows = await Holiday.findAll({ order: [['date', 'DESC']] });
    return res.json({ ok: true, data: rows.map(serializeHoliday) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat hari libur.' });
  }
});

app.post('/api/holidays', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const date = toText(req.body.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return res.status(400).json({ ok: false, message: 'Tanggal libur wajib diisi.' });
    }
    const [row] = await Holiday.findOrCreate({ where: { date }, defaults: { date, label: toText(req.body.label) } });
    return res.status(201).json({ ok: true, data: serializeHoliday(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah hari libur.' });
  }
});

app.delete('/api/holidays/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await Holiday.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Hari libur berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus hari libur.' });
  }
});

const ATTENDANCE_STATUSES = ['hadir', 'izin', 'sakit', 'cuti'];

app.get('/api/employees/roster', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const rows = await User.findAll({
      where: { active: true },
      attributes: ['id', 'fullName'],
      order: [['fullName', 'ASC']],
    });
    return res.json({ ok: true, data: rows.map((row) => ({ id: row.id, fullName: row.fullName })) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat daftar karyawan.' });
  }
});

app.post('/api/attendance', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const userId = toInt(req.body.userId, null);
    const workDate = toText(req.body.workDate);
    const status = toText(req.body.status) || 'hadir';
    if (!userId || !workDate) {
      return res.status(400).json({ ok: false, message: 'Karyawan dan tanggal wajib diisi.' });
    }
    if (!ATTENDANCE_STATUSES.includes(status)) {
      return res.status(400).json({ ok: false, message: 'Status tidak valid.' });
    }
    const target = await User.findByPk(userId);
    if (!target) {
      return res.status(404).json({ ok: false, message: 'Karyawan tidak ditemukan.' });
    }
    let row = await Attendance.findOne({ where: { userId, workDate } });
    if (row) {
      row.status = status;
      row.note = toText(req.body.note) || row.note;
      await row.save();
    } else {
      row = await Attendance.create({
        userId,
        workDate,
        status,
        method: 'manual',
        note: toText(req.body.note),
      });
    }
    const output = await Attendance.findByPk(row.id, { include: [{ model: User }] });
    return res.status(201).json({ ok: true, message: 'Absensi tersimpan.', data: serializeAttendance(output) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah absensi.' });
  }
});

app.put('/api/attendance/:id', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const row = await Attendance.findByPk(req.params.id, { include: [{ model: User }] });
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data absensi tidak ditemukan.' });
    }
    if (req.body.status !== undefined) {
      const status = toText(req.body.status);
      if (!ATTENDANCE_STATUSES.includes(status)) {
        return res.status(400).json({ ok: false, message: 'Status tidak valid.' });
      }
      row.status = status;
    }
    if (req.body.overtimeHours !== undefined) {
      row.overtimeHours = Math.max(0, toFloat(req.body.overtimeHours, 0));
    }
    if (req.body.note !== undefined) {
      row.note = toText(req.body.note);
    }
    await row.save();
    return res.json({ ok: true, message: 'Absensi tersimpan.', data: serializeAttendance(row) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menyimpan absensi.' });
  }
});

app.delete('/api/attendance/:id', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const row = await Attendance.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Data absensi berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus data absensi.' });
  }
});

app.post('/api/attendance/check-in', ...authRequired, async (req, res) => {
  try {
    if (req.user.role !== 'karyawan') {
      return res.status(403).json({ ok: false, message: 'Absensi wajah hanya untuk karyawan.' });
    }
    const photoUrl = toText(req.body.photoUrl);
    if (!photoUrl) {
      return res.status(400).json({ ok: false, message: 'Foto absensi wajib diambil dari kamera terlebih dahulu.' });
    }
    const locationName = toText(req.body.locationName);
    const settings = await getAttendanceSettings();
    const shiftWindow = resolveShiftWindow(req.user, settings);
    const today = jakartaDate();
    let row = await Attendance.findOne({
      where: {
        userId: req.user.id,
        workDate: today,
      },
      include: [{ model: User }],
    });

    let distanceMeters = null;
    if (settings.radiusEnabled && settings.latitude !== null && settings.longitude !== null) {
      const lat = toFloat(req.body.latitude, null);
      const lng = toFloat(req.body.longitude, null);
      if (lat === null || lng === null) {
        return res.status(400).json({ ok: false, message: 'Lokasi tidak terdeteksi. Aktifkan izin lokasi lalu coba lagi.' });
      }
      distanceMeters = haversineMeters(lat, lng, settings.latitude, settings.longitude);
      if (distanceMeters > settings.radiusMeters) {
        return res.status(422).json({
          ok: false,
          message: `Anda berada di luar radius lokasi absensi (±${Math.round(distanceMeters)}m dari ${settings.locationLabel || 'lokasi toko'}, radius diizinkan ${settings.radiusMeters}m).`,
        });
      }
    }

    if (!row) {
      row = await Attendance.create({
        userId: req.user.id,
        workDate: today,
        checkInAt: now(),
        status: 'hadir',
        method: 'wajah',
        confidence: req.body.confidence === undefined ? null : toFloat(req.body.confidence, null),
        note: toText(req.body.note),
        isLate: isCheckInLate(shiftWindow),
        checkInDistanceMeters: distanceMeters,
        checkInPhotoUrl: photoUrl,
        checkInLocationName: locationName || null,
        shiftName: shiftWindow.name,
      });
      row = await Attendance.findByPk(row.id, { include: [{ model: User }] });
      return res.json({ ok: true, message: row.isLate ? 'Absensi masuk berhasil dicatat (Telat).' : 'Absensi masuk berhasil dicatat.', data: serializeAttendance(row) });
    }
    if (!row.checkInAt) {
      row.checkInAt = now();
      row.status = 'hadir';
      row.method = 'wajah';
      row.confidence = req.body.confidence === undefined ? null : toFloat(req.body.confidence, null);
      row.note = toText(req.body.note);
      row.isLate = isCheckInLate(shiftWindow);
      row.checkInDistanceMeters = distanceMeters;
      row.checkInPhotoUrl = photoUrl;
      row.checkInLocationName = locationName || null;
      row.shiftName = shiftWindow.name;
      await row.save();
      row = await Attendance.findByPk(row.id, { include: [{ model: User }] });
      return res.json({ ok: true, message: row.isLate ? 'Absensi masuk berhasil dicatat (Telat).' : 'Absensi masuk berhasil dicatat.', data: serializeAttendance(row) });
    }
    if (!row.checkOutAt) {
      row.checkOutAt = now();
      row.checkOutDistanceMeters = distanceMeters;
      row.checkOutPhotoUrl = photoUrl;
      row.checkOutLocationName = locationName || null;
      if (row.checkInAt && await isOvertimeDay(row.workDate)) {
        row.overtimeHours = Math.max(0, (row.checkOutAt.getTime() - row.checkInAt.getTime()) / 3600000);
      }
      await row.save();
      row = await Attendance.findByPk(row.id, { include: [{ model: User }] });
      return res.json({ ok: true, message: 'Absensi pulang berhasil dicatat.', data: serializeAttendance(row) });
    }
    return res.status(409).json({ ok: false, message: 'Absensi hari ini sudah lengkap.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal mencatat absensi.' });
  }
});

function salaryScopeWhere(requester) {
  if (requester.role === 'admin') {
    return null;
  }
  if (requester.role === 'karyawan') {
    return { id: requester.id, active: true };
  }
  if (requester.role === 'leader') {
    return { role: { [Op.ne]: 'owner' }, active: true };
  }
  return { active: true };
}

async function computeSalaryReport(weekStart, scopeWhere) {
  const { start, end } = weekBounds(weekStart);
  const users = await User.findAll({ where: scopeWhere, order: [['fullName', 'ASC']] });
  const attendanceRows = await Attendance.findAll({
    where: { workDate: { [Op.between]: [start, end] } },
    include: [{ model: User }],
  });
  const grouped = new Map();
  for (const row of attendanceRows) {
    const key = row.userId;
    if (!grouped.has(key)) {
      grouped.set(key, []);
    }
    grouped.get(key).push(row);
  }
  const payments = await SalaryPayment.findAll({
    where: { weekStart, userId: { [Op.in]: users.map((user) => user.id) } },
  });
  const paymentsByUser = new Map(payments.map((payment) => [payment.userId, payment]));
  const data = users.map((user) => {
    const records = grouped.get(user.id) || [];
    const presentDays = records.filter((record) => record.status === 'hadir').length;
    return buildSalaryRow(user, presentDays, records, paymentsByUser.get(user.id));
  });
  return { data, start, end };
}

app.get('/api/reports/salary', ...authRequired, async (req, res) => {
  try {
    const scopeWhere = salaryScopeWhere(req.user);
    if (!scopeWhere) {
      return res.status(403).json({ ok: false, message: 'Akses laporan gaji tidak tersedia untuk role Anda.' });
    }
    const weekStart = sundayOfWeek(toText(req.query.week) || jakartaDate());
    const { data, start, end } = await computeSalaryReport(weekStart, scopeWhere);
    return res.json({ ok: true, data, period: { weekStart, start, end } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat laporan gaji.' });
  }
});

app.post('/api/reports/salary/adjust', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const userId = toInt(req.body.userId, null);
    if (!userId) {
      return res.status(400).json({ ok: false, message: 'Karyawan tidak valid.' });
    }
    const target = await User.findByPk(userId);
    if (!target) {
      return res.status(404).json({ ok: false, message: 'Karyawan tidak ditemukan.' });
    }
    if (target.role === 'owner' && req.user.role !== 'owner') {
      return res.status(403).json({ ok: false, message: 'Gaji owner tidak bisa diubah oleh leader.' });
    }
    const weekStart = sundayOfWeek(toText(req.body.weekStart) || jakartaDate());
    let payment = await SalaryPayment.findOne({ where: { userId, weekStart } });
    if (!payment) {
      payment = await SalaryPayment.create({ userId, weekStart, paid: false });
    }
    payment.bonus = money(req.body.bonus);
    payment.thr = money(req.body.thr);
    payment.deduction = money(req.body.deduction);
    await payment.save();
    return res.json({ ok: true, message: 'Penyesuaian gaji tersimpan.', data: { userId, weekStart, bonus: payment.bonus, thr: payment.thr, deduction: payment.deduction } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menyimpan penyesuaian gaji.' });
  }
});

app.post('/api/reports/salary/mark-paid', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const userId = toInt(req.body.userId, null);
    if (!userId) {
      return res.status(400).json({ ok: false, message: 'Karyawan tidak valid.' });
    }
    const target = await User.findByPk(userId);
    if (!target) {
      return res.status(404).json({ ok: false, message: 'Karyawan tidak ditemukan.' });
    }
    if (target.role === 'owner' && req.user.role !== 'owner') {
      return res.status(403).json({ ok: false, message: 'Gaji owner tidak bisa diubah oleh leader.' });
    }
    const weekStart = sundayOfWeek(toText(req.body.weekStart) || jakartaDate());
    const paid = Boolean(req.body.paid);
    let payment = await SalaryPayment.findOne({ where: { userId, weekStart } });
    if (!payment) {
      payment = await SalaryPayment.create({ userId, weekStart, paid: false });
    }
    payment.paid = paid;
    payment.paidAt = paid ? now() : null;
    await payment.save();
    return res.json({ ok: true, message: paid ? 'Ditandai sudah dibayar.' : 'Ditandai belum dibayar.', data: { userId, weekStart, paid: payment.paid, paidAt: payment.paidAt } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memperbarui status gajian.' });
  }
});

app.post('/api/reports/salary/mark-all-paid', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const weekStart = sundayOfWeek(toText(req.body.weekStart) || jakartaDate());
    const paid = req.body.paid === undefined ? true : Boolean(req.body.paid);
    const scopeWhere = salaryScopeWhere(req.user) || { active: true };
    const users = await User.findAll({ where: scopeWhere });
    for (const user of users) {
      let payment = await SalaryPayment.findOne({ where: { userId: user.id, weekStart } });
      if (!payment) {
        payment = await SalaryPayment.create({ userId: user.id, weekStart, paid: false });
      }
      payment.paid = paid;
      payment.paidAt = paid ? now() : null;
      await payment.save();
    }
    return res.json({ ok: true, message: paid ? 'Semua karyawan ditandai sudah dibayar.' : 'Semua karyawan ditandai belum dibayar.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memperbarui status gajian.' });
  }
});

app.get('/api/reports/profit', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const from = toText(req.query.from);
    const to = toText(req.query.to);
    const data = await buildProfitSummary(from, to);
    return res.json({ ok: true, data });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat laba rugi.' });
  }
});

function loadPdfKit() {
  try {
    return require('pdfkit');
  } catch (error) {
    return require(path.join(__dirname, '.pdf_vendor', 'node_modules', 'pdfkit'));
  }
}

function buildProfitReportPdf(report) {
  const PDFDocument = loadPdfKit();
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', margin: 48, bufferPages: true });
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const colors = {
      heading: '#0f172a',
      muted: '#64748b',
      line: '#cbd5e1',
      accent: '#0e7490',
      positive: '#047857',
      negative: '#b91c1c',
      boxFill: '#f1f5f9',
    };
    const left = doc.page.margins.left;
    const right = doc.page.width - doc.page.margins.right;
    const pageWidth = right - left;
    const labelWidth = pageWidth * 0.7;
    const amountWidth = pageWidth - labelWidth;
    const bottomLimit = doc.page.height - doc.page.margins.bottom;
    const money = (value) => `Rp ${Math.round(value || 0).toLocaleString('id-ID')}`;

    const ensureSpace = (height) => {
      if (doc.y + height > bottomLimit) {
        doc.addPage();
      }
    };

    const hRule = (color = colors.line, width = 0.75) => {
      doc.moveTo(left, doc.y).lineTo(right, doc.y).strokeColor(color).lineWidth(width).stroke();
    };

    const row = (label, amount, { bold = false, indent = 0, fontSize = 9.5 } = {}) => {
      ensureSpace(16);
      const y = doc.y;
      doc.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(fontSize).fillColor(colors.heading);
      doc.text(label, left + indent, y, { width: labelWidth - indent, lineBreak: false, ellipsis: true });
      doc.text(money(amount), left + labelWidth, y, { width: amountWidth, align: 'right', lineBreak: false });
      doc.y = y + 15;
    };

    const section = (title, rows, total) => {
      ensureSpace(40);
      doc.moveDown(0.6);
      doc.font('Helvetica-Bold').fontSize(10.5).fillColor(colors.accent).text(title.toUpperCase(), left, doc.y, { width: pageWidth });
      doc.moveDown(0.3);
      const visibleRows = rows.filter((item) => item.amount);
      for (const item of visibleRows) row(item.label, item.amount, { indent: 8 });
      doc.moveDown(0.15);
      hRule();
      doc.moveDown(0.15);
      row('Jumlah', total, { bold: true });
    };

    // Header
    doc.font('Helvetica-Bold').fontSize(18).fillColor(colors.heading).text('LAPORAN LABA RUGI', { align: 'center' });
    doc.moveDown(0.15);
    doc.font('Helvetica').fontSize(10).fillColor(colors.muted).text(report.periodLabel || 'Semua data', { align: 'center' });
    const printedAt = new Intl.DateTimeFormat('id-ID', { dateStyle: 'long', timeStyle: 'short', timeZone: 'Asia/Jakarta' }).format(new Date());
    doc.text(`Dicetak: ${printedAt} WIB`, { align: 'center' });
    doc.moveDown(0.6);
    hRule(colors.heading, 1.25);
    doc.moveDown(0.8);

    section('Pendapatan', report.sections.revenue, report.totals.revenue);
    section('Beban Pokok Penjualan', report.sections.costOfGoods, report.totals.hpp);
    section('Beban Operasional', report.sections.operational, report.totals.operational);
    section('Beban Marketing', report.sections.marketing, report.totals.marketing);
    section('Biaya Admin', report.sections.admin, report.totals.adminFee);
    section('Pendapatan Non Operasional', report.sections.nonOperationalIncome, report.totals.nonOperationalIncome);
    section('Beban Non Operasional', report.sections.nonOperationalExpense, report.totals.nonOperationalExpense);

    // Summary box
    const summaryRows = [
      { label: 'Laba Kotor', value: report.totals.grossProfit },
      { label: 'Total Beban Operasional', value: report.totals.operatingExpense },
      { label: 'Laba Bersih', value: report.totals.netProfit, emphasize: true },
    ];
    const rowHeight = 22;
    const boxPadding = 12;
    const boxHeight = boxPadding * 2 + rowHeight * summaryRows.length;
    ensureSpace(boxHeight + 16);
    doc.moveDown(0.6);
    const boxTop = doc.y;
    doc.rect(left, boxTop, pageWidth, boxHeight).fillAndStroke(colors.boxFill, colors.line);
    let sy = boxTop + boxPadding;
    for (const item of summaryRows) {
      const fontSize = item.emphasize ? 12.5 : 10;
      const color = item.emphasize ? (item.value >= 0 ? colors.positive : colors.negative) : colors.heading;
      doc.font('Helvetica-Bold').fontSize(fontSize).fillColor(color);
      doc.text(item.label, left + boxPadding, sy, { width: labelWidth - boxPadding, lineBreak: false });
      doc.text(money(item.value), left + labelWidth, sy, { width: amountWidth - boxPadding, align: 'right', lineBreak: false });
      sy += rowHeight;
    }
    doc.y = boxTop + boxHeight;

    // Footer with page numbers (temporarily zero the bottom margin so the
    // footer itself doesn't trigger pdfkit's auto page-break)
    const pageRange = doc.bufferedPageRange();
    const originalBottomMargin = doc.page.margins.bottom;
    for (let i = pageRange.start; i < pageRange.start + pageRange.count; i++) {
      doc.switchToPage(i);
      doc.page.margins.bottom = 0;
      doc.font('Helvetica').fontSize(8).fillColor(colors.muted);
      doc.text(
        `Halaman ${i + 1} dari ${pageRange.count}`,
        left,
        doc.page.height - originalBottomMargin + 12,
        { width: pageWidth, align: 'center' },
      );
      doc.page.margins.bottom = originalBottomMargin;
    }

    doc.end();
  });
}

app.get('/api/reports/profit/pdf', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const from = toText(req.query.from);
    const to = toText(req.query.to);
    const data = await buildProfitSummary(from, to);
    const buffer = await buildProfitReportPdf(data);
    res.setHeader('Content-Disposition', 'attachment; filename="laba_rugi.pdf"');
    res.setHeader('Content-Type', 'application/pdf');
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal membuat PDF laba rugi.' });
  }
});

app.get('/api/dashboard', ...authRequired, async (req, res) => {
  try {
    if (req.user.role === 'admin') {
      return res.status(403).json({ ok: false, message: 'Dashboard tidak tersedia untuk role Anda.' });
    }
    const from = toText(req.query.from);
    const to = toText(req.query.to);
    const data = await buildDashboard(from, to, { includeHpp: req.user.role === 'owner' });
    return res.json({ ok: true, data });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat dashboard.' });
  }
});

app.use((err, req, res, next) => {
  if (err) {
    return res.status(500).json({ ok: false, message: 'Terjadi kesalahan server.' });
  }
  return next();
});

app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

async function seedExpenseCategories() {
  const seeds = [
    ...DEFAULT_OPERATIONAL_ORDER.map((name) => ({ name, kind: 'operasional' })),
    ...DEFAULT_MARKETING_ORDER.map((name) => ({ name, kind: 'marketing' })),
  ];
  for (const seed of seeds) {
    await ExpenseCategory.findOrCreate({ where: seed, defaults: seed });
  }
}

async function bootstrap() {
  await ensureDatabase();
  await sequelize.authenticate();
  await sequelize.sync({ alter: true });
  await migrateOwnerRole();
  await createInitialAdmin();
  await getAttendanceSettings();
  await seedExpenseCategories();
}

bootstrap()
  .then(() => {
    app.listen(PORT, '127.0.0.1', () => {
      console.log(`Server running on http://127.0.0.1:${PORT}`);
    });
  })
  .catch((error) => {
    console.error('Gagal menjalankan server:', error);
    process.exit(1);
  });
