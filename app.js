const express = require('express');
const cors = require('cors');
const path = require('path');
const bcrypt = require('bcryptjs');
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
const PORT = process.env.PORT || 3299;
const JWT_SECRET =  'topi';
const DB_HOST = process.env.MYSQL_HOST || '127.0.0.1';
const DB_PORT = Number(process.env.MYSQL_PORT || 3306);
const DB_USER = process.env.MYSQL_USER || 'root';
const DB_PASSWORD = process.env.MYSQL_PASSWORD || 'alen';
// Database khusus aplikasi ini. Sebelumnya bernama 'topi', tapi database itu
// dipakai bersama project lain (app builder project_196) yang menjalankan
// sequelize.sync({ alter: true }) dengan model berbeda, sehingga kolom
// shifts.business_id berulang kali di-DROP dan bikin bootstrap gagal + login 500.
const DB_NAME = process.env.MYSQL_DATABASE || 'topi_pos';

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
// Batch version of isOvertimeDay so payroll can classify a whole period with a
// single query instead of one lookup per attendance row.
async function overtimeDateSet(dates) {
  const unique = [...new Set((dates || []).map((value) => toText(value)).filter(Boolean))];
  if (!unique.length) {
    return new Set();
  }
  const holidays = await Holiday.findAll({ where: { date: { [Op.in]: unique } } });
  const holidayDates = new Set(holidays.map((row) => toText(row.date)));
  const result = new Set();
  for (const date of unique) {
    const [y, m, d] = date.split('-').map(Number);
    if (new Date(y, m - 1, d).getDay() === 0 || holidayDates.has(date)) {
      result.add(date);
    }
  }
  return result;
}
function shiftDurationHours(user, settings) {
  const source = user && user.Shift ? user.Shift : settings;
  const start = String((source && source.checkInStart) || '07:00').split(':').map(Number);
  const end = String((source && source.checkOutStart) || '17:00').split(':').map(Number);
  const startMinutes = (start[0] || 0) * 60 + (start[1] || 0);
  let endMinutes = (end[0] || 0) * 60 + (end[1] || 0);
  if (endMinutes <= startMinutes) {
    endMinutes += 24 * 60;
  }
  return Math.round(((endMinutes - startMinutes) / 60) * 100) / 100;
}
async function resolveManualOvertimeHours(user, workDate, status, businessId) {
  if (status !== 'hadir') {
    return 0;
  }
  // Owner tidak pernah mendapat lembur, hanya gaji pokok.
  if (user && user.role === 'owner') {
    return 0;
  }
  if (!(await isOvertimeDay(workDate))) {
    return 0;
  }
  const settings = await getAttendanceSettings(businessId);
  return shiftDurationHours(user, settings);
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
  dialectOptions: {
    // Raw sequelize.query() calls (used for sales/dashboard aggregation) bypass
    // Sequelize's own DATEONLY parsing and get the mysql2 driver's value directly.
    // mysql2 hands back DATE columns as JS Date objects at local midnight; once
    // Express JSON-serializes that Date it calls toISOString(), which converts to
    // UTC and shifts the calendar date back a day for any timezone ahead of UTC
    // (e.g. WIB). Keeping DATE columns as plain 'YYYY-MM-DD' strings avoids that
    // round trip entirely. DATETIME/TIMESTAMP columns are left untouched.
    dateStrings: ['DATE'],
  },
  define: {
    underscored: true,
    timestamps: true,
  },
});

const Business = sequelize.define('Business', {
  name: {
    type: DataTypes.STRING(120),
    allowNull: false,
  },
  slug: {
    type: DataTypes.STRING(120),
    allowNull: false,
    unique: true,
  },
  active: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: true,
  },
}, {
  tableName: 'businesses',
});

const User = sequelize.define('User', {
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  indexes: [
    // Not unique: production data already has ~30 pre-existing duplicate SKUs
    // (never enforced before this feature). Kept as a plain index for lookup
    // performance; app logic scopes by businessId but doesn't hard-fail on dupes.
    { fields: ['business_id', 'sku'] },
  ],
});

const Sale = sequelize.define('Sale', {
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
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
});

// Owner accounts are global so they can access every business, but their salary
// components belong to a specific business. Keeping these values in a separate
// table prevents an edit in one store from overwriting the other store.
const OwnerSalarySetting = sequelize.define('OwnerSalarySetting', {
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
  userId: {
    type: DataTypes.INTEGER,
    allowNull: false,
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
}, {
  tableName: 'owner_salary_settings',
  indexes: [
    { unique: true, fields: ['business_id', 'user_id'] },
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
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
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
    { unique: true, fields: ['business_id', 'name', 'kind'] },
  ],
});

const ProfitSharePartner = sequelize.define('ProfitSharePartner', {
  businessId: {
    type: DataTypes.INTEGER,
    allowNull: true,
  },
  name: {
    type: DataTypes.STRING(120),
    allowNull: false,
  },
  percentage: {
    type: DataTypes.FLOAT,
    allowNull: false,
    defaultValue: 0,
  },
  sortOrder: {
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
  tableName: 'profit_share_partners',
});

const ProfitShareTarget = sequelize.define('ProfitShareTarget', {
  partnerId: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
  period: {
    type: DataTypes.STRING(7),
    allowNull: false,
  },
  amount: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
}, {
  tableName: 'profit_share_targets',
  indexes: [
    { unique: true, fields: ['partner_id', 'period'] },
  ],
});

const ProfitShareInstallment = sequelize.define('ProfitShareInstallment', {
  partnerId: {
    type: DataTypes.INTEGER,
    allowNull: false,
  },
  period: {
    type: DataTypes.STRING(7),
    allowNull: false,
  },
  paidDate: {
    type: DataTypes.DATEONLY,
    allowNull: false,
  },
  amount: {
    type: DataTypes.INTEGER,
    allowNull: false,
    defaultValue: 0,
  },
  note: {
    type: DataTypes.STRING(255),
    allowNull: true,
  },
}, {
  tableName: 'profit_share_installments',
});

ProfitSharePartner.hasMany(ProfitShareTarget, { foreignKey: 'partnerId' });
ProfitShareTarget.belongsTo(ProfitSharePartner, { foreignKey: 'partnerId' });
ProfitSharePartner.hasMany(ProfitShareInstallment, { foreignKey: 'partnerId' });
ProfitShareInstallment.belongsTo(ProfitSharePartner, { foreignKey: 'partnerId' });

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
User.hasMany(OwnerSalarySetting, { foreignKey: 'userId' });
OwnerSalarySetting.belongsTo(User, { foreignKey: 'userId' });
Business.hasMany(OwnerSalarySetting, { foreignKey: 'businessId' });
OwnerSalarySetting.belongsTo(Business, { foreignKey: 'businessId' });

async function getAttendanceSettings(businessId) {
  let row = await AttendanceSetting.findOne({ where: { businessId }, order: [['id', 'ASC']] });
  if (!row) {
    row = await AttendanceSetting.create({ businessId });
  }
  return row;
}

const serializeBusiness = (row) => ({
  id: row.id,
  name: row.name,
  slug: row.slug,
  active: Boolean(row.active),
});

const serializeUser = (user, salarySetting = null) => ({
  id: user.id,
  businessId: salarySetting?.businessId || user.businessId || null,
  username: user.username,
  fullName: user.fullName,
  role: user.role,
  jobTitle: user.jobTitle || '',
  workShift: user.workShift || '',
  phone: user.phone || '',
  dailyWage: salarySetting ? (salarySetting.dailyWage || 0) : (user.dailyWage || 0),
  mealAllowance: salarySetting ? (salarySetting.mealAllowance || 0) : (user.mealAllowance || 0),
  overtimeRatePerHour: salarySetting ? (salarySetting.overtimeRatePerHour || 0) : (user.overtimeRatePerHour || 0),
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

async function ownerSalarySettingsByUser(users, businessId) {
  const ownerIds = users.filter((user) => user.role === 'owner').map((user) => user.id);
  if (!ownerIds.length || businessId === null || businessId === undefined) return new Map();
  const rows = await OwnerSalarySetting.findAll({
    where: { businessId, userId: { [Op.in]: ownerIds } },
  });
  return new Map(rows.map((row) => [row.userId, row]));
}

async function attendanceUserScopeWhere(req, extraWhere = {}) {
  if (req.businessId === null || req.businessId === undefined) return extraWhere;
  const settings = await OwnerSalarySetting.findAll({
    where: { businessId: req.businessId },
    attributes: ['userId'],
  });
  const ownerIds = settings.map((setting) => setting.userId);
  return {
    ...extraWhere,
    [Op.or]: [
      { businessId: req.businessId },
      ...(ownerIds.length ? [{ id: { [Op.in]: ownerIds } }] : []),
    ],
  };
}

async function isUserInAttendanceBusiness(user, businessId) {
  if (!user || businessId === null || businessId === undefined) return Boolean(user);
  if (user.role !== 'owner') return user.businessId === businessId;
  return Boolean(await OwnerSalarySetting.findOne({
    where: { userId: user.id, businessId },
    attributes: ['id'],
  }));
}

async function upsertOwnerSalarySetting(user, businessId, values = {}) {
  const [row] = await OwnerSalarySetting.findOrCreate({
    where: { userId: user.id, businessId },
    defaults: {
      userId: user.id,
      businessId,
      dailyWage: user.dailyWage || 0,
      mealAllowance: user.mealAllowance || 0,
      overtimeRatePerHour: user.overtimeRatePerHour || 0,
    },
  });
  row.dailyWage = money(values.dailyWage);
  row.mealAllowance = money(values.mealAllowance);
  row.overtimeRatePerHour = money(values.overtimeRatePerHour);
  await row.save();
  return row;
}

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
  businessId: row.businessId || null,
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
  businessId: row.businessId || null,
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
  businessId: row.businessId || null,
  sku: row.sku,
  name: row.name,
  variant: row.variant,
  hpp: row.hpp,
  createdAt: row.createdAt,
  updatedAt: row.updatedAt,
});

const serializeSale = (row, { includeHpp = true } = {}) => ({
  id: row.id,
  businessId: row.businessId || null,
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
  businessId: row.businessId || null,
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
  businessId: row.businessId || null,
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
  businessId: row.businessId || null,
  name: row.name,
  kind: row.kind,
});

const serializeProfitSharePartner = (row) => ({
  id: row.id,
  businessId: row.businessId || null,
  name: row.name,
  percentage: row.percentage,
  sortOrder: row.sortOrder,
  active: Boolean(row.active),
});

const serializeProfitShareInstallment = (row) => ({
  id: row.id,
  partnerId: row.partnerId,
  period: row.period,
  paidDate: row.paidDate,
  amount: row.amount,
  note: row.note || '',
  createdAt: row.createdAt,
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
    return res.status(403).json({ ok: false, message: 'Akses ini hanya untuk owner atau kepala toko.' });
  }
  return next();
};

const salesRecapAccess = (req, res, next) => {
  if (!['owner', 'leader', 'admin'].includes(req.user?.role)) {
    return res.status(403).json({ ok: false, message: 'Akses ini tidak tersedia untuk role Anda.' });
  }
  return next();
};

const ownerLeaderAdmin = (req, res, next) => {
  if (!['owner', 'leader', 'admin'].includes(req.user?.role)) {
    return res.status(403).json({ ok: false, message: 'Akses ini tidak tersedia untuk role Anda.' });
  }
  return next();
};

const selfServiceRole = (req, res, next) => {
  if (!['owner', 'leader', 'karyawan', 'admin'].includes(req.user?.role)) {
    return res.status(403).json({ ok: false, message: 'Fitur ini tidak tersedia untuk role Anda.' });
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

function resolveBusiness(req, res, next) {
  const rawHeader = req.headers['x-business-id'];
  const requested = rawHeader !== undefined ? toText(rawHeader) : toText(req.query.businessId);
  if (req.user.role === 'owner') {
    req.businessId = (!requested || requested.toLowerCase() === 'all') ? null : toInt(requested, null);
    return next();
  }
  if (!req.user.businessId) {
    return res.status(403).json({ ok: false, message: 'Akun Anda belum ditugaskan ke bisnis manapun. Hubungi owner.' });
  }
  req.businessId = req.user.businessId;
  return next();
}

// req.businessId === null means "all businesses" (owner cross-business view, read-only).
function businessWhere(req, extraWhere = {}) {
  if (req.businessId === null || req.businessId === undefined) return extraWhere;
  return { ...extraWhere, businessId: req.businessId };
}

// Like businessWhere, but keeps owner-role users visible even when a specific
// business is selected — owner accounts have businessId NULL by design (they
// span all businesses), so a plain businessId filter would hide them.
function businessOrOwnerWhere(req, extraWhere = {}) {
  if (req.businessId === null || req.businessId === undefined) return extraWhere;
  return { ...extraWhere, [Op.or]: [{ businessId: req.businessId }, { role: 'owner' }] };
}

function assertOwnedByBusiness(row, req) {
  if (!row) return false;
  if (req.businessId === null || req.businessId === undefined) return true;
  return row.businessId === req.businessId;
}

// Owner accounts are global (businessId NULL) and remain visible in every
// business view, so user mutations must treat them as part of the same scope.
function assertUserInBusinessScope(row, req) {
  return Boolean(row) && (row.role === 'owner' || assertOwnedByBusiness(row, req));
}

function requireConcreteBusiness(req, res) {
  if (req.businessId === null || req.businessId === undefined) {
    res.status(400).json({ ok: false, message: 'Pilih bisnis terlebih dahulu. Data tidak bisa disimpan pada tampilan "Semua Bisnis".' });
    return false;
  }
  return true;
}

const authRequired = [auth, loadCurrentUser, resolveBusiness];

// Hari lembur (minggu/hari libur) dibayar sebagai lembur saja: gaji pokok dan
// uang makan hanya dihitung pada hari kerja biasa. Owner tidak punya lembur,
// sehingga seluruh kehadirannya dibayar sebagai gaji pokok.
function splitPaidDays(user, attendanceRecords, presentDays, overtimeDates) {
  const isOwner = user && user.role === 'owner';
  const records = attendanceRecords || [];
  if (isOwner) {
    return { paidDays: presentDays, overtimeHours: 0 };
  }
  const overtimeHours = records.reduce((sum, record) => sum + (record.overtimeHours || 0), 0);
  if (!overtimeDates || !overtimeDates.size) {
    return { paidDays: presentDays, overtimeHours };
  }
  const overtimeDayCount = records
    .filter((record) => record.status === 'hadir' && overtimeDates.has(toText(record.workDate)))
    .length;
  return { paidDays: Math.max(0, presentDays - overtimeDayCount), overtimeHours };
}

function buildSalaryRow(user, presentDays, attendanceRecords, paymentRow, salarySetting = null, overtimeDates = null) {
  const salarySource = salarySetting || user;
  const dailyWage = salarySource.dailyWage || 0;
  const mealAllowance = salarySource.mealAllowance || 0;
  const overtimeRatePerHour = user && user.role === 'owner' ? 0 : (salarySource.overtimeRatePerHour || 0);
  const { paidDays, overtimeHours } = splitPaidDays(user, attendanceRecords, presentDays, overtimeDates);
  const overtimePay = Math.round(overtimeHours * overtimeRatePerHour);
  const bonus = paymentRow?.bonus || 0;
  const thr = paymentRow?.thr || 0;
  const deduction = paymentRow?.deduction || 0;
  const salaryBase = paidDays * dailyWage;
  const mealTotal = paidDays * mealAllowance;
  const gross = salaryBase + mealTotal + overtimePay + bonus + thr;
  const net = gross - deduction;
  return {
    user: serializeUser(user, salarySetting),
    presentDays,
    paidDays,
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

const SALE_GROUP_KEY_SQL = "CASE WHEN `order_number` IS NULL OR `order_number` = '' THEN CONCAT('__row_', `id`) ELSE `order_number` END";

function buildDashboardSaleConditions({ from, to, channel, storeName, sku, businessId } = {}) {
  const conditions = [];
  const replacements = {};
  if (from) {
    conditions.push('`sales`.`sale_date` >= :from');
    replacements.from = from;
  }
  if (to) {
    conditions.push('`sales`.`sale_date` <= :to');
    replacements.to = to;
  }
  if (channel) {
    conditions.push('`sales`.`channel` = :channel');
    replacements.channel = channel;
  }
  if (storeName) {
    conditions.push('`sales`.`store_name` = :storeName');
    replacements.storeName = storeName;
  }
  if (sku) {
    conditions.push('`sales`.`sku` = :sku');
    replacements.sku = sku;
  }
  if (businessId !== null && businessId !== undefined) {
    conditions.push('`sales`.`business_id` = :businessId');
    replacements.businessId = businessId;
  }
  return { whereSql: conditions.length ? `WHERE ${conditions.join(' AND ')}` : '', replacements };
}

async function countDistinctOrders(filters) {
  const { whereSql, replacements } = buildDashboardSaleConditions(filters);
  const rows = await sequelize.query(
    `SELECT COUNT(*) AS total FROM (SELECT 1 FROM \`sales\` ${whereSql} GROUP BY ${SALE_GROUP_KEY_SQL}) t`,
    { replacements, type: sequelize.QueryTypes.SELECT },
  );
  return toInt(rows[0]?.total, 0);
}

async function buildRevenueTrend(filters, monthsBack = 11) {
  const anchor = filters.to ? new Date(`${filters.to}T00:00:00`) : now();
  const start = new Date(anchor.getFullYear(), anchor.getMonth() - monthsBack, 1);
  const startMonth = `${start.getFullYear()}-${pad(start.getMonth() + 1)}-01`;
  const { whereSql, replacements } = buildDashboardSaleConditions({ ...filters, from: startMonth, to: undefined });
  const rows = await sequelize.query(
    `SELECT DATE_FORMAT(\`sale_date\`, '%Y-%m') AS month, SUM(\`subtotal\`) AS revenue
     FROM \`sales\`
     ${whereSql}
     GROUP BY month
     ORDER BY month ASC`,
    { replacements, type: sequelize.QueryTypes.SELECT },
  );
  const byMonth = new Map(rows.map((row) => [row.month, toInt(row.revenue, 0)]));
  const series = [];
  for (let i = monthsBack; i >= 0; i -= 1) {
    const d = new Date(anchor.getFullYear(), anchor.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
    series.push({ month: key, revenue: byMonth.get(key) || 0 });
  }
  return series;
}

async function buildPlatformContribution(filters) {
  const { whereSql, replacements } = buildDashboardSaleConditions(filters);
  const rows = await sequelize.query(
    `SELECT \`channel\`, \`store_name\` AS storeName,
       COUNT(DISTINCT ${SALE_GROUP_KEY_SQL}) AS totalOrders,
       SUM(\`qty\`) AS totalQty,
       SUM(\`subtotal\`) AS totalSales
     FROM \`sales\`
     ${whereSql}
     GROUP BY \`channel\`, \`store_name\`
     ORDER BY totalSales DESC`,
    { replacements, type: sequelize.QueryTypes.SELECT },
  );
  return rows.map((row) => ({
    label: `${row.channel} ${row.storeName}`.trim(),
    totalOrders: toInt(row.totalOrders, 0),
    totalQty: toInt(row.totalQty, 0),
    totalSales: toInt(row.totalSales, 0),
  }));
}

async function buildTopProducts(filters, limit = 5) {
  const { whereSql, replacements } = buildDashboardSaleConditions(filters);
  const rows = await sequelize.query(
    `SELECT \`sales\`.\`sku\` AS sku, MIN(\`products\`.\`name\`) AS name,
       SUM(\`sales\`.\`qty\`) AS qty, SUM(\`sales\`.\`subtotal\`) AS totalSales
     FROM \`sales\`
     LEFT JOIN \`products\` ON \`products\`.\`sku\` = \`sales\`.\`sku\` AND \`products\`.\`business_id\` <=> \`sales\`.\`business_id\`
     ${whereSql}
     GROUP BY \`sales\`.\`sku\`
     ORDER BY totalSales DESC
     LIMIT ${Number(limit)}`,
    { replacements, type: sequelize.QueryTypes.SELECT },
  );
  return rows.map((row, index) => ({
    rank: index + 1,
    sku: row.sku,
    name: row.name || row.sku,
    qty: toInt(row.qty, 0),
    totalSales: toInt(row.totalSales, 0),
  }));
}

async function buildProfitSummary(from, to, saleFilters = {}, businessId = null) {
  const bizWhere = (businessId === null || businessId === undefined) ? {} : { businessId };
  const saleWhere = { ...dateWhere('saleDate', from, to), ...saleFilters, ...bizWhere };
  const expenseWhere = { ...dateWhere('expenseDate', from, to), ...bizWhere };
  const reportDate = new Intl.DateTimeFormat('id-ID', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });

  const marketingWhere = { ...expenseWhere };
  if (saleFilters.channel || saleFilters.storeName) {
    const matchingStores = await Store.findAll({
      where: {
        ...bizWhere,
        ...(saleFilters.channel ? { platform: saleFilters.channel } : {}),
        ...(saleFilters.storeName ? { name: saleFilters.storeName } : {}),
      },
      raw: true,
    });
    const platformStoreValues = matchingStores.map((store) => store.platformStore).filter(Boolean);
    marketingWhere.platformStore = { [Op.in]: platformStoreValues };
  }

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
    where: marketingWhere,
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

  // Gaji dari Laporan Gaji ikut masuk sebagai beban operasional toko supaya
  // laba rugi tidak perlu input manual lagi.
  const gajiKaryawan = await computeGajiTotalForPeriod(from, to, businessId);
  bucket(operationalMap, 'Beban Gaji, Upah & Honorer', gajiKaryawan);

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
      gajiKaryawan,
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

async function computeGajiTotalForPeriod(from, to, businessId = null) {
  const attendanceWhere = dateWhere('workDate', from, to);
  const paymentWhere = {
    ...dateWhere('weekStart', from, to),
    ...((businessId === null || businessId === undefined) ? {} : { businessId }),
  };
  const userWhere = (businessId === null || businessId === undefined)
    ? { active: true }
    : { active: true, [Op.or]: [{ businessId }, { role: 'owner' }] };
  const candidateUsers = await User.findAll({ where: userWhere });
  const ownerSettings = await ownerSalarySettingsByUser(candidateUsers, businessId);
  const users = (businessId === null || businessId === undefined)
    ? candidateUsers
    : candidateUsers.filter((user) => user.role !== 'owner' || ownerSettings.has(user.id));
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

  const overtimeDates = await overtimeDateSet(attendanceRows.map((row) => row.workDate));

  let total = 0;
  for (const user of users) {
    const salarySource = ownerSettings.get(user.id) || user;
    const records = attendanceByUser.get(user.id) || [];
    const presentDays = records.filter((record) => record.status === 'hadir').length;
    const { paidDays, overtimeHours } = splitPaidDays(user, records, presentDays, overtimeDates);
    const overtimeRate = user.role === 'owner' ? 0 : (salarySource.overtimeRatePerHour || 0);
    const overtimePay = Math.round(overtimeHours * overtimeRate);
    const salaryBase = paidDays * (salarySource.dailyWage || 0);
    const mealTotal = paidDays * (salarySource.mealAllowance || 0);
    const userPayments = paymentsByUser.get(user.id) || [];
    const bonus = userPayments.reduce((sum, payment) => sum + (payment.bonus || 0), 0);
    const thr = userPayments.reduce((sum, payment) => sum + (payment.thr || 0), 0);
    const deduction = userPayments.reduce((sum, payment) => sum + (payment.deduction || 0), 0);
    total += salaryBase + mealTotal + overtimePay + bonus + thr - deduction;
  }
  return total;
}

async function buildDashboard(from, to, { platform = '', storeName = '', sku = '', businessId = null } = {}) {
  const saleFilters = { from, to, channel: platform, storeName, sku, businessId };
  const profitSummary = await buildProfitSummary(from, to, {
    ...(platform ? { channel: platform } : {}),
    ...(storeName ? { storeName } : {}),
    ...(sku ? { sku } : {}),
  }, businessId);
  // Gaji sudah termasuk di dalam total beban operasional laba rugi.
  const gajiKaryawan = profitSummary.totals.gajiKaryawan;
  const totalBebanPeriode = profitSummary.totals.operational + profitSummary.totals.marketing + profitSummary.totals.adminFee;

  const [totalOrders, revenueTrend, platformContribution, topProducts] = await Promise.all([
    countDistinctOrders(saleFilters),
    buildRevenueTrend(saleFilters),
    buildPlatformContribution(saleFilters),
    buildTopProducts(saleFilters),
  ]);

  return {
    period: { from, to },
    counts: {
      sales: profitSummary.counts.sales,
      operational: profitSummary.counts.operational,
      marketing: profitSummary.counts.marketing,
      totalOrders,
    },
    totals: {
      ...profitSummary.totals,
      gajiKaryawan,
      totalBebanPeriode,
    },
    revenueTrend,
    platformContribution,
    topProducts,
  };
}

async function migrateOwnerRole() {
  // Only the original administrator account used the legacy `admin` role as
  // the system owner. Converting every admin on each boot promotes real store
  // admins into owners and incorrectly puts them in the owner salary report.
  await sequelize.query("UPDATE `users` SET `role` = 'owner' WHERE `username` = 'admin' AND `role` = 'admin'");
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
      const where = businessWhere(req, {
        ...(options.searchFields ? searchWhere(req.query.search, options.searchFields) : {}),
        ...(options.dateField ? dateWhere(options.dateField, toText(req.query.from), toText(req.query.to)) : {}),
      });
      if (options.paginated) {
        const { page, limit, offset } = getPagination(req.query);
        const requestedSort = toText(req.query.sortBy);
        const sortBy = (options.sortFields || []).includes(requestedSort) ? requestedSort : 'createdAt';
        const sortDir = toText(req.query.sortDir).toLowerCase() === 'asc' ? 'ASC' : 'DESC';
        const { count, rows } = await model.findAndCountAll({
          where,
          include: options.include || [],
          order: [[sortBy, sortDir], ['id', 'DESC']],
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
      if (!requireConcreteBusiness(req, res)) return;
      const payload = await parser(req.body, null, req.businessId);
      const row = await model.create({ ...payload, businessId: req.businessId });
      const output = options.reload ? await model.findByPk(row.id, { include: options.include || [] }) : row;
      return res.status(201).json({ ok: true, data: serializer(output) });
    } catch (error) {
      return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan data.' });
    }
  });

  app.put(`/api/${prefix}/:id`, ...authRequired, writeMiddleware, async (req, res) => {
    try {
      const row = await model.findByPk(req.params.id);
      if (!row || !assertOwnedByBusiness(row, req)) {
        return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
      }
      const payload = await parser(req.body, row, req.businessId);
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
      if (!row || !assertOwnedByBusiness(row, req)) {
        return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
      }
      await row.destroy();
      return res.json({ ok: true, message: 'Data berhasil dihapus.' });
    } catch (error) {
      return res.status(500).json({ ok: false, message: 'Gagal menghapus data.' });
    }
  });

  app.post(`/api/${prefix}/bulk-delete`, ...authRequired, writeMiddleware, async (req, res) => {
    try {
      const ids = Array.isArray(req.body.ids) ? req.body.ids.map((id) => toInt(id, null)).filter((id) => id !== null) : [];
      if (!ids.length) {
        return res.status(400).json({ ok: false, message: 'Pilih minimal satu data untuk dihapus.' });
      }
      const deleted = await model.destroy({ where: businessWhere(req, { id: { [Op.in]: ids } }) });
      return res.json({ ok: true, message: `${deleted} data berhasil dihapus.`, data: { deleted } });
    } catch (error) {
      return res.status(500).json({ ok: false, message: 'Gagal menghapus data terpilih.' });
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
  if (!expenseDate || !category) {
    throw new Error('Tanggal dan kategori wajib diisi.');
  }
  return {
    expenseDate,
    category,
    description: description || '-',
    nominal: money(body.nominal),
    paymentMethod: toText(body.paymentMethod) || 'Transfer',
  };
}

async function parseMarketing(body, existingRow, businessId) {
  const expenseDate = toText(body.expenseDate || body.date);
  const category = toText(body.category);
  const platformStore = toText(body.platformStore || body.platformAndStore);
  const description = toText(body.description || body.note);
  if (!expenseDate || !category || !platformStore) {
    throw new Error('Tanggal, kategori, dan platform & toko wajib diisi.');
  }
  const nominal = money(body.nominal);
  let taxRate = body.taxRate === undefined || body.taxRate === null || body.taxRate === '' ? 0 : Number(body.taxRate);
  if (!Number.isFinite(taxRate)) taxRate = 0;
  if (taxRate === 0) {
    const store = await Store.findOne({ where: { platformStore, ...(businessId ? { businessId } : {}) } });
    if (store) {
      taxRate = store.taxRate || 0;
    }
  }
  const totalTaxInput = body.totalTax === undefined || body.totalTax === null || body.totalTax === '' ? 0 : Number(body.totalTax);
  const totalTax = totalTaxInput ? money(body.totalTax) : Math.round(nominal * taxRate);
  return {
    expenseDate,
    category,
    platformId: null,
    platformStore,
    description: description || '-',
    nominal,
    taxRate,
    totalTax,
  };
}

async function parseSale(body, existingRow, businessId) {
  const saleDate = toText(body.saleDate || body.date);
  const channelInput = toText(body.channel);
  const channel = channelInput;
  const storeName = toText(body.storeName || body.platformStore);
  const orderNumber = toText(body.orderNumber || body.noPesanan);
  const sku = toText(body.sku);
  if (!saleDate || !channel || !storeName || !sku) {
    throw new Error('Tanggal, channel, toko, dan SKU wajib diisi.');
  }
  const product = await Product.findOne({ where: { sku, ...(businessId ? { businessId } : {}) } });
  const qty = Math.max(0, money(body.qty ?? 1));
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
    // Keep the calendar date as entered in Excel. Converting a local midnight
    // value to ISO/UTC shifts Indonesian dates back one day (for example,
    // "1 Februari 2026" became "2026-01-31").
    return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
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

// Normalizes a raw parsed row into its final field values and flags rows that are
// missing required fields. Pure/synchronous so it can run for every row up front,
// before any database round trip.
function normalizeImportedSaleRecord(record, summary) {
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

  if (!channel || !storeName || !saleDate || !sku) {
    return 'invalid';
  }
  if (!summary.firstSaleDate || saleDate < summary.firstSaleDate) {
    summary.firstSaleDate = saleDate;
  }
  if (!summary.lastSaleDate || saleDate > summary.lastSaleDate) {
    summary.lastSaleDate = saleDate;
  }

  return {
    saleDate,
    channel,
    storeName,
    orderNumber,
    customer,
    sku,
    qty,
    price,
    subtotal,
    adminFee,
    hpp: record.hpp,
    totalHpp: record.totalHpp,
    productName: record.productName,
    variant: record.variant,
  };
}

// Ensures every SKU referenced by `records` has a Product row, in at most two queries
// (one lookup, one bulk insert) instead of one lookup+insert pair per row. Mirrors the
// previous per-row behavior: a SKU already known to the "Data Barang" sheet or already
// in the database keeps its existing hpp; a genuinely new SKU is created once, using the
// first row that mentions it, even if that SKU repeats across many rows.
async function ensureImportedProducts(records, productHpp, summary, businessId, transaction) {
  const hppBySku = new Map();
  const firstRecordBySku = new Map();
  for (const record of records) {
    if (record.sku && !firstRecordBySku.has(record.sku)) {
      firstRecordBySku.set(record.sku, record);
    }
  }
  const skusNeeded = new Set(firstRecordBySku.keys());
  if (!skusNeeded.size) return hppBySku;

  const existing = await Product.findAll({
    where: { sku: { [Op.in]: Array.from(skusNeeded) }, businessId },
    transaction,
  });
  for (const product of existing) {
    hppBySku.set(product.sku, product.hpp || 0);
    skusNeeded.delete(product.sku);
  }

  if (skusNeeded.size) {
    const toCreate = [];
    for (const sku of skusNeeded) {
      const record = firstRecordBySku.get(sku);
      const hpp = productHpp.get(sku) || record.hpp || 0;
      hppBySku.set(sku, hpp);
      toCreate.push({
        businessId,
        sku,
        name: record.productName || sku,
        variant: record.variant || '-',
        hpp,
      });
    }
    await Product.bulkCreate(toCreate, { transaction });
    summary.productsCreated += toCreate.length;
  }
  return hppBySku;
}

// A single order number legitimately can repeat the exact same SKU more than once (e.g. an
// embroidery add-on fee line per item). A plain existence check can't tell "already
// imported before" apart from "another genuinely identical line in this order", so we
// track how many times this order+SKU has been seen so far in this import and only skip
// once that many rows already exist in the database. Fetching those pre-import counts
// once up front (rather than re-querying after every row) gives identical create/skip
// decisions, because the comparison only ever depends on the count of matching rows that
// existed *before* this import started, not on rows this same import has inserted so far.
async function fetchExistingOrderSkuCounts(records, businessId, transaction) {
  const orderNumbers = Array.from(new Set(records.map((r) => r.orderNumber).filter(Boolean)));
  const counts = new Map();
  if (!orderNumbers.length) return counts;
  const rows = await Sale.findAll({
    attributes: ['orderNumber', 'sku', [sequelize.fn('COUNT', sequelize.col('id')), 'cnt']],
    where: { businessId, orderNumber: { [Op.in]: orderNumbers } },
    group: ['orderNumber', 'sku'],
    raw: true,
    transaction,
  });
  for (const row of rows) {
    counts.set(JSON.stringify([row.orderNumber, row.sku]), Number(row.cnt) || 0);
  }
  return counts;
}

async function importSalesWorkbook(buffer, fileName = '', businessId) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const summary = {
    productsCreated: 0,
    productsUpdated: 0,
    salesCreated: 0,
    salesSkipped: 0,
    firstSaleDate: '',
    lastSaleDate: '',
    sheets: [],
  };
  const sheetFailureGroups = [];

  // The whole import runs inside one transaction: previously each parsed row was written
  // with its own Product/Sale queries as soon as it was parsed, so a failure partway
  // through (e.g. while building the failed-rows report) left the rows seen so far
  // committed even though the request ultimately responded with an error. Wrapping the
  // writes in a transaction makes the import all-or-nothing, so a reported failure can no
  // longer leave data behind that only shows up after a refresh.
  await sequelize.transaction(async (t) => {
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
          where: { sku, businessId },
          defaults: {
            businessId,
            sku,
            name,
            variant: variant || '-',
            hpp,
          },
          transaction: t,
        });
        if (created) {
          summary.productsCreated += 1;
        } else {
          await product.update({
            name,
            variant: variant || product.variant || '-',
            hpp,
          }, { transaction: t });
          summary.productsUpdated += 1;
        }
      }
    }

    const defaultStoreName = detectStoreNameFromFile(fileName);
    const saleSheetNames = workbook.SheetNames.filter((name) => name !== 'Data Barang');

    // Pass 1: parse every sheet's rows into candidate records without touching the
    // database, so the (much slower) DB work below can run in a few batched queries
    // instead of two or three queries per row.
    const sheetResults = [];
    for (const sheetName of saleSheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
      const headerRowIndex = rows.findIndex((row) => detectSaleTemplate(row));
      if (headerRowIndex < 0) {
        sheetResults.push({ name: sheetName, format: 'tidak dikenali', created: 0, skipped: 0 });
        continue;
      }
      const headers = rows[headerRowIndex];
      const map = headerIndex(headers);
      const format = detectSaleTemplate(headers);
      const sheetStats = { created: 0, skipped: 0 };
      const sheetFailedRows = [];
      const dataRows = rows.slice(headerRowIndex + 1);
      const candidates = [];

      for (let rowIdx = 0; rowIdx < dataRows.length; rowIdx += 1) {
        const row = dataRows[rowIdx];
        const rowNumber = headerRowIndex + rowIdx + 2;
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
          if ((row || []).some((cell) => cell !== null && cell !== undefined && cell !== '')) {
            sheetFailedRows.push({ rowNumber, reason: 'SKU atau Tanggal kosong', raw: row || [] });
          }
          continue;
        }
        const normalized = normalizeImportedSaleRecord(record, summary);
        if (normalized === 'invalid') {
          sheetFailedRows.push({ rowNumber, reason: 'Channel, toko, tanggal, SKU, atau jumlah tidak lengkap', raw: row || [] });
          continue;
        }
        candidates.push(normalized);
      }

      sheetResults.push({ sheetName, format, headers, sheetStats, sheetFailedRows, candidates });
    }

    // Pass 2: ensure every SKU referenced across all sheets has a Product row, in a
    // couple of batched queries.
    const allCandidates = sheetResults.flatMap((item) => item.candidates || []);
    const hppBySku = await ensureImportedProducts(allCandidates, productHpp, summary, businessId, t);

    // Pass 3: fetch pre-import duplicate counts in bulk, then decide create/skip for
    // every row and build the final batch insert.
    const existingCounts = await fetchExistingOrderSkuCounts(allCandidates, businessId, t);
    const occurrenceCounts = new Map();
    const toInsert = [];

    for (const item of sheetResults) {
      if (!item.candidates) continue;
      for (const record of item.candidates) {
        const baseProductHpp = hppBySku.get(record.sku) || 0;
        const totalHpp = record.totalHpp || record.qty * (record.hpp || baseProductHpp || 0);
        const hpp = record.hpp || (record.qty > 0 ? Math.round(totalHpp / record.qty) : 0) || baseProductHpp || 0;

        const duplicateWhere = record.orderNumber
          ? { orderNumber: record.orderNumber, sku: record.sku, businessId }
          : {
            saleDate: record.saleDate,
            channel: record.channel,
            storeName: record.storeName,
            customer: record.customer,
            sku: record.sku,
            qty: record.qty,
            price: record.price,
            subtotal: record.subtotal,
            businessId,
          };
        const duplicateKey = JSON.stringify(duplicateWhere);
        const occurrence = (occurrenceCounts.get(duplicateKey) || 0) + 1;
        occurrenceCounts.set(duplicateKey, occurrence);
        const existingCount = record.orderNumber
          ? (existingCounts.get(JSON.stringify([record.orderNumber, record.sku])) || 0)
          : await Sale.count({ where: duplicateWhere, transaction: t });

        if (existingCount >= occurrence) {
          summary.salesSkipped += 1;
          item.sheetStats.skipped += 1;
          continue;
        }

        toInsert.push({
          businessId,
          saleDate: record.saleDate,
          platformId: null,
          channel: record.channel,
          storeName: record.storeName,
          orderNumber: record.orderNumber,
          customer: record.customer,
          sku: record.sku,
          qty: record.qty,
          price: record.price,
          subtotal: record.subtotal,
          adminFee: record.adminFee,
          hpp,
          totalHpp: totalHpp || record.qty * hpp,
        });
        summary.salesCreated += 1;
        item.sheetStats.created += 1;
      }
    }

    if (toInsert.length) {
      await Sale.bulkCreate(toInsert, { transaction: t });
    }

    for (const item of sheetResults) {
      if (!item.candidates) {
        summary.sheets.push(item);
        continue;
      }
      summary.sheets.push({
        name: item.sheetName,
        format: item.format,
        created: item.sheetStats.created,
        skipped: item.sheetStats.skipped,
      });
      if (item.sheetFailedRows.length) {
        const headerRow = (item.headers || []).map((cell) => excelText(cell));
        sheetFailureGroups.push({ sheetName: item.sheetName, headerRow, failedRows: item.sheetFailedRows });
      }
    }
  });

  const totalFailed = sheetFailureGroups.reduce((sum, group) => sum + group.failedRows.length, 0);
  if (totalFailed) {
    summary.failedRowsCount = totalFailed;
    summary.failedRowsBase64 = buildFailedRowsWorkbook(sheetFailureGroups).toString('base64');
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

function buildFailedRowsWorkbook(sheetGroups) {
  const workbook = XLSX.utils.book_new();
  let appended = 0;
  for (const group of sheetGroups) {
    if (!group.failedRows.length) continue;
    const sheetRows = [['Baris Excel', 'Alasan Gagal', ...group.headerRow]];
    for (const failed of group.failedRows) {
      sheetRows.push([failed.rowNumber, failed.reason, ...failed.raw.map((cell) => (cell === undefined || cell === null ? '' : cell))]);
    }
    const sheet = XLSX.utils.aoa_to_sheet(sheetRows);
    const safeName = String(group.sheetName || 'Data Gagal').slice(0, 28).replace(/[\\/*?:[\]]/g, '_') || 'Data Gagal';
    XLSX.utils.book_append_sheet(workbook, sheet, safeName);
    appended += 1;
  }
  if (!appended) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Baris Excel', 'Alasan Gagal']]), 'Data Gagal');
  }
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

async function importProductsWorkbook(buffer, businessId) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
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
  const failedRows = [];

  const dataRows = rows.slice(headerRowIndex + 1);
  for (let i = 0; i < dataRows.length; i += 1) {
    const row = dataRows[i];
    const rowNumber = headerRowIndex + i + 2;
    const sku = textByHeader(row, map, PRODUCT_SKU_ALIASES);
    const name = textByHeader(row, map, ['Nama Barang']);
    const variant = textByHeader(row, map, PRODUCT_VARIANT_ALIASES);
    const hpp = numberByHeader(row, map, ['HPP']);
    if (!sku || !name) {
      if ((row || []).some((cell) => cell !== null && cell !== undefined && cell !== '')) {
        failedRows.push({ rowNumber, reason: 'SKU atau Nama Barang kosong', raw: row || [] });
      }
      continue;
    }

    // Satu baris bermasalah tidak boleh membatalkan seluruh import: baris yang
    // sudah tersimpan tetap dihitung dan baris gagal dilaporkan ke pengguna.
    try {
      const [product, isCreated] = await Product.findOrCreate({
        where: { sku, businessId },
        defaults: {
          businessId,
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
    } catch (error) {
      failedRows.push({ rowNumber, reason: error.message || 'Baris gagal disimpan', raw: row || [] });
    }
  }

  summary.sheets.push({
    name: sheetName,
    created,
    updated,
  });

  if (failedRows.length) {
    const headerRow = (rows[headerRowIndex] || []).map((cell) => excelText(cell));
    summary.failedRowsCount = failedRows.length;
    summary.failedRowsBase64 = buildFailedRowsWorkbook([{ sheetName: 'Data Gagal', headerRow, failedRows }]).toString('base64');
  }

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

async function importMarketingWorkbook(buffer, businessId) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
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
  const failedRows = [];

  const dataRows = rows.slice(headerRowIndex + 1);
  for (let i = 0; i < dataRows.length; i += 1) {
    const row = dataRows[i];
    const rowNumber = headerRowIndex + i + 2;
    const expenseDate = dateByHeader(row, map, ['Tanggal']);
    const category = textByHeader(row, map, ['Kategori']);
    const platformStore = textByHeader(row, map, ['Platform & Toko']);
    const description = textByHeader(row, map, ['Keterangan']);
    const nominal = numberByHeader(row, map, ['Nominal']);
    const taxCell = cellByHeader(row, map, ['Pajak Iklan']);
    const taxCellBlank = taxCell === undefined || taxCell === null || taxCell === '';
    let taxRate = excelPercent(taxCell);
    if (taxCellBlank && platformStore) {
      const matchingStore = await Store.findOne({ where: { platformStore, businessId } });
      if (matchingStore) {
        taxRate = matchingStore.taxRate || 0;
      }
    }
    const totalTax = numberByHeader(row, map, ['Total Pajak']) || Math.round(nominal * taxRate);

    if (!expenseDate || !category || !platformStore) {
      if ((row || []).some((cell) => cell !== null && cell !== undefined && cell !== '')) {
        failedRows.push({ rowNumber, reason: 'Tanggal, kategori, atau platform & toko kosong', raw: row || [] });
      }
      continue;
    }

    const duplicate = await MarketingExpense.findOne({
      where: {
        businessId,
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
      businessId,
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

  if (failedRows.length) {
    const headerRow = (rows[headerRowIndex] || []).map((cell) => excelText(cell));
    summary.failedRowsCount = failedRows.length;
    summary.failedRowsBase64 = buildFailedRowsWorkbook([{ sheetName: 'Data Gagal', headerRow, failedRows }]).toString('base64');
  }

  return summary;
}

// Scoped to a single business: deletes only that business's data, never touches the
// other business, shared reference data (expense categories, holidays), or owner
// accounts (businessId IS NULL). Owner accounts and reference data survive a reset.
async function resetDatabase(businessId) {
  await sequelize.query(
    'DELETE FROM `attendance` WHERE `user_id` IN (SELECT `id` FROM `users` WHERE `business_id` = :businessId)',
    { replacements: { businessId } },
  );
  await sequelize.query(
    'DELETE FROM `attendance` WHERE `user_id` IN (SELECT `user_id` FROM `owner_salary_settings` WHERE `business_id` = :businessId)',
    { replacements: { businessId } },
  );
  await SalaryPayment.destroy({ where: { businessId } });
  await OwnerSalarySetting.destroy({ where: { businessId } });
  const scopedTables = ['sales', 'marketing_expenses', 'operational_expenses', 'stores', 'products', 'shifts', 'attendance_settings', 'profit_share_partners'];
  for (const tableName of scopedTables) {
    await sequelize.query(`DELETE FROM \`${tableName}\` WHERE \`business_id\` = :businessId`, { replacements: { businessId } });
  }
  await sequelize.query('DELETE FROM `users` WHERE `business_id` = :businessId', { replacements: { businessId } });
  await getAttendanceSettings(businessId);
}

createSimpleCrudRoutes('shifts', Shift, serializeShift, parseShift, {});
createSimpleCrudRoutes('stores', Store, serializeStore, parseStore, {
  include: [Platform],
  reload: true,
  searchFields: ['platform', 'name', 'platformStore'],
});
createSimpleCrudRoutes('products', Product, serializeProduct, parseProduct, {
  paginated: true,
  searchFields: ['sku', 'name', 'variant'],
  sortFields: ['sku', 'name', 'variant', 'hpp', 'createdAt'],
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
    if (!requireConcreteBusiness(req, res)) return;
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'File Excel wajib diunggah.' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) {
      return res.status(400).json({ ok: false, message: 'Format file harus .xlsx atau .xls.' });
    }
    const summary = await importProductsWorkbook(req.file.buffer, req.businessId);
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


async function fetchGroupedSales({ search, channel, storeName, from, to, sortBy, sortDir, limit, offset, businessId }) {
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
  if (channel) {
    conditions.push('`channel` = :channel');
    replacements.channel = channel;
  }
  if (from) {
    conditions.push('`sale_date` >= :from');
    replacements.from = from;
  }
  if (to) {
    conditions.push('`sale_date` <= :to');
    replacements.to = to;
  }
  if (businessId !== null && businessId !== undefined) {
    conditions.push('`business_id` = :businessId');
    replacements.businessId = businessId;
  }
  const whereSql = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';
  const sortColumns = {
    saleDate: 'saleDate',
    channel: 'channel',
    storeName: 'storeName',
    orderNumber: 'orderNumber',
    customer: 'customer',
    itemCount: 'itemCount',
    totalQty: 'totalQty',
    totalSubtotal: 'totalSubtotal',
    totalAdminFee: 'totalAdminFee',
    totalHpp: 'totalHpp',
    profit: 'profit',
  };
  const orderColumn = sortColumns[sortBy] || sortColumns.saleDate;
  const orderDirection = String(sortDir).toLowerCase() === 'asc' ? 'ASC' : 'DESC';

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
       SUM(\`admin_fee\`) AS totalAdminFee,
       SUM(\`total_hpp\`) AS totalHpp,
       SUM(\`subtotal\`) - SUM(\`admin_fee\`) - SUM(\`total_hpp\`) AS profit
     FROM \`sales\`
     ${whereSql}
     GROUP BY ${SALE_GROUP_KEY_SQL}
     ORDER BY ${orderColumn} ${orderDirection}, id DESC
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
  totalAdminFee: toInt(row.totalAdminFee, 0),
  ...(includeHpp ? { totalHpp: toInt(row.totalHpp, 0) } : {}),
});

app.get('/api/sales', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const includeHpp = req.user.role === 'owner';
    const { page, limit, offset } = getPagination(req.query);
    const searchText = toText(req.query.search);
    const storeId = toInt(req.query.storeId, null);
    const selectedStore = storeId ? await Store.findByPk(storeId) : null;
    if (storeId && (!selectedStore || !assertOwnedByBusiness(selectedStore, req))) {
      return res.status(400).json({ ok: false, message: 'Toko yang dipilih tidak ditemukan.' });
    }
    const storeName = selectedStore ? selectedStore.name : toText(req.query.storeName);
    const channel = selectedStore ? selectedStore.platform : toText(req.query.channel);
    const from = toText(req.query.from);
    const to = toText(req.query.to);
    const sortBy = toText(req.query.sortBy);
    const sortDir = toText(req.query.sortDir);

    if (toText(req.query.grouped) === 'true') {
      const { rows, total } = await fetchGroupedSales({ search: searchText, channel, storeName, from, to, sortBy, sortDir, limit, offset, businessId: req.businessId });
      return res.json({
        ok: true,
        data: rows.map((row) => serializeGroupedSale(row, { includeHpp })),
        pagination: paginationMeta(page, limit, total),
      });
    }

    const where = businessWhere(req, searchWhere(searchText, ['orderNumber', 'customer', 'sku', 'storeName', 'channel']));
    if (storeName) {
      where.storeName = storeName;
    }
    if (channel) {
      where.channel = channel;
    }
    if (from || to) {
      where.saleDate = {};
      if (from) where.saleDate[Op.gte] = from;
      if (to) where.saleDate[Op.lte] = to;
    }

    const saleSortFields = ['saleDate', 'channel', 'storeName', 'orderNumber', 'customer', 'qty', 'subtotal', 'adminFee', 'totalHpp'];
    const saleSortBy = saleSortFields.includes(sortBy) ? sortBy : 'saleDate';
    const saleSortDir = sortDir.toLowerCase() === 'asc' ? 'ASC' : 'DESC';
    const { count, rows } = await Sale.findAndCountAll({
      where,
      include: [Platform],
      order: [[saleSortBy, saleSortDir], ['createdAt', 'DESC']],
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
    const where = businessWhere(req, orderNumber ? { orderNumber } : { id });
    const rows = await Sale.findAll({ where, include: [Platform], order: [['id', 'ASC']] });
    return res.json({ ok: true, data: rows.map((row) => serializeSale(row, { includeHpp })) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat detail penjualan.' });
  }
});

app.post('/api/sales', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const payload = await parseSale(req.body, null, req.businessId);
    const row = await Sale.create({ ...payload, businessId: req.businessId });
    const output = await Sale.findByPk(row.id, { include: [Platform] });
    return res.status(201).json({ ok: true, data: serializeSale(output, { includeHpp: req.user.role === 'owner' }) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan penjualan.' });
  }
});

app.post('/api/sales/import-excel', ...authRequired, salesRecapAccess, upload.single('file'), async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'File Excel wajib diunggah.' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) {
      return res.status(400).json({ ok: false, message: 'Format file harus .xlsx atau .xls.' });
    }
    const summary = await importSalesWorkbook(req.file.buffer, req.file.originalname, req.businessId);
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
    if (!row || !assertOwnedByBusiness(row, req)) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    const payload = await parseSale(req.body, row, row.businessId);
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
    if (!row || !assertOwnedByBusiness(row, req)) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Penjualan berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus penjualan.' });
  }
});

app.post('/api/sales/bulk-delete', ...authRequired, salesRecapAccess, async (req, res) => {
  try {
    const items = Array.isArray(req.body.items) ? req.body.items : [];
    if (!items.length) {
      return res.status(400).json({ ok: false, message: 'Pilih minimal satu data untuk dihapus.' });
    }
    let deleted = 0;
    for (const item of items) {
      const orderNumber = toText(item?.orderNumber);
      if (orderNumber) {
        deleted += await Sale.destroy({ where: businessWhere(req, { orderNumber }) });
      } else {
        const id = toInt(item?.id, null);
        if (id !== null) {
          deleted += await Sale.destroy({ where: businessWhere(req, { id }) });
        }
      }
    }
    return res.json({ ok: true, message: `${deleted} baris penjualan berhasil dihapus.`, data: { deleted } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus penjualan terpilih.' });
  }
});
createSimpleCrudRoutes('expenses/operasional', OperationalExpense, serializeOperational, parseOperational, {
  writeMiddleware: ownerOrLeader,
  searchFields: ['category', 'description'],
  dateField: 'expenseDate',
});
createSimpleCrudRoutes('expenses/marketing', MarketingExpense, serializeMarketing, parseMarketing, {
  include: [Platform],
  reload: true,
  writeMiddleware: ownerLeaderAdmin,
  searchFields: ['category', 'platformStore', 'description'],
  dateField: 'expenseDate',
});

app.get('/api/expense-categories', ...authRequired, async (req, res) => {
  try {
    const kind = toText(req.query.kind);
    const where = businessWhere(req, kind ? { kind } : {});
    const rows = await ExpenseCategory.findAll({ where, order: [['name', 'ASC']] });
    return res.json({ ok: true, data: rows.map(serializeExpenseCategory) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat kategori.' });
  }
});

app.post('/api/expense-categories', ...authRequired, ownerLeaderAdmin, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const name = toText(req.body.name);
    const kind = toText(req.body.kind);
    if (!name || !['operasional', 'marketing'].includes(kind)) {
      return res.status(400).json({ ok: false, message: 'Nama dan jenis kategori wajib diisi.' });
    }
    const [row] = await ExpenseCategory.findOrCreate({
      where: { name, kind, businessId: req.businessId },
      defaults: { name, kind, businessId: req.businessId },
    });
    return res.status(201).json({ ok: true, data: serializeExpenseCategory(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah kategori.' });
  }
});

app.get('/api/expenses/marketing/template', ...authRequired, ownerLeaderAdmin, async (req, res) => {
  try {
    const buffer = buildMarketingTemplateBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="template_beban_marketing.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal membuat template marketing.' });
  }
});

app.post('/api/expenses/marketing/import-excel', ...authRequired, ownerLeaderAdmin, upload.single('file'), async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    if (!req.file) {
      return res.status(400).json({ ok: false, message: 'File Excel wajib diunggah.' });
    }
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) {
      return res.status(400).json({ ok: false, message: 'Format file harus .xlsx atau .xls.' });
    }
    const summary = await importMarketingWorkbook(req.file.buffer, req.businessId);
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
    if (!requireConcreteBusiness(req, res)) return;
    const password = toText(req.body.password);
    if (!password) {
      return res.status(400).json({ ok: false, message: 'Password admin wajib diisi.' });
    }
    const valid = await bcrypt.compare(password, req.user.passwordHash);
    if (!valid) {
      return res.status(401).json({ ok: false, message: 'Password admin tidak sesuai.' });
    }
    await resetDatabase(req.businessId);
    return res.json({
      ok: true,
      message: 'Data bisnis yang sedang aktif berhasil direset.',
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal mereset database.' });
  }
});

async function availableBusinessesFor(user) {
  if (user.role === 'owner') {
    const rows = await Business.findAll({ where: { active: true }, order: [['id', 'ASC']] });
    return rows.map(serializeBusiness);
  }
  if (!user.businessId) return [];
  const row = await Business.findByPk(user.businessId);
  return row ? [serializeBusiness(row)] : [];
}

app.get('/api/businesses', ...authRequired, async (req, res) => {
  try {
    const data = await availableBusinessesFor(req.user);
    return res.json({ ok: true, data });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat daftar bisnis.' });
  }
});

app.post('/api/businesses', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const name = toText(req.body.name);
    const slug = toText(req.body.slug) || name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
    if (!name || !slug) {
      return res.status(400).json({ ok: false, message: 'Nama bisnis wajib diisi.' });
    }
    const row = await Business.create({ name, slug, active: true });
    return res.status(201).json({ ok: true, data: serializeBusiness(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah bisnis.' });
  }
});

app.put('/api/businesses/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await Business.findByPk(req.params.id);
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Bisnis tidak ditemukan.' });
    }
    if (req.body.name !== undefined) row.name = toText(req.body.name);
    if (req.body.active !== undefined) row.active = Boolean(req.body.active);
    await row.save();
    return res.json({ ok: true, data: serializeBusiness(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal memperbarui bisnis.' });
  }
});

app.get('/api/auth/me', ...authRequired, async (req, res) => {
  const availableBusinesses = await availableBusinessesFor(req.user);
  return res.json({ ok: true, data: serializeUser(req.user), availableBusinesses });
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
    const availableBusinesses = await availableBusinessesFor(user);
    return res.json({
      ok: true,
      token,
      data: serializeUser(user),
      availableBusinesses,
      message: 'Masuk berhasil.',
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal masuk ke sistem.' });
  }
});

app.get('/api/users', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const where = businessOrOwnerWhere(req, searchWhere(req.query.search, ['fullName', 'username', 'jobTitle']));
    const rows = await User.findAll({ where, order: [['createdAt', 'DESC']], include: [Shift] });
    const ownerSettings = await ownerSalarySettingsByUser(rows, req.businessId);
    const scopedRows = (req.businessId === null || req.businessId === undefined)
      ? rows
      : rows.filter((row) => row.role !== 'owner' || ownerSettings.has(row.id));
    return res.json({ ok: true, data: scopedRows.map((row) => serializeUser(row, ownerSettings.get(row.id))) });
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
    const role = ROLES.includes(req.body.role) ? req.body.role : 'karyawan';
    let businessId = null;
    let salaryBusinessId = null;
    if (role === 'owner') {
      salaryBusinessId = req.body.businessId !== undefined
        ? toInt(req.body.businessId, null)
        : req.businessId;
      if (!salaryBusinessId || !(await Business.findByPk(salaryBusinessId))) {
        return res.status(400).json({ ok: false, message: 'Bisnis gaji owner wajib dipilih.' });
      }
    } else {
      businessId = req.body.businessId !== undefined ? toInt(req.body.businessId, null) : req.businessId;
      if (!businessId) {
        return res.status(400).json({ ok: false, message: 'Bisnis wajib dipilih untuk role selain owner.' });
      }
    }
    const exists = await User.findOne({ where: { username } });
    if (exists) {
      return res.status(409).json({ ok: false, message: 'Username sudah dipakai.' });
    }
    const passwordHash = await bcrypt.hash(password || '12345678', 10);
    const row = await User.create({
      businessId,
      username,
      passwordHash,
      fullName,
      role,
      jobTitle: toText(req.body.jobTitle),
      workShift: toText(req.body.workShift),
      phone: toText(req.body.phone),
      dailyWage: role === 'owner' ? 0 : money(req.body.dailyWage),
      mealAllowance: role === 'owner' ? 0 : money(req.body.mealAllowance),
      overtimeRatePerHour: role === 'owner' ? 0 : money(req.body.overtimeRatePerHour),
      shiftId: req.body.shiftId ? toInt(req.body.shiftId, null) : null,
      active: req.body.active === undefined ? true : Boolean(req.body.active),
    });
    const output = await User.findByPk(row.id, { include: [Shift] });
    let salarySetting = null;
    if (role === 'owner') {
      salarySetting = await upsertOwnerSalarySetting(output, salaryBusinessId, req.body);
    }
    return res.status(201).json({ ok: true, data: serializeUser(output, salarySetting) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah karyawan.' });
  }
});

app.put('/api/users/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await User.findByPk(req.params.id);
    if (!assertUserInBusinessScope(row, req)) {
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
    let salaryBusinessId = null;
    if (row.role === 'owner') {
      salaryBusinessId = req.body.businessId !== undefined
        ? toInt(req.body.businessId, null)
        : req.businessId;
      if (!salaryBusinessId || !(await Business.findByPk(salaryBusinessId))) {
        return res.status(400).json({ ok: false, message: 'Bisnis gaji owner wajib dipilih.' });
      }
      row.businessId = null;
    } else if (req.body.businessId !== undefined) {
      row.businessId = toInt(req.body.businessId, row.businessId);
    }
    row.jobTitle = toText(req.body.jobTitle);
    row.workShift = toText(req.body.workShift);
    row.phone = toText(req.body.phone);
    if (row.role !== 'owner') {
      row.dailyWage = money(req.body.dailyWage);
      row.mealAllowance = money(req.body.mealAllowance);
      row.overtimeRatePerHour = money(req.body.overtimeRatePerHour);
    }
    row.shiftId = req.body.shiftId ? toInt(req.body.shiftId, null) : null;
    if (req.body.active !== undefined) {
      row.active = Boolean(req.body.active);
    }
    await row.save();
    const output = await User.findByPk(row.id, { include: [Shift] });
    const salarySetting = row.role === 'owner'
      ? await upsertOwnerSalarySetting(output, salaryBusinessId, req.body)
      : null;
    return res.json({ ok: true, data: serializeUser(output, salarySetting) });
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
    if (!assertUserInBusinessScope(row, req)) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Karyawan berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus karyawan.' });
  }
});

app.post('/api/users/bulk-delete', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const ids = Array.isArray(req.body.ids)
      ? req.body.ids.map((id) => toInt(id, null)).filter((id) => id !== null && id !== req.user.id)
      : [];
    if (!ids.length) {
      return res.status(400).json({ ok: false, message: 'Pilih minimal satu karyawan untuk dihapus.' });
    }
    const deleted = await User.destroy({ where: businessOrOwnerWhere(req, { id: { [Op.in]: ids } }) });
    return res.json({ ok: true, message: `${deleted} karyawan berhasil dihapus.`, data: { deleted } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus karyawan terpilih.' });
  }
});

app.get('/api/face/me', ...authRequired, async (req, res) => {
  if (!['owner', 'leader', 'karyawan', 'admin'].includes(req.user.role)) {
    return res.status(403).json({ ok: false, message: 'Daftar wajah tidak tersedia untuk role Anda.' });
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
    if (!['owner', 'leader', 'karyawan', 'admin'].includes(req.user.role)) {
      return res.status(403).json({ ok: false, message: 'Daftar wajah tidak tersedia untuk role Anda.' });
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
    if (!(await isUserInAttendanceBusiness(req.user, req.businessId))) {
      return res.json({ ok: true, data: [] });
    }
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
    const search = toText(req.query.search);
    const from = toText(req.query.from);
    const to = toText(req.query.to);
    const where = { ...dateWhere('workDate', from, to) };
    if (search) {
      where[Op.and] = [sequelize.where(sequelize.col('User.full_name'), { [Op.like]: `%${search}%` })];
    }
    const rows = await Attendance.findAll({
      where,
      include: [{ model: User, where: await attendanceUserScopeWhere(req) }],
      order: [['workDate', 'DESC'], ['createdAt', 'DESC']],
      limit: 400,
      subQuery: false,
    });
    return res.json({ ok: true, data: rows.map(serializeAttendance) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat absensi.' });
  }
});

app.get('/api/attendance/settings', ...authRequired, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const settings = await getAttendanceSettings(req.businessId);
    return res.json({ ok: true, data: serializeAttendanceSettings(settings) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat pengaturan absensi.' });
  }
});

app.put('/api/attendance/settings', ...authRequired, ownerOnly, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
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
    const settings = await getAttendanceSettings(req.businessId);
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
      where: await attendanceUserScopeWhere(req, { active: true }),
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
    const target = await User.findByPk(userId, { include: [{ model: Shift }] });
    if (!target || !(await isUserInAttendanceBusiness(target, req.businessId))) {
      return res.status(404).json({ ok: false, message: 'Karyawan tidak ditemukan.' });
    }
    const overtimeHours = target.role === 'owner'
      ? 0
      : (req.body.overtimeHours !== undefined && toText(req.body.overtimeHours) !== ''
        ? Math.max(0, toFloat(req.body.overtimeHours, 0))
        : await resolveManualOvertimeHours(target, workDate, status, req.businessId));
    let row = await Attendance.findOne({ where: { userId, workDate } });
    if (row) {
      row.status = status;
      row.note = toText(req.body.note) || row.note;
      row.overtimeHours = overtimeHours;
      await row.save();
    } else {
      row = await Attendance.create({
        userId,
        workDate,
        status,
        method: 'manual',
        note: toText(req.body.note),
        overtimeHours,
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
    if (!row || !(await isUserInAttendanceBusiness(row.User, req.businessId))) {
      return res.status(404).json({ ok: false, message: 'Data absensi tidak ditemukan.' });
    }
    if (req.body.status !== undefined) {
      const status = toText(req.body.status);
      if (!ATTENDANCE_STATUSES.includes(status)) {
        return res.status(400).json({ ok: false, message: 'Status tidak valid.' });
      }
      row.status = status;
    }
    if (row.status !== 'hadir' || (row.User && row.User.role === 'owner')) {
      row.overtimeHours = 0;
    } else if (req.body.overtimeHours !== undefined && toText(req.body.overtimeHours) !== '') {
      row.overtimeHours = Math.max(0, toFloat(req.body.overtimeHours, 0));
    } else if (req.body.status !== undefined) {
      const target = await User.findByPk(row.userId, { include: [{ model: Shift }] });
      row.overtimeHours = await resolveManualOvertimeHours(target, row.workDate, row.status, req.businessId);
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
    const row = await Attendance.findByPk(req.params.id, { include: [{ model: User }] });
    if (!row || !(await isUserInAttendanceBusiness(row.User, req.businessId))) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Data absensi berhasil dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus data absensi.' });
  }
});

app.post('/api/attendance/bulk-delete', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    const ids = Array.isArray(req.body.ids) ? req.body.ids.map((id) => toInt(id, null)).filter((id) => id !== null) : [];
    if (!ids.length) {
      return res.status(400).json({ ok: false, message: 'Pilih minimal satu data absensi untuk dihapus.' });
    }
    const scopedRows = await Attendance.findAll({
      where: { id: { [Op.in]: ids } },
      include: [{ model: User, where: await attendanceUserScopeWhere(req) }],
      attributes: ['id'],
    });
    const scopedIds = scopedRows.map((row) => row.id);
    const deleted = scopedIds.length ? await Attendance.destroy({ where: { id: { [Op.in]: scopedIds } } }) : 0;
    return res.json({ ok: true, message: `${deleted} data absensi berhasil dihapus.`, data: { deleted } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus data absensi terpilih.' });
  }
});

app.post('/api/attendance/check-in', ...authRequired, async (req, res) => {
  try {
    if (!['owner', 'leader', 'karyawan', 'admin'].includes(req.user.role)) {
      return res.status(403).json({ ok: false, message: 'Absensi wajah tidak tersedia untuk role Anda.' });
    }
    if (!(await isUserInAttendanceBusiness(req.user, req.businessId))) {
      return res.status(403).json({ ok: false, message: 'Akun Anda tidak ditugaskan ke absensi bisnis aktif.' });
    }
    const photoUrl = toText(req.body.photoUrl);
    if (!photoUrl) {
      return res.status(400).json({ ok: false, message: 'Foto absensi wajib diambil dari kamera terlebih dahulu.' });
    }
    const locationName = toText(req.body.locationName);
    const settings = await getAttendanceSettings(req.businessId);
    let shiftWindow = resolveShiftWindow(req.user, settings);
    if (req.body.shiftId) {
      const overrideShift = await Shift.findByPk(toInt(req.body.shiftId, null));
      if (overrideShift && assertOwnedByBusiness(overrideShift, req)) {
        shiftWindow = {
          checkInDeadline: overrideShift.checkInDeadline,
          lateToleranceMinutes: overrideShift.lateToleranceMinutes,
          name: overrideShift.name,
        };
      }
    }
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
      if (row.checkInAt && req.user.role !== 'owner' && await isOvertimeDay(row.workDate)) {
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

function salaryScopeWhere(req) {
  const requester = req.user;
  if (requester.role === 'karyawan' || requester.role === 'admin') {
    return { id: requester.id, active: true };
  }
  if (requester.role === 'leader') {
    return businessWhere(req, { role: { [Op.ne]: 'owner' }, active: true });
  }
  return businessOrOwnerWhere(req, { active: true });
}

async function computeSalaryReport(weekStart, scopeWhere, businessId) {
  const { start, end } = weekBounds(weekStart);
  const candidateUsers = await User.findAll({ where: scopeWhere, order: [['fullName', 'ASC']] });
  const ownerSettings = await ownerSalarySettingsByUser(candidateUsers, businessId);
  const users = candidateUsers.filter((user) => user.role !== 'owner' || ownerSettings.has(user.id));
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
    where: { weekStart, businessId, userId: { [Op.in]: users.map((user) => user.id) } },
  });
  const paymentsByUser = new Map(payments.map((payment) => [payment.userId, payment]));
  const overtimeDates = await overtimeDateSet(attendanceRows.map((row) => row.workDate));
  const data = users.map((user) => {
    const records = grouped.get(user.id) || [];
    const presentDays = records.filter((record) => record.status === 'hadir').length;
    return buildSalaryRow(user, presentDays, records, paymentsByUser.get(user.id), ownerSettings.get(user.id), overtimeDates);
  });
  return { data, start, end };
}

async function findOrCreateSalaryPayment(userId, weekStart, businessId) {
  const [payment] = await SalaryPayment.findOrCreate({
    where: { userId, weekStart, businessId },
    defaults: { userId, weekStart, businessId, paid: false },
  });
  return payment;
}

app.get('/api/reports/salary', ...authRequired, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const scopeWhere = salaryScopeWhere(req);
    if (!scopeWhere) {
      return res.status(403).json({ ok: false, message: 'Akses laporan gaji tidak tersedia untuk role Anda.' });
    }
    const weekStart = sundayOfWeek(toText(req.query.week) || jakartaDate());
    const { data, start, end } = await computeSalaryReport(weekStart, scopeWhere, req.businessId);
    return res.json({ ok: true, data, period: { weekStart, start, end } });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat laporan gaji.' });
  }
});

app.post('/api/reports/salary/adjust', ...authRequired, ownerOrLeader, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const userId = toInt(req.body.userId, null);
    if (!userId) {
      return res.status(400).json({ ok: false, message: 'Karyawan tidak valid.' });
    }
    const target = await User.findByPk(userId);
    if (!target || !assertUserInBusinessScope(target, req)) {
      return res.status(404).json({ ok: false, message: 'Karyawan tidak ditemukan.' });
    }
    if (target.role === 'owner' && req.user.role !== 'owner') {
      return res.status(403).json({ ok: false, message: 'Gaji owner tidak bisa diubah oleh kepala toko.' });
    }
    if (target.role === 'owner' && !(await OwnerSalarySetting.findOne({
      where: { userId: target.id, businessId: req.businessId },
    }))) {
      return res.status(404).json({ ok: false, message: 'Owner tidak ditugaskan ke laporan gaji bisnis ini.' });
    }
    const weekStart = sundayOfWeek(toText(req.body.weekStart) || jakartaDate());
    const payment = await findOrCreateSalaryPayment(userId, weekStart, req.businessId);
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
    if (!requireConcreteBusiness(req, res)) return;
    const userId = toInt(req.body.userId, null);
    if (!userId) {
      return res.status(400).json({ ok: false, message: 'Karyawan tidak valid.' });
    }
    const target = await User.findByPk(userId);
    if (!target || !assertUserInBusinessScope(target, req)) {
      return res.status(404).json({ ok: false, message: 'Karyawan tidak ditemukan.' });
    }
    if (target.role === 'owner' && req.user.role !== 'owner') {
      return res.status(403).json({ ok: false, message: 'Gaji owner tidak bisa diubah oleh kepala toko.' });
    }
    if (target.role === 'owner' && !(await OwnerSalarySetting.findOne({
      where: { userId: target.id, businessId: req.businessId },
    }))) {
      return res.status(404).json({ ok: false, message: 'Owner tidak ditugaskan ke laporan gaji bisnis ini.' });
    }
    const weekStart = sundayOfWeek(toText(req.body.weekStart) || jakartaDate());
    const paid = Boolean(req.body.paid);
    const payment = await findOrCreateSalaryPayment(userId, weekStart, req.businessId);
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
    if (!requireConcreteBusiness(req, res)) return;
    const weekStart = sundayOfWeek(toText(req.body.weekStart) || jakartaDate());
    const paid = req.body.paid === undefined ? true : Boolean(req.body.paid);
    const scopeWhere = salaryScopeWhere(req) || { active: true };
    const candidateUsers = await User.findAll({ where: scopeWhere });
    const ownerSettings = await ownerSalarySettingsByUser(candidateUsers, req.businessId);
    const users = candidateUsers.filter((user) => user.role !== 'owner' || ownerSettings.has(user.id));
    for (const user of users) {
      const payment = await findOrCreateSalaryPayment(user.id, weekStart, req.businessId);
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
    const data = await buildProfitSummary(from, to, {}, req.businessId);
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
    const data = await buildProfitSummary(from, to, {}, req.businessId);
    const buffer = await buildProfitReportPdf(data);
    res.setHeader('Content-Disposition', 'attachment; filename="laba_rugi.pdf"');
    res.setHeader('Content-Type', 'application/pdf');
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal membuat PDF laba rugi.' });
  }
});

function monthRange(period) {
  const [y, m] = period.split('-').map(Number);
  const from = `${period}-01`;
  const to = `${period}-${pad(new Date(y, m, 0).getDate())}`;
  return { from, to };
}

function formatPeriodMonthLabel(period) {
  const [y, m] = period.split('-').map(Number);
  return new Intl.DateTimeFormat('id-ID', { month: 'long', year: 'numeric' }).format(new Date(y, m - 1, 1));
}

app.get('/api/profit-share/partners', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const rows = await ProfitSharePartner.findAll({
      where: businessWhere(req),
      order: [['sortOrder', 'ASC'], ['name', 'ASC']],
    });
    return res.json({ ok: true, data: rows.map(serializeProfitSharePartner) });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat data pembagian keuntungan.' });
  }
});

app.post('/api/profit-share/partners', ...authRequired, ownerOnly, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const name = toText(req.body.name);
    const percentage = toFloat(req.body.percentage, 0);
    if (!name) {
      return res.status(400).json({ ok: false, message: 'Nama pihak wajib diisi.' });
    }
    if (percentage < 0 || percentage > 100) {
      return res.status(400).json({ ok: false, message: 'Persentase harus di antara 0 - 100.' });
    }
    const row = await ProfitSharePartner.create({ name, percentage, businessId: req.businessId });
    return res.status(201).json({ ok: true, data: serializeProfitSharePartner(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menambah pihak pembagian keuntungan.' });
  }
});

app.put('/api/profit-share/partners/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await ProfitSharePartner.findByPk(toInt(req.params.id, null));
    if (!row || !assertOwnedByBusiness(row, req)) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    const name = toText(req.body.name);
    const percentage = toFloat(req.body.percentage, row.percentage);
    if (!name) {
      return res.status(400).json({ ok: false, message: 'Nama pihak wajib diisi.' });
    }
    if (percentage < 0 || percentage > 100) {
      return res.status(400).json({ ok: false, message: 'Persentase harus di antara 0 - 100.' });
    }
    row.name = name;
    row.percentage = percentage;
    if (req.body.active !== undefined) row.active = Boolean(req.body.active);
    await row.save();
    return res.json({ ok: true, data: serializeProfitSharePartner(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan perubahan.' });
  }
});

app.delete('/api/profit-share/partners/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await ProfitSharePartner.findByPk(toInt(req.params.id, null));
    if (!row || !assertOwnedByBusiness(row, req)) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await ProfitShareInstallment.destroy({ where: { partnerId: row.id } });
    await ProfitShareTarget.destroy({ where: { partnerId: row.id } });
    await row.destroy();
    return res.json({ ok: true, message: 'Pihak pembagian keuntungan dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus data.' });
  }
});

app.get('/api/profit-share/summary', ...authRequired, ownerOnly, async (req, res) => {
  try {
    if (!requireConcreteBusiness(req, res)) return;
    const period = /^\d{4}-\d{2}$/.test(toText(req.query.period)) ? toText(req.query.period) : jakartaDate().slice(0, 7);
    const { from, to } = monthRange(period);

    const partners = await ProfitSharePartner.findAll({
      where: businessWhere(req, { active: true }),
      order: [['sortOrder', 'ASC'], ['name', 'ASC']],
    });
    const partnerIds = partners.map((p) => p.id);

    const profit = await buildProfitSummary(from, to, {}, req.businessId);
    const netProfit = profit.totals.netProfit;

    const targets = partnerIds.length
      ? await ProfitShareTarget.findAll({ where: { period, partnerId: { [Op.in]: partnerIds } } })
      : [];
    const targetByPartner = new Map(targets.map((t) => [t.partnerId, t]));

    const installments = partnerIds.length
      ? await ProfitShareInstallment.findAll({ where: { period, partnerId: { [Op.in]: partnerIds } }, order: [['paidDate', 'ASC'], ['id', 'ASC']] })
      : [];
    const installmentsByPartner = new Map();
    for (const row of installments) {
      if (!installmentsByPartner.has(row.partnerId)) installmentsByPartner.set(row.partnerId, []);
      installmentsByPartner.get(row.partnerId).push(serializeProfitShareInstallment(row));
    }

    let percentageSum = 0;
    let targetSum = 0;
    let paidSum = 0;
    let remainingSum = 0;

    const partnerRows = partners.map((partner) => {
      const autoShare = Math.round(netProfit * (partner.percentage / 100));
      const targetRow = targetByPartner.get(partner.id);
      const target = targetRow ? targetRow.amount : autoShare;
      const partnerInstallments = installmentsByPartner.get(partner.id) || [];
      const totalPaid = partnerInstallments.reduce((sum, item) => sum + item.amount, 0);
      const remaining = target - totalPaid;
      percentageSum += partner.percentage;
      targetSum += target;
      paidSum += totalPaid;
      remainingSum += remaining;
      return {
        ...serializeProfitSharePartner(partner),
        autoShare,
        target,
        isOverride: Boolean(targetRow),
        totalPaid,
        remaining,
        status: remaining <= 0 ? 'TUNTAS' : 'BELUM TUNTAS',
        installments: partnerInstallments,
      };
    });

    return res.json({
      ok: true,
      data: {
        period,
        periodLabel: formatPeriodMonthLabel(period),
        netProfit,
        percentageSum,
        partners: partnerRows,
        totals: { targetSum, paidSum, remainingSum },
      },
    });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal memuat pembagian keuntungan.' });
  }
});

app.put('/api/profit-share/target', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const partnerId = toInt(req.body.partnerId, null);
    const period = toText(req.body.period);
    if (!partnerId || !/^\d{4}-\d{2}$/.test(period)) {
      return res.status(400).json({ ok: false, message: 'Pihak dan periode wajib valid.' });
    }
    const partner = await ProfitSharePartner.findByPk(partnerId);
    if (!partner || !assertOwnedByBusiness(partner, req)) {
      return res.status(404).json({ ok: false, message: 'Pihak tidak ditemukan.' });
    }
    if (req.body.amount === null || req.body.amount === '') {
      await ProfitShareTarget.destroy({ where: { partnerId, period } });
      return res.json({ ok: true, message: 'Target dikembalikan ke perhitungan otomatis.' });
    }
    const amount = toInt(req.body.amount, 0);
    const [row] = await ProfitShareTarget.findOrCreate({ where: { partnerId, period }, defaults: { partnerId, period, amount } });
    row.amount = amount;
    await row.save();
    return res.json({ ok: true, message: 'Target pembulatan tersimpan.' });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan target.' });
  }
});

app.post('/api/profit-share/installments', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const partnerId = toInt(req.body.partnerId, null);
    const period = toText(req.body.period);
    const paidDate = toText(req.body.paidDate) || jakartaDate();
    const amount = toInt(req.body.amount, 0);
    if (!partnerId || !/^\d{4}-\d{2}$/.test(period)) {
      return res.status(400).json({ ok: false, message: 'Pihak dan periode wajib valid.' });
    }
    if (amount <= 0) {
      return res.status(400).json({ ok: false, message: 'Nominal cicilan harus lebih dari 0.' });
    }
    const partner = await ProfitSharePartner.findByPk(partnerId);
    if (!partner || !assertOwnedByBusiness(partner, req)) {
      return res.status(404).json({ ok: false, message: 'Pihak tidak ditemukan.' });
    }
    const row = await ProfitShareInstallment.create({
      partnerId,
      period,
      paidDate,
      amount,
      note: toText(req.body.note) || null,
    });
    return res.status(201).json({ ok: true, data: serializeProfitShareInstallment(row) });
  } catch (error) {
    return res.status(400).json({ ok: false, message: error.message || 'Gagal menyimpan cicilan.' });
  }
});

app.delete('/api/profit-share/installments/:id', ...authRequired, ownerOnly, async (req, res) => {
  try {
    const row = await ProfitShareInstallment.findByPk(toInt(req.params.id, null));
    if (!row) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    const partner = await ProfitSharePartner.findByPk(row.partnerId);
    if (!partner || !assertOwnedByBusiness(partner, req)) {
      return res.status(404).json({ ok: false, message: 'Data tidak ditemukan.' });
    }
    await row.destroy();
    return res.json({ ok: true, message: 'Cicilan dihapus.' });
  } catch (error) {
    return res.status(500).json({ ok: false, message: 'Gagal menghapus cicilan.' });
  }
});

app.get('/api/dashboard', ...authRequired, async (req, res) => {
  try {
    const from = toText(req.query.from);
    const to = toText(req.query.to);
    const storeId = toInt(req.query.storeId, null);
    const selectedStore = storeId ? await Store.findByPk(storeId) : null;
    if (storeId && (!selectedStore || !assertOwnedByBusiness(selectedStore, req))) {
      return res.status(400).json({ ok: false, message: 'Toko yang dipilih tidak ditemukan.' });
    }
    const platform = selectedStore ? selectedStore.platform : toText(req.query.platform);
    const storeName = selectedStore ? selectedStore.name : toText(req.query.storeName);
    const sku = toText(req.query.sku);
    const data = await buildDashboard(from, to, { platform, storeName, sku, businessId: req.businessId });
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

async function seedExpenseCategories(businessId) {
  const seeds = [
    ...DEFAULT_OPERATIONAL_ORDER.map((name) => ({ name, kind: 'operasional' })),
    ...DEFAULT_MARKETING_ORDER.map((name) => ({ name, kind: 'marketing' })),
  ];
  for (const seed of seeds) {
    await ExpenseCategory.findOrCreate({ where: { ...seed, businessId }, defaults: { ...seed, businessId } });
  }
}

// One-off idempotent backfill (same pattern as migrateOwnerRole) run on every boot: seeds the
// two Business rows and assigns any pre-existing (pre-multi-business) data to "SDS Hatpro",
// since that's the business all data belonged to before this feature existed. Never touches
// owner accounts (businessId stays NULL = unrestricted access to all businesses).
async function migrateBusinessScoping() {
  const [headwear] = await Business.findOrCreate({
    where: { slug: 'sds-headwear' },
    defaults: { name: 'SDS Hatpro', slug: 'sds-headwear', active: true },
  });
  const [fashion] = await Business.findOrCreate({
    where: { slug: 'sds-fashion' },
    defaults: { name: 'SDS Hivvy', slug: 'sds-fashion', active: true },
  });

  const scopedTables = ['stores', 'products', 'sales', 'operational_expenses', 'marketing_expenses', 'expense_categories', 'shifts', 'attendance_settings'];
  for (const tableName of scopedTables) {
    await sequelize.query(
      `UPDATE \`${tableName}\` SET \`business_id\` = :businessId WHERE \`business_id\` IS NULL`,
      { replacements: { businessId: headwear.id } },
    );
  }
  await sequelize.query(
    "UPDATE `users` SET `business_id` = :businessId WHERE `business_id` IS NULL AND `role` != 'owner'",
    { replacements: { businessId: headwear.id } },
  );

  for (const tableName of scopedTables) {
    const rows = await sequelize.query(
      `SELECT COUNT(*) AS total FROM \`${tableName}\` WHERE \`business_id\` IS NULL`,
      { type: sequelize.QueryTypes.SELECT },
    );
    const remaining = toInt(rows[0]?.total, 0);
    if (remaining > 0) {
      console.warn(`[migrateBusinessScoping] ${tableName}: ${remaining} baris masih business_id NULL setelah backfill.`);
    }
  }

  return { headwearId: headwear.id, fashionId: fashion.id };
}

async function migrateSalaryBusinessScoping() {
  const queryInterface = sequelize.getQueryInterface();
  const columns = await queryInterface.describeTable('salary_payments');
  if (!columns.business_id) {
    await queryInterface.addColumn('salary_payments', 'business_id', {
      type: DataTypes.INTEGER,
      allowNull: true,
    });
  }

  // The previous key made one payment row global for a user/week. It must be
  // removed before those rows can be copied into each business.
  const oldIndexes = await queryInterface.showIndex('salary_payments');
  const hasUserForeignKeyIndex = oldIndexes.some((index) => {
    const fields = (index.fields || []).map((field) => field.attribute || field.name);
    return !index.unique && fields.length === 1 && fields[0] === 'user_id';
  });
  if (!hasUserForeignKeyIndex) {
    // MySQL will not drop the old unique index while the user foreign key relies
    // on it. This narrow index keeps that constraint valid during the migration.
    await queryInterface.addIndex('salary_payments', ['user_id'], {
      name: 'salary_payments_user_id_fk',
    });
  }
  for (const index of oldIndexes) {
    const fields = (index.fields || []).map((field) => field.attribute || field.name);
    if (index.unique && fields.length === 2 && fields.includes('user_id') && fields.includes('week_start')) {
      await queryInterface.removeIndex('salary_payments', index.name);
    }
  }

  const businesses = await Business.findAll({ order: [['id', 'ASC']] });
  const defaultBusinessId = businesses[0]?.id || null;
  const legacyPayments = await SalaryPayment.findAll({
    where: { businessId: null },
    include: [{ model: User }],
  });
  for (const legacy of legacyPayments) {
    const targetBusinessIds = legacy.User?.role === 'owner'
      ? businesses.map((business) => business.id)
      : [legacy.User?.businessId || defaultBusinessId].filter(Boolean);
    for (const businessId of targetBusinessIds) {
      await SalaryPayment.findOrCreate({
        where: { userId: legacy.userId, weekStart: legacy.weekStart, businessId },
        defaults: {
          userId: legacy.userId,
          weekStart: legacy.weekStart,
          businessId,
          paid: legacy.paid,
          paidAt: legacy.paidAt,
          bonus: legacy.bonus,
          thr: legacy.thr,
          deduction: legacy.deduction,
        },
      });
    }
    await legacy.destroy();
  }

  if (defaultBusinessId) {
    await SalaryPayment.update(
      { businessId: defaultBusinessId },
      { where: { businessId: null } },
    );
  }
  await queryInterface.changeColumn('salary_payments', 'business_id', {
    type: DataTypes.INTEGER,
    allowNull: false,
  });

  const indexes = await queryInterface.showIndex('salary_payments');
  const hasScopedIndex = indexes.some((index) => {
    const fields = (index.fields || []).map((field) => field.attribute || field.name);
    return index.unique
      && fields.length === 3
      && fields.includes('user_id')
      && fields.includes('week_start')
      && fields.includes('business_id');
  });
  if (!hasScopedIndex) {
    await queryInterface.addIndex('salary_payments', ['user_id', 'week_start', 'business_id'], {
      name: 'salary_payments_user_week_business_unique',
      unique: true,
    });
  }

  // Accounts explicitly named for one business must only appear in that
  // business's salary report. Earlier versions made every owner global and
  // copied its salary profile to every business.
  const businessByOwnerPrefix = new Map([
    ['hatpro.', businesses.find((business) => business.slug === 'sds-headwear')],
    ['hivvy.', businesses.find((business) => business.slug === 'sds-fashion')],
  ]);
  const owners = await User.findAll({ where: { role: 'owner' } });
  for (const owner of owners) {
    const username = String(owner.username || '').toLowerCase();
    const matchedEntry = [...businessByOwnerPrefix.entries()]
      .find(([prefix, business]) => business && username.startsWith(prefix));
    if (!matchedEntry) continue;

    const targetBusiness = matchedEntry[1];
    const settings = await OwnerSalarySetting.findAll({
      where: { userId: owner.id },
      order: [['updatedAt', 'DESC']],
    });
    const source = settings.find((setting) => (
      setting.dailyWage || setting.mealAllowance || setting.overtimeRatePerHour
    )) || settings[0] || owner;
    let target = settings.find((setting) => setting.businessId === targetBusiness.id);
    if (!target) {
      target = await OwnerSalarySetting.create({
        userId: owner.id,
        businessId: targetBusiness.id,
        dailyWage: source.dailyWage || 0,
        mealAllowance: source.mealAllowance || 0,
        overtimeRatePerHour: source.overtimeRatePerHour || 0,
      });
    } else if (
      !(target.dailyWage || target.mealAllowance || target.overtimeRatePerHour)
      && (source.dailyWage || source.mealAllowance || source.overtimeRatePerHour)
    ) {
      target.dailyWage = source.dailyWage || 0;
      target.mealAllowance = source.mealAllowance || 0;
      target.overtimeRatePerHour = source.overtimeRatePerHour || 0;
      await target.save();
    }
    await OwnerSalarySetting.destroy({
      where: { userId: owner.id, businessId: { [Op.ne]: targetBusiness.id } },
    });

    const misplacedPayments = await SalaryPayment.findAll({
      where: { userId: owner.id, businessId: { [Op.ne]: targetBusiness.id } },
    });
    for (const misplaced of misplacedPayments) {
      const targetPayment = await SalaryPayment.findOne({
        where: { userId: owner.id, weekStart: misplaced.weekStart, businessId: targetBusiness.id },
      });
      if (!targetPayment) {
        misplaced.businessId = targetBusiness.id;
        await misplaced.save();
      } else {
        const targetScore = (targetPayment.bonus || 0) + (targetPayment.thr || 0)
          + (targetPayment.deduction || 0) + (targetPayment.paid ? 1 : 0);
        const misplacedScore = (misplaced.bonus || 0) + (misplaced.thr || 0)
          + (misplaced.deduction || 0) + (misplaced.paid ? 1 : 0);
        if (misplacedScore > targetScore) {
          targetPayment.bonus = misplaced.bonus;
          targetPayment.thr = misplaced.thr;
          targetPayment.deduction = misplaced.deduction;
          targetPayment.paid = misplaced.paid;
          targetPayment.paidAt = misplaced.paidAt;
          await targetPayment.save();
        }
        await misplaced.destroy();
      }
    }
  }
}

async function bootstrap() {
  await ensureDatabase();
  await sequelize.authenticate();
  // alter: true menambah unique index duplikat tiap boot (slug_2, slug_3, ...) sampai
  // menabrak limit 64 key MySQL dan bikin bootstrap gagal permanen.
  // Default sekarang sync() biasa; aktifkan alter hanya saat deploy yang mengubah skema:
  //   SEQUELIZE_ALTER=1 pm2 restart topi --update-env
  await sequelize.sync(process.env.SEQUELIZE_ALTER === '1' ? { alter: true } : undefined);
  await migrateOwnerRole();
  await createInitialAdmin();
  const { headwearId, fashionId } = await migrateBusinessScoping();
  await migrateSalaryBusinessScoping();
  await getAttendanceSettings(headwearId);
  await getAttendanceSettings(fashionId);
  await seedExpenseCategories(headwearId);
  await seedExpenseCategories(fashionId);
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
