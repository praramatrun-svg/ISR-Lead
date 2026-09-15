const APP = Object.freeze({
  SHEETS: { COMPANIES: 'Companies', USERS: 'Users', LEADS: 'Leads', ACTIVITIES: 'Activities', SETTINGS: 'Settings' },
  SESSION_TTL: 21600,
  ROLES: ['EMPLOYEE', 'TELESALES', 'ADMIN', 'SUPER_ADMIN']
});

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.refParam = String((e && e.parameter && e.parameter.ref) || '').trim();
  t.companyParam = String((e && e.parameter && e.parameter.company) || '').trim();
  
  let rawUrl = '';
  try {
    rawUrl = ScriptApp.getService().getUrl();
  } catch (err) {
    rawUrl = '';
  }
  let execUrl = rawUrl.replace(/\/u\/\d+\//, '/');
  if (execUrl.indexOf('/exec') === -1 && execUrl.indexOf('/dev') === -1 && execUrl !== '') {
    execUrl = execUrl.replace(/\/edit.*$/, '/exec');
  }
  t.execUrl = execUrl;

  return t.evaluate()
    .setTitle('Autocorp Holding - Lead Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  const schemas = {};
  schemas[APP.SHEETS.COMPANIES] = ['company_id','company_name','company_code','logo_url','primary_color','secondary_color','background_url','welcome_message','contact_phone','privacy_notice_url','consent_version','status','created_at','updated_at'];
  schemas[APP.SHEETS.USERS] = ['user_id','company_id','username','employee_code','display_name','email','phone','branch','role','plain_password','account_status','must_change_password','temporary_password_expires_at','failed_login_count','locked_until','last_login_at','password_changed_at','session_version','created_by','created_at','updated_at'];
  schemas[APP.SHEETS.LEADS] = ['lead_id','lead_reference','company_id','referrer_user_id','branch','assigned_to','customer_title','customer_first_name','customer_last_name','customer_phone','customer_phone_normalized','customer_email','line_id','province','preferred_contact_channel','preferred_contact_time','insurance_type','car_brand','car_model','car_year','has_current_insurance','insurance_expiry_period','lead_status','verification_status','verified_at','next_follow_up_at','last_contact_result','last_note','policy_reference','premium_amount','commission_amount','commission_status','risk_level','created_at','updated_at'];
  schemas[APP.SHEETS.ACTIVITIES] = ['activity_id','company_id','lead_id','user_id','activity_type','old_value','new_value','note','created_at'];
  schemas[APP.SHEETS.SETTINGS] = ['setting_id','company_id','category','setting_key','setting_value','display_label','parent_value','sort_order','is_active'];
  
  Object.keys(schemas).forEach(name => ensureSheet_(ss, name, schemas[name]));
  return 'Setup complete';
}

function apiLogin(username, password) {
  username = normalizeUsername_(username);
  const user = findUserByUsername_(username);
  const now = new Date();

  if (!user) throw new Error(`ไม่พบชื่อผู้ใช้ "${username}" ในระบบ`);

  const status = String(user.account_status || '').trim().toUpperCase();
  if (status !== 'ACTIVE') {
    throw new Error(status === 'PENDING_APPROVAL' ? 'บัญชีกำลังรอ Admin อนุมัติ' : `บัญชีนี้ไม่พร้อมใช้งาน (สถานะ: ${status})`);
  }

  if (user.locked_until && new Date(user.locked_until) > now) {
    throw new Error('บัญชีถูกล็อกชั่วคราว กรุณารอ 15 นาที หรือล้างค่าในช่อง locked_until');
  }

  const savedPass = String(user.plain_password != null ? user.plain_password : '').trim();
  const inputPass = String(password || '').trim();

  if (!savedPass || savedPass !== inputPass) {
    const count = Number(user.failed_login_count || 0) + 1;
    const max = Number(getSetting_('GLOBAL', 'MAX_LOGIN_ATTEMPTS', '5'));
    const patch = { failed_login_count: count };
    if (count >= max) {
      patch.locked_until = new Date(Date.now() + Number(getSetting_('GLOBAL', 'ACCOUNT_LOCK_MINUTES', '15')) * 60000);
    }
    updateById_(APP.SHEETS.USERS, 'user_id', user.user_id, patch);
    logActivity_(user.company_id, '', user.user_id, 'LOGIN_FAILED', '', '', 'Invalid password');
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }

  updateById_(APP.SHEETS.USERS, 'user_id', user.user_id, {
    failed_login_count: 0,
    locked_until: '',
    last_login_at: now
  });

  const token = createSession_(user);
  logActivity_(user.company_id, '', user.user_id, 'LOGIN_SUCCESS', '', '', '');
  return { token, mustChangePassword: truthy_(user.must_change_password), user: publicUser_(user) };
}

function apiLogout(token) {
  if (token) CacheService.getScriptCache().remove('SESSION_' + hashText_(token));
  return true;
}

function apiChangePassword(token, currentPassword, newPassword) {
  const ctx = requireSession_(token);
  const user = findById_(APP.SHEETS.USERS, 'user_id', ctx.user_id);
  if (String(user.plain_password || '').trim() !== String(currentPassword || '').trim()) {
    throw new Error('รหัสผ่านปัจจุบันไม่ถูกต้อง');
  }
  validatePassword_(newPassword);
  updateById_(APP.SHEETS.USERS, 'user_id', user.user_id, {
    plain_password: String(newPassword).trim(),
    must_change_password: false,
    temporary_password_expires_at: '',
    password_changed_at: new Date(),
    session_version: Number(user.session_version || 1) + 1
  });
  CacheService.getScriptCache().remove('SESSION_' + hashText_(token));
  logActivity_(user.company_id, '', user.user_id, 'PASSWORD_CHANGED', '', '', '');
  return true;
}

function apiResetPasswordWithVerification(username, phone, newPassword) {
  username = normalizeUsername_(username);
  validatePassword_(newPassword);
  
  const user = findUserByUsername_(username);
  if (!user) throw new Error('ไม่พบชื่อผู้ใช้นี้ในระบบ');

  const userPhone = String(user.phone || '').replace(/\D/g, '');
  const inputPhone = String(phone || '').replace(/\D/g, '');

  if (!userPhone || userPhone !== inputPhone) {
    throw new Error('เบอร์โทรศัพท์ไม่ตรงกับข้อมูลที่ลงทะเบียนไว้');
  }

  updateById_(APP.SHEETS.USERS, 'user_id', user.user_id, {
    plain_password: String(newPassword).trim(),
    failed_login_count: 0,
    locked_until: '',
    updated_at: new Date()
  });

  logActivity_(user.company_id, '', user.user_id, 'PASSWORD_RESET_SELF', '', '', user.username);
  return { success: true };
}

function apiMe(token) { 
  return publicUser_(findById_(APP.SHEETS.USERS, 'user_id', requireSession_(token).user_id)); 
}

function apiGetBootstrap(token, companyHint, ref) {
  let user = null;
  if (token) { try { user = apiMe(token); } catch (err) {} }
  
  const refUser = ref ? findReferralUser_(ref) : null;
  const companyId = user ? user.company_id : (refUser ? refUser.company_id : (companyHint || ''));
  
  return { 
    user, 
    company: companyId ? publicCompany_(getCompany_(companyId)) : null, 
    companies: apiListPublicCompanies(), 
    options: getPublicOptions_(companyId || 'GLOBAL'),
    referralAgent: refUser ? {
      display_name: refUser.display_name,
      employee_code: refUser.employee_code,
      branch: refUser.branch || 'ไม่ระบุ',
      company_id: refUser.company_id
    } : null
  };
}

function apiListPublicCompanies() {
  return readObjectsCached_(APP.SHEETS.COMPANIES, 600)
    .filter(c => String(c.status || '').toUpperCase() === 'ACTIVE')
    .map(c => ({ company_id: c.company_id, company_name: c.company_name, primary_color: c.primary_color }));
}

function apiRegisterEmployee(data) {
  data = data || {};
  const company = getCompany_(String(data.company_id || ''));
  if (!company || String(company.status || '').toUpperCase() !== 'ACTIVE') throw new Error('กรุณาเลือกบริษัท');
  if (findUserByUsername_(data.username)) throw new Error('ชื่อผู้ใช้นี้ถูกใช้งานแล้ว');
  if (!clean_(data.display_name, 100)) throw new Error('กรุณากรอกชื่อ–นามสกุล');
  if (!clean_(data.branch, 80)) throw new Error('กรุณาระบุสาขาประจำการ');
  
  const empCode = clean_(data.employee_code, 40) || ('EMP' + Math.floor(1000 + Math.random() * 9000));
  validatePassword_(data.password);
  const targetRole = ['EMPLOYEE', 'TELESALES'].includes(data.role) ? data.role : 'EMPLOYEE';

  const user = createUserRecord_({
    company_id: company.company_id,
    username: data.username,
    employee_code: empCode,
    display_name: data.display_name,
    email: data.email || '',
    phone: data.phone || '',
    branch: clean_(data.branch, 80),
    role: targetRole,
    password: data.password,
    must_change_password: false,
    account_status: 'PENDING_APPROVAL',
    created_by: 'SELF_REGISTRATION'
  });

  logActivity_(company.company_id, '', user.user_id, 'USER_REGISTERED', '', 'PENDING_APPROVAL', user.username);
  return { submitted: true, employee_code: empCode };
}

function apiApproveUserWithRole(token, userId, newRole) {
  const ctx = requireRole_(token, ['ADMIN', 'SUPER_ADMIN']);
  const user = findById_(APP.SHEETS.USERS, 'user_id', userId);
  if (!user || (ctx.role !== 'SUPER_ADMIN' && user.company_id !== ctx.company_id)) {
    throw new Error('ไม่พบบัญชีหรือไม่มีสิทธิ์');
  }
  const roleToSet = ['EMPLOYEE', 'TELESALES', 'ADMIN'].includes(newRole) ? newRole : user.role;
  
  updateById_(APP.SHEETS.USERS, 'user_id', userId, {
    account_status: 'ACTIVE',
    role: roleToSet,
    updated_at: new Date(),
    session_version: Number(user.session_version || 1) + 1
  });

  logActivity_(user.company_id, '', ctx.user_id, 'USER_APPROVED', user.account_status, 'ACTIVE', `${user.username} as ${roleToSet}`);
  return true;
}

function apiGetAdminDashboardBundle(token) {
  const ctx = requireRole_(token, ['ADMIN', 'SUPER_ADMIN']);
  const sh = sheet_(APP.SHEETS.LEADS);
  const lastRow = sh.getLastRow();
  let leads = [];
  if (lastRow >= 2) {
    const h = headers_(sh);
    const readCount = Math.min(lastRow - 1, 500);
    const startRow = lastRow - readCount + 1;
    const rawVals = sh.getRange(startRow, 1, readCount, h.length).getValues();

    leads = rawVals
      .filter(r => r.some(x => x !== ''))
      .map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])))
      .filter(r => ctx.role === 'SUPER_ADMIN' || r.company_id === ctx.company_id)
      .reverse()
      .map(publicLead_);
  }

  const pendingUsers = readObjects_(APP.SHEETS.USERS)
    .filter(u => (ctx.role === 'SUPER_ADMIN' || u.company_id === ctx.company_id) && u.account_status === 'PENDING_APPROVAL')
    .map(publicUser_);

  return { leads, pendingUsers };
}

function apiCreateLead(payload) {
  payload = payload || {};
  const refUser = payload.ref ? findReferralUser_(payload.ref) : null;
  const companyId = refUser ? refUser.company_id : String(payload.company_id || '');
  const company = getCompany_(companyId);
  if (!company || String(company.status || '').toUpperCase() !== 'ACTIVE') throw new Error('ไม่พบบริษัทหรือบริษัทไม่เปิดใช้งาน');
  validateLead_(payload);
  
  const phone = normalizeThaiPhone_(payload.customer_phone);
  const duplicate = findDuplicateLead_(companyId, phone, payload.insurance_type);
  const now = new Date();
  const leadId = Utilities.getUuid();
  const reference = (company.company_code || 'LD') + '-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + leadId.slice(0,6).toUpperCase();
  
  const branchTag = refUser ? (refUser.branch || '') : clean_(payload.branch, 80);

  const row = {
    lead_id: leadId,
    lead_reference: reference,
    company_id: companyId,
    referrer_user_id: refUser ? refUser.user_id : '',
    branch: branchTag,
    assigned_to: '',
    customer_title: clean_(payload.customer_title, 20),
    customer_first_name: clean_(payload.customer_first_name, 80),
    customer_last_name: clean_(payload.customer_last_name, 80),
    customer_phone: "'" + maskPhone_(phone),
    customer_phone_normalized: "'" + phone,
    customer_email: clean_(payload.customer_email, 120),
    line_id: clean_(payload.line_id, 80),
    province: clean_(payload.province, 80),
    preferred_contact_channel: clean_(payload.preferred_contact_channel, 30),
    preferred_contact_time: clean_(payload.preferred_contact_time, 80),
    insurance_type: clean_(payload.insurance_type, 40),
    car_brand: clean_(payload.car_brand, 60),
    car_model: clean_(payload.car_model, 60),
    car_year: clean_(payload.car_year, 10),
    has_current_insurance: clean_(payload.has_current_insurance, 10),
    insurance_expiry_period: clean_(payload.insurance_expiry_period, 40),
    lead_status: duplicate ? 'DUPLICATE' : 'PENDING_VERIFICATION',
    verification_status: 'PENDING',
    verified_at: '',
    next_follow_up_at: '',
    last_contact_result: '',
    last_note: '',
    policy_reference: '',
    premium_amount: '',
    commission_amount: '',
    commission_status: '',
    risk_level: duplicate ? 'MEDIUM' : 'LOW',
    created_at: now,
    updated_at: now
  };
  appendObject_(APP.SHEETS.LEADS, row);
  appendConsentActivities_(row, payload);
  logActivity_(companyId, leadId, refUser ? refUser.user_id : '', 'LEAD_CREATED', '', row.lead_status, duplicate ? 'Possible duplicate' : '');
  return { leadId, reference, verificationRequired: true, duplicate: !!duplicate };
}

function apiRequestOtp(leadId) {
  const lead = findById_(APP.SHEETS.LEADS, 'lead_id', leadId);
  if (!lead) throw new Error('ไม่พบ Lead');
  if (lead.verification_status === 'VERIFIED') throw new Error('Lead นี้ยืนยันแล้ว');
  
  const cache = CacheService.getScriptCache();
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const minutes = Number(getSetting_('GLOBAL','OTP_EXPIRY_MINUTES','5'));
  cache.put('OTP_' + leadId, JSON.stringify({ code: otp, attempts: 0 }), minutes * 60);
  
  sendOtp_(lead.customer_phone_normalized, otp);
  logActivity_(lead.company_id, leadId, '', 'OTP_SENT', '', '', '');
  
  const mode = PropertiesService.getScriptProperties().getProperty('OTP_MODE') || 'TEST';
  return { sent: true, testCode: (mode === 'TEST') ? otp : '' };
}

function apiVerifyOtp(leadId, otp) {
  const lead = findById_(APP.SHEETS.LEADS, 'lead_id', leadId);
  if (!lead) throw new Error('ไม่พบ Lead');
  const cache = CacheService.getScriptCache();
  const key = 'OTP_' + leadId;
  const raw = cache.get(key);
  if (!raw) throw new Error('OTP หมดอายุ กรุณาขอรหัสใหม่');
  const item = JSON.parse(raw);
  item.attempts = Number(item.attempts || 0) + 1;
  if (item.attempts > 5) { cache.remove(key); throw new Error('กรอกรหัสผิดเกินกำหนด'); }
  if (String(otp || '').trim() !== String(item.code).trim()) { cache.put(key, JSON.stringify(item), 300); throw new Error('OTP ไม่ถูกต้อง'); }
  cache.remove(key);
  updateById_(APP.SHEETS.LEADS, 'lead_id', leadId, { verification_status: 'VERIFIED', verified_at: new Date(), lead_status: lead.lead_status === 'DUPLICATE' ? 'DUPLICATE' : 'VERIFIED', updated_at: new Date() });
  logActivity_(lead.company_id, leadId, '', 'OTP_VERIFIED', 'PENDING', 'VERIFIED', '');
  return { verified: true, reference: lead.lead_reference };
}

function apiListLeads(token, filters) {
  const ctx = requireSession_(token);
  filters = filters || {};
  const limit = Math.min(Number(filters.limit || 200), 500);

  const sh = sheet_(APP.SHEETS.LEADS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const h = headers_(sh);
  const readCount = Math.min(lastRow - 1, 500);
  const startRow = lastRow - readCount + 1;
  const rawVals = sh.getRange(startRow, 1, readCount, h.length).getValues();

  let rows = rawVals
    .filter(r => r.some(x => x !== ''))
    .map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));

  if (ctx.role !== 'SUPER_ADMIN' && ctx.role !== 'TELESALES') {
    rows = rows.filter(r => r.company_id === ctx.company_id);
  }

  if (ctx.role === 'EMPLOYEE') {
    rows = rows.filter(r => r.referrer_user_id === ctx.user_id);
  }

  if (filters.status && filters.status !== 'ALL') {
    rows = rows.filter(r => r.lead_status === filters.status);
  }

  return rows.reverse().slice(0, limit).map(publicLead_);
}

function apiUpdateLead(token, leadId, patch) {
  const ctx = requireSession_(token);
  const lead = findById_(APP.SHEETS.LEADS, 'lead_id', leadId);
  assertLeadAccess_(ctx, lead, true);
  const allowed = ['lead_status','assigned_to','next_follow_up_at','last_contact_result','last_note','policy_reference','premium_amount','commission_amount','commission_status'];
  const cleanPatch = { updated_at: new Date() };
  allowed.forEach(k => { if (Object.prototype.hasOwnProperty.call(patch || {}, k)) cleanPatch[k] = clean_(patch[k], k === 'last_note' ? 500 : 100); });
  if (cleanPatch.assigned_to && !['ADMIN','SUPER_ADMIN'].includes(ctx.role)) delete cleanPatch.assigned_to;
  if ((cleanPatch.commission_amount || cleanPatch.commission_status) && !['ADMIN','SUPER_ADMIN'].includes(ctx.role)) { delete cleanPatch.commission_amount; delete cleanPatch.commission_status; }
  updateById_(APP.SHEETS.LEADS, 'lead_id', leadId, cleanPatch);
  logActivity_(lead.company_id, leadId, ctx.user_id, 'LEAD_UPDATED', JSON.stringify({status:lead.lead_status}), JSON.stringify(cleanPatch), cleanPatch.last_note || '');
  return true;
}

function apiListUsers(token) {
  const ctx = requireRole_(token, ['ADMIN','SUPER_ADMIN']);
  return readObjects_(APP.SHEETS.USERS).filter(u => ctx.role === 'SUPER_ADMIN' || u.company_id === ctx.company_id).map(publicUser_);
}

function apiCreateUser(token, data) {
  const ctx = requireRole_(token, ['ADMIN','SUPER_ADMIN']);
  const companyId = ctx.role === 'SUPER_ADMIN' ? String(data.company_id || ctx.company_id) : ctx.company_id;
  if (findUserByUsername_(data.username)) throw new Error('Username นี้ถูกใช้แล้ว');
  if (!APP.ROLES.includes(data.role) || (data.role === 'SUPER_ADMIN' && ctx.role !== 'SUPER_ADMIN')) throw new Error('Role ไม่ถูกต้อง');
  const tempPassword = makeTempPassword_();
  const user = createUserRecord_({ company_id: companyId, username: data.username, employee_code: data.employee_code, display_name: data.display_name, email: data.email, phone: data.phone, branch: data.branch, role: data.role, password: tempPassword, must_change_password: true, created_by: ctx.user_id });
  logActivity_(companyId, '', ctx.user_id, 'USER_CREATED', '', user.user_id, user.username);
  return { user: publicUser_(user), temporaryPassword: tempPassword };
}

function sendOtp_(phone, otp) {
  const props = PropertiesService.getScriptProperties();
  const mode = props.getProperty('OTP_MODE') || 'TEST';
  if (mode === 'TEST') return;
  if (mode !== 'HTTP_API') return;
  const url = props.getProperty('OTP_API_URL');
  const token = props.getProperty('OTP_API_TOKEN');
  if (!url || !token) return;
  UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+token}, payload:JSON.stringify({to:phone, code:otp}), muteHttpExceptions:true });
}

function createUserRecord_(d) {
  validatePassword_(d.password);
  const now = new Date();
  const row = { 
    user_id: Utilities.getUuid(), 
    company_id: d.company_id, 
    username: normalizeUsername_(d.username), 
    employee_code: clean_(d.employee_code, 40), 
    display_name: clean_(d.display_name, 100), 
    email: clean_(d.email, 120), 
    phone: clean_(d.phone, 30), 
    branch: clean_(d.branch, 80), 
    role: d.role, 
    plain_password: String(d.password).trim(),
    account_status: d.account_status || 'ACTIVE', 
    must_change_password: !!d.must_change_password, 
    temporary_password_expires_at: d.must_change_password ? new Date(Date.now() + 86400000) : '', 
    failed_login_count: 0, 
    locked_until: '', 
    last_login_at: '', 
    password_changed_at: '', 
    session_version: 1, 
    created_by: d.created_by || 'SYSTEM', 
    created_at: now, 
    updated_at: now 
  };
  appendObject_(APP.SHEETS.USERS, row); 
  return row;
}

function createSession_(user) {
  const token = randomToken_() + randomToken_();
  CacheService.getScriptCache().put('SESSION_' + hashText_(token), JSON.stringify({ user_id: user.user_id, company_id: user.company_id, role: user.role, session_version: Number(user.session_version || 1) }), APP.SESSION_TTL);
  return token;
}

function requireSession_(token) {
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  const raw = CacheService.getScriptCache().get('SESSION_' + hashText_(token));
  if (!raw) throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const ctx = JSON.parse(raw), user = findById_(APP.SHEETS.USERS, 'user_id', ctx.user_id);
  if (!user || String(user.account_status || '').toUpperCase() !== 'ACTIVE' || Number(user.session_version || 1) !== Number(ctx.session_version)) throw new Error('Session ไม่ถูกต้อง');
  return ctx;
}

function requireRole_(token, roles) { const ctx = requireSession_(token); if (!roles.includes(ctx.role)) throw new Error('คุณไม่มีสิทธิ์ทำรายการนี้'); return ctx; }

function assertLeadAccess_(ctx, lead, write) { 
  if (!lead) throw new Error('ไม่พบ Lead'); 
  if (ctx.role === 'SUPER_ADMIN' || ctx.role === 'TELESALES') return;
  if (lead.company_id !== ctx.company_id) throw new Error('ไม่มีสิทธิ์เข้าถึงข้อมูล'); 
  if (ctx.role === 'EMPLOYEE' && lead.referrer_user_id !== ctx.user_id) throw new Error('ไม่มีสิทธิ์เข้าถึงข้อมูล'); 
}

function getCompany_(id) { return readObjectsCached_(APP.SHEETS.COMPANIES, 600).find(r => r.company_id === id); }
function publicCompany_(c) { return c ? { company_id: c.company_id, company_name: c.company_name, logo_url: c.logo_url, primary_color: c.primary_color || '#0b2239', secondary_color: c.secondary_color || '#eff6ff', background_url: c.background_url, welcome_message: c.welcome_message, contact_phone: c.contact_phone, privacy_notice_url: c.privacy_notice_url, consent_version: c.consent_version } : null; }
function publicUser_(u) { return { user_id: u.user_id, company_id: u.company_id, username: u.username, employee_code: u.employee_code, display_name: u.display_name, email: u.email, phone: u.phone, branch: u.branch, role: u.role, account_status: u.account_status, must_change_password: truthy_(u.must_change_password) }; }
function publicLead_(r) { return { lead_id: r.lead_id, lead_reference: r.lead_reference, company_id: r.company_id, referrer_user_id: r.referrer_user_id, branch: r.branch || 'ไม่ระบุ', assigned_to: r.assigned_to, customer_name: [r.customer_title, r.customer_first_name, r.customer_last_name].filter(Boolean).join(' '), customer_phone: r.customer_phone, province: r.province, insurance_type: r.insurance_type, car_brand: r.car_brand, car_model: r.car_model, car_year: r.car_year, lead_status: r.lead_status, verification_status: r.verification_status, next_follow_up_at: r.next_follow_up_at, last_contact_result: r.last_contact_result, last_note: r.last_note, premium_amount: r.premium_amount, commission_amount: r.commission_amount, commission_status: r.commission_status, risk_level: r.risk_level, created_at: r.created_at }; }

function getPublicOptions_(companyId) {
  const all = readObjectsCached_(APP.SHEETS.SETTINGS, 600).filter(r => truthy_(r.is_active));
  const consents = all.filter(r => r.category === 'CONSENT' && (r.company_id === 'GLOBAL' || r.company_id === companyId)).map(r => ({ key: r.setting_key, text: r.setting_value, label: r.display_label }));
  
  const branchMap = {};
  all.filter(r => r.category === 'BRANCH').forEach(r => {
    const cId = r.company_id;
    if (!branchMap[cId]) branchMap[cId] = [];
    branchMap[cId].push({ name: r.setting_value, sort: Number(r.sort_order || 99) });
  });
  Object.keys(branchMap).forEach(k => branchMap[k].sort((a,b) => a.sort - b.sort));
  return { consents, branches: branchMap };
}

function getSetting_(companyId, key, fallback) {
  const row = readObjectsCached_(APP.SHEETS.SETTINGS, 600).find(r => (r.company_id === companyId || r.company_id === 'GLOBAL') && r.setting_key === key && truthy_(r.is_active));
  return row ? row.setting_value : fallback;
}

function findReferralUser_(ref) { return readObjects_(APP.SHEETS.USERS).find(u => u.employee_code === String(ref) && String(u.account_status || '').toUpperCase() === 'ACTIVE'); }

function findUserByUsername_(username) {
  username = normalizeUsername_(username);
  const sh = sheet_(APP.SHEETS.USERS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const h = headers_(sh).map(x => String(x).trim());
  const userIdx = h.indexOf('username');
  if (userIdx === -1) throw new Error('ไม่พบคอลัมน์ username ในชีต Users');

  const data = sh.getRange(2, 1, lastRow - 1, h.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (normalizeUsername_(data[i][userIdx]) === username) {
      return Object.fromEntries(h.map((k, idx) => [k, data[i][idx]]));
    }
  }
  return null;
}

function findDuplicateLead_(companyId, phone, type) { 
  const days = Number(getSetting_(companyId, 'DUPLICATE_LEAD_DAYS', '90')); 
  const since = Date.now() - days * 86400000; 
  return readObjects_(APP.SHEETS.LEADS).find(r => r.company_id === companyId && r.customer_phone_normalized === phone && r.insurance_type === type && new Date(r.created_at).getTime() >= since); 
}

function appendConsentActivities_(lead, p) { 
  const consents = p.consents || {}; 
  Object.keys(consents).forEach(k => logActivity_(lead.company_id, lead.lead_id, lead.referrer_user_id, 'CONSENT_RECORDED', '', String(!!consents[k]), k)); 
}

function validateLead_(p) {
  const fieldLabels = {
    'customer_first_name': 'ชื่อจริง',
    'customer_last_name': 'นามสกุล',
    'customer_phone': 'เบอร์โทรศัพท์',
    'province': 'จังหวัด',
    'preferred_contact_time': 'วันและเวลาที่สะดวกติดต่อ',
    'insurance_type': 'ประเภทประกันที่สนใจ',
    'car_brand': 'ยี่ห้อรถ',
    'car_model': 'รุ่นรถ',
    'car_year': 'ปีรถ'
  };

  for (const [key, label] of Object.entries(fieldLabels)) {
    if (!String(p[key] || '').trim()) throw new Error(`กรุณากรอกข้อมูล: ${label}`);
  }

  if (!p.consents || p.consents.CONTACT_AND_MARKETING !== true) {
    throw new Error('กรุณาให้ความยินยอมเพื่อให้เจ้าหน้าที่ติดต่อกลับ (ข้อ 3)');
  }
  normalizeThaiPhone_(p.customer_phone);
}

function validatePassword_(p) { 
  p = String(p || '').trim(); 
  if (p.length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร'); 
}

function normalizeUsername_(s) { return String(s || '').trim().toLowerCase(); }
function normalizeThaiPhone_(s) { let d = String(s || '').replace(/\D/g, ''); if (d.startsWith('66')) d = '0' + d.slice(2); if (!/^0[689]\d{8}$/.test(d)) throw new Error('รูปแบบเบอร์โทรศัพท์ไม่ถูกต้อง'); return '+66' + d.slice(1); }
function maskPhone_(p) { return p.replace(/^(\+66\d)(\d{5})(\d{3})$/, '$1-XXX-XX$3'); }
function clean_(v, max) { let s = String(v == null ? '' : v).trim(); if (/^[=+\-@]/.test(s)) s = "'" + s; return s.slice(0, max || 500); }
function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function randomToken_() { return Utilities.getUuid().replace(/-/g, ''); }
function hashText_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s))); }
function makeTempPassword_() { return 'Tmp' + Math.floor(100000 + Math.random() * 900000); }

function ensureSheet_(ss, name, headers) { 
  let sh = ss.getSheetByName(name) || ss.insertSheet(name); 
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]); 
  const current = sh.getRange(1, 1, 1, headers.length).getValues()[0]; 
  if (current.join('|') !== headers.join('|')) sh.getRange(1, 1, 1, headers.length).setValues([headers]); 
  sh.setFrozenRows(1); 
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0B3A6E').setFontColor('#FFFFFF'); 
}

function sheet_(name) { const sh = SpreadsheetApp.getActive().getSheetByName(name); if (!sh) throw new Error('Missing sheet: ' + name); return sh; }
function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }

function readObjects_(name) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  const h = headers_(sh), v = sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getValues();
  return v.filter(r => r.some(x => x !== '')).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}

function readObjectsCached_(name, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'CACHE_SHEET_' + name;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const data = readObjects_(name);
  try { cache.put(cacheKey, JSON.stringify(data), ttlSeconds || 300); } catch (e) {}
  return data;
}

function appendObject_(name, obj) {
  const sh = sheet_(name), h = headers_(sh);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sh.appendRow(h.map(k => Object.prototype.hasOwnProperty.call(obj, k) ? obj[k] : ''));
    CacheService.getScriptCache().remove('CACHE_SHEET_' + name);
  } finally {
    lock.releaseLock();
  }
}

function findById_(name, key, id) { return readObjects_(name).find(r => String(r[key]) === String(id)); }

function updateById_(name, key, id, patch) {
  const sh = sheet_(name), h = headers_(sh), idx = h.indexOf(key);
  if (idx < 0) throw new Error('Missing key column');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), h.length).getValues();
    const n = vals.findIndex(r => String(r[idx]) === String(id));
    if (n < 0) throw new Error('Record not found');
    Object.keys(patch).forEach(k => { const c = h.indexOf(k); if (c >= 0) vals[n][c] = patch[k]; });
    sh.getRange(n + 2, 1, 1, h.length).setValues([vals[n]]);
    CacheService.getScriptCache().remove('CACHE_SHEET_' + name);
  } finally {
    lock.releaseLock();
  }
}

function logActivity_(companyId, leadId, userId, type, oldVal, newVal, note) {
  appendObject_(APP.SHEETS.ACTIVITIES, {
    activity_id: Utilities.getUuid(),
    company_id: companyId,
    lead_id: leadId,
    user_id: userId,
    activity_type: type,
    old_value: clean_(oldVal, 500),
    new_value: clean_(newVal, 500),
    note: clean_(note, 500),
    created_at: new Date()
  });
}
