/**
 * Máy chủ Google Apps Script cho web Khảo sát nhu cầu CGKT 2027 – BV Phổi Hải Phòng
 * Dữ liệu lưu trực tiếp trong Google Sheet này (mỗi bảng 1 tab).
 *
 * CÀI ĐẶT (1 lần):
 *  1. Tạo 1 Google Sheet mới → Tiện ích mở rộng → Apps Script → dán toàn bộ file này.
 *  2. Đổi ADMIN_PW bên dưới thành mật khẩu quản trị của anh → Lưu.
 *  3. Chọn hàm "caiDat" → Chạy (cấp quyền khi được hỏi) → các tab được tạo.
 *  4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *       Thực thi với tư cách: Tôi   |   Người có quyền truy cập: Bất kỳ ai
 *     → Sao chép URL dạng https://script.google.com/macros/s/.../exec
 *  5. Dán URL đó vào GSHEET_URL trong file index.html.
 *  Sửa code sau này: Triển khai → Quản lý triển khai → Sửa → Phiên bản mới.
 */
const ADMIN_PW = 'doi-mat-khau-nay';

const TABLES = {
  cau_hinh: ['id','mo','han_chot','thong_bao','phu_luc'],
  khoa:     ['id','ten','thu_tu'],
  don_vi:   ['id','ten','tuyen','ghi_chu','duyet'],
  ky_thuat: ['id','ten','chuyen_nganh','duyet','de_xuat_khoa','created_at'],
  lua_chon: ['id','ky_thuat_id','don_vi_id','doi_tuong','thoi_gian','goi','hoc_phi','nguon_tt','duyet'],
  dang_ky:  ['id','khoa_id','loai','ky_thuat_id','lua_chon_id','don_vi_id','ten_kt','don_vi_nhan','huong','don_vi_khac','chuyen_khoa',
             'hinh_thuc','so_nguoi','doi_tuong','thoi_gian','ghi_chu','nguoi_khai','sdt','created_at','updated_at',
             'nhom_dt','loai_dt','ho_ten','chuc_danh','so_thang','co_so_dt','kinh_phi']
};
const BOOL = {mo:1, duyet:1};
const NUM  = {so_nguoi:1, thu_tu:1};
const VIEW = 'Xem tổng hợp';
const VIEW5 = 'PL5 – Đào tạo';
const TEN_BV = 'Bệnh viện Phổi Hải Phòng';

/* ---------- Sheet helpers ---------- */
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(t){
  let s = ss_().getSheetByName(t);
  if(!s){
    s = ss_().insertSheet(t);
    const h = TABLES[t];
    s.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#E8F0FA');
    s.setFrozenRows(1);
    s.getRange(1,1,s.getMaxRows(),h.length).setNumberFormat('@');
  }
  return s;
}
function ensureCols_(t){ // bổ sung cột mới nếu sheet cũ còn thiếu (chỉ chạy khi ghi / cài đặt)
  const s = sh_(t);
  if(!TABLES[t]) return s;
  const cur = s.getRange(1,1,1,Math.max(s.getLastColumn(),1)).getValues()[0].map(String);
  TABLES[t].filter(k => cur.indexOf(k) < 0).forEach(k => { const c = s.getLastColumn()+1; s.getRange(1,c).setValue(k).setFontWeight('bold').setBackground('#E8F0FA'); s.getRange(1,c,s.getMaxRows(),1).setNumberFormat('@'); });
  return s;
}
function header_(s){ return s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(String); }
let MEMO_ = {};
function readT_(t){ // đọc cả bảng bằng 1 lệnh, nhớ tạm trong 1 lượt chạy
  if(MEMO_[t]) return MEMO_[t];
  const all = sh_(t).getDataRange().getValues();
  if(all.length < 2) return (MEMO_[t] = []);
  const h = all[0].map(String);
  return (MEMO_[t] = all.slice(1).filter(r => String(r[0]) !== '').map(r => {
    const o = {};
    h.forEach((k,i) => {
      if(!k) return;
      let x = r[i];
      if(BOOL[k]) x = (x === true || String(x).toUpperCase() === 'TRUE');
      else if(NUM[k] || (t==='cau_hinh' && k==='id')) x = (x === '' ? null : Number(x));
      else x = (x === '' || x === null ? null : (x instanceof Date ? x.toISOString() : String(x)));
      o[k] = x;
    });
    return o;
  }));
}

/* ---------- Bộ nhớ đệm (CacheService) cho doGet ---------- */
const CK_ = 'all_v1';
function cacheGet_(){
  const c = CacheService.getScriptCache(), n = c.get(CK_+'_n');
  if(!n) return null;
  const keys = []; for(let i=0;i<+n;i++) keys.push(CK_+'_'+i);
  const m = c.getAll(keys); let s = '';
  for(const k of keys){ if(m[k] == null) return null; s += m[k]; }
  return s;
}
function cachePut_(s){
  const c = CacheService.getScriptCache(), o = {}, size = 90000; let n = 0;
  for(let i=0;i<s.length;i+=size) o[CK_+'_'+(n++)] = s.slice(i,i+size);
  o[CK_+'_n'] = String(n);
  try{ c.putAll(o, 21600); }catch(e){}
}
function cacheClear_(){ try{ CacheService.getScriptCache().remove(CK_+'_n'); }catch(e){} MEMO_ = {}; }
function allJson_(){
  let s = cacheGet_();
  if(!s){ const out = {}; Object.keys(TABLES).forEach(t => out[t] = readT_(t)); s = JSON.stringify(out); cachePut_(s); }
  return s;
}

function cell_(v){ if(v === null || v === undefined) return ''; if(typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; return String(v); }
function rowIndex_(s, col, val){ // trả về các số dòng (1-based) khớp
  const h = header_(s), c = h.indexOf(col), last = s.getLastRow();
  if(c < 0 || last < 2) return [];
  const v = s.getRange(2, c+1, last-1, 1).getValues(), out = [];
  v.forEach((r,i) => { if(String(r[0]) === String(val)) out.push(i+2); });
  return out;
}
function rowObj_(s, h, r){ const v = s.getRange(r,1,1,h.length).getValues()[0], o = {}; h.forEach((k,i)=>o[k]=v[i]); return o; }

/* ---------- Thao tác ---------- */
function insert_(t, rows){
  const s = ensureCols_(t), h = header_(s);
  if(!rows.length) return;
  const vals = rows.map(r => h.map(k => cell_(r[k])));
  const start = s.getLastRow() + 1;
  s.getRange(start, 1, vals.length, h.length).setNumberFormat('@').setValues(vals);
}
function patchRow_(s, h, r, patch){
  Object.keys(patch).forEach(k => { const c = h.indexOf(k); if(c >= 0 && k !== 'id') s.getRange(r, c+1).setNumberFormat('@').setValue(cell_(patch[k])); });
}
function update_(t, id, patch){ const s = sh_(t), h = header_(s); rowIndex_(s,'id',id).forEach(r => patchRow_(s,h,r,patch)); }
function updateWhere_(t, col, val, patch){ const s = sh_(t), h = header_(s); rowIndex_(s,col,val).forEach(r => patchRow_(s,h,r,patch)); }
function remove_(t, id){ const s = sh_(t); rowIndex_(s,'id',id).reverse().forEach(r => s.deleteRow(r)); }
function clear_(t){ const s = sh_(t), last = s.getLastRow(); if(last >= 2) s.deleteRows(2, last-1); }

function isOpen_(){
  const c = readT_('cau_hinh')[0];
  if(!c) return true;
  return !!c.mo && (!c.han_chot || new Date() <= new Date(c.han_chot));
}

// phu_luc: trống/null (dữ liệu cũ) = mở cả 4; 'none' = đóng hết; '1,3' = chỉ mở PL1, PL3
function plOn_(n){
  const c = readT_('cau_hinh')[0] || {}, v = c.phu_luc;
  if(v == null || v === '') return true;
  return String(v).split(',').map(x => x.trim()).indexOf(String(n)) >= 0;
}

/* ---------- Kiểm tra quyền ---------- */
function check_(req){
  const admin = req.pw && req.pw === ADMIN_PW;
  if(admin) return true;
  const t = req.table, a = req.action;
  if(!isOpen_()) throw new Error('Đã khóa nhập liệu');
  if(t === 'dang_ky'){
    const tuyen = {}; readT_('don_vi').forEach(d => tuyen[d.id] = d.tuyen);
    const plOf = r => r.loai === 'kc' ? String(r.huong || '') : r.loai === 'dt' ? '5' : r.loai === 'cg' ? '3' : r.loai === 'lp' ? '4' : (tuyen[r.don_vi_id] === 'TW' ? '1' : tuyen[r.don_vi_id] ? '2' : null);
    const need = pl => { if(pl && !plOn_(pl)) throw new Error('Phụ lục ' + pl + ' đang đóng'); };
    if(a === 'insert'){ [].concat(req.rows).forEach(r => need(plOf(r))); return true; }
    if(a === 'update' || a === 'remove'){
      const s = sh_(t), h = header_(s), rows = rowIndex_(s,'id',req.id);
      if(!rows.length) return true;
      const o = rowObj_(s,h,rows[0]);
      if(String(o.khoa_id) !== String(req.khoa_id)) throw new Error('Chỉ khoa/phòng đã đề xuất mới được sửa/xóa');
      need(plOf(o));
      if(a === 'update' && req.patch){
        if(req.patch.khoa_id && String(req.patch.khoa_id) !== String(o.khoa_id)) throw new Error('Không được đổi khoa');
        need(plOf(Object.assign({}, o, req.patch)));
      }
      return true;
    }
  }
  if(['ky_thuat','don_vi','lua_chon'].indexOf(t) >= 0 && a === 'insert' && [].concat(req.rows).every(r => !r.duyet)) return true;
  throw new Error('Không có quyền thực hiện thao tác này');
}

/* ---------- Bảng xem dễ đọc ---------- */
function buildView_(){
  const k = {}, kt = {}, dv = {}, lc = {};
  readT_('khoa').forEach(x => k[x.id] = x);
  readT_('ky_thuat').forEach(x => kt[x.id] = x);
  readT_('don_vi').forEach(x => dv[x.id] = x);
  readT_('lua_chon').forEach(x => lc[x.id] = x);
  const H = ['STT','Khoa/phòng đề xuất','Phụ lục','Chuyên ngành','Kỹ thuật / nội dung','Đơn vị chuyển giao/hỗ trợ','Đơn vị nhận','Hình thức','Số người','Đối tượng','Thời gian','Gói/khóa','Học phí','Ghi chú','Người đề xuất','SĐT','Cập nhật'];
  const rows = readT_('dang_ky').map(d => {
    const kh = (k[d.khoa_id]||{}).ten || '?', me = TEN_BV + ' (' + kh + ')';
    const o = lc[d.lua_chon_id] || {}, v = dv[d.don_vi_id] || {}, t = kt[d.ky_thuat_id] || {};
    let pl, ten, a, b;
    if(d.loai === 'nhan'){ pl = v.tuyen === 'TW' ? 'PL1' : 'PL2'; ten = t.ten; a = v.ten; b = me; }
    else if(d.loai === 'cg'){ pl = 'PL3'; ten = d.ten_kt; a = me; b = d.don_vi_nhan; }
    else if(d.loai === 'dt'){ pl = 'PL5'; ten = [d.loai_dt, d.chuyen_khoa].filter(Boolean).join(' – ') + (d.ho_ten ? ' (' + d.ho_ten + ')' : ''); a = d.co_so_dt; b = me; }
    else if(d.loai === 'kc'){ pl = 'PL' + d.huong; ten = d.ten_kt || 'Không có nhu cầu'; a = ''; b = me; }
    else { pl = 'PL4'; ten = 'Chuyên khoa: ' + d.chuyen_khoa; a = d.huong === 'nhan' ? d.don_vi_khac : me; b = d.huong === 'nhan' ? me : d.don_vi_khac; }
    return [kh, pl, t.chuyen_nganh || '', ten || '', a || '', b || '', d.hinh_thuc || '', d.so_nguoi == null ? '' : d.so_nguoi, d.doi_tuong || '', d.thoi_gian || '',
            [o.goi, o.thoi_gian].filter(Boolean).join(', '), o.hoc_phi || '', d.ghi_chu || '', d.nguoi_khai || '', d.sdt || '', d.updated_at || ''];
  }).sort((x,y) => (x[0]+x[1]).localeCompare(y[0]+y[1], 'vi')).map((r,i) => [i+1].concat(r));
  let s = ss_().getSheetByName(VIEW);
  if(!s){ s = ss_().insertSheet(VIEW, 0); }
  s.clear();
  s.getRange(1,1,1,H.length).setValues([H]).setFontWeight('bold').setBackground('#E8F0FA');
  s.setFrozenRows(1);
  if(rows.length) s.getRange(2,1,rows.length,H.length).setValues(rows);
  s.getRange(1,1,Math.max(rows.length,1)+1,H.length).setWrap(true).setVerticalAlignment('top');
  buildView5_(k);
}

/* ---------- Tab riêng PL5 – Đào tạo văn bằng, chứng chỉ ---------- */
function buildView5_(k){
  if(!k){ k = {}; readT_('khoa').forEach(x => k[x.id] = x); }
  const kTen = d => (k[d.khoa_id]||{}).ten || '?', kOrd = d => { const x = k[d.khoa_id]; return x && x.thu_tu != null ? x.thu_tu : 999; };
  const nguoi = d => { const t = String(d.ho_ten||'').split(';').map(x=>x.trim()), c = String(d.chuc_danh||'').split(';').map(x=>x.trim());
    return t.map((x,i) => x ? x + (c[i] ? ' (' + c[i] + ')' : '') : '').filter(Boolean).join('\n'); };
  const H = ['STT','Nhóm','Khoa/Phòng','Loại đào tạo','Chuyên ngành / nội dung','Người được cử (chức danh)','Số lượng','Cơ sở đào tạo dự kiến','Bắt đầu dự kiến','Thời gian (tháng)','Hình thức','Nguồn kinh phí','Ghi chú','Người đề xuất','SĐT','Cập nhật'];
  const all = readT_('dang_ky'), rows = [];
  ['Văn bằng','Chứng chỉ'].forEach(nhom => {
    all.filter(d => d.loai === 'dt' && (d.nhom_dt || 'Văn bằng') === nhom)
      .sort((a,b) => String(a.loai_dt||'').localeCompare(String(b.loai_dt||''),'vi') || kOrd(a) - kOrd(b))
      .forEach(d => rows.push([nhom, kTen(d), d.loai_dt||'', d.chuyen_khoa||'', nguoi(d), d.so_nguoi == null ? '' : d.so_nguoi, d.co_so_dt||'', d.thoi_gian||'',
        d.so_thang||'', d.hinh_thuc||'', d.kinh_phi||'', d.ghi_chu||'', d.nguoi_khai||'', d.sdt||'', d.updated_at||'']));
  });
  all.filter(d => d.loai === 'kc' && String(d.huong) === '5').sort((a,b) => kOrd(a) - kOrd(b))
    .forEach(d => rows.push(['Không có nhu cầu', kTen(d), 'Không có nhu cầu đào tạo văn bằng, chứng chỉ', '', '', '', '', '', '', '', '', d.ghi_chu||'', d.nguoi_khai||'', d.sdt||'', d.updated_at||'']));
  const out = rows.map((r,i) => [i+1].concat(r));
  let s = ss_().getSheetByName(VIEW5);
  if(!s){ s = ss_().insertSheet(VIEW5, 1); }
  s.clear();
  s.getRange(1,1,1,H.length).setValues([H]).setFontWeight('bold').setBackground('#E8F0FA');
  s.setFrozenRows(1);
  if(out.length) s.getRange(2,1,out.length,H.length).setValues(out);
  s.getRange(1,1,Math.max(out.length,1)+1,H.length).setWrap(true).setVerticalAlignment('top');
}

/* ---------- Web app ---------- */
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e){
  try{
    const data = allJson_();
    if(e && e.parameter && e.parameter.only === 'cau_hinh') return json_({ok:true, data:{cau_hinh: JSON.parse(data).cau_hinh}});
    return ContentService.createTextOutput('{"ok":true,"data":' + data + '}').setMimeType(ContentService.MimeType.JSON);
  }catch(err){ return json_({ok:false, error:String(err.message||err)}); }
}
function doPost(e){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    const req = JSON.parse(e.postData.contents);
    if(req.action === 'login'){ if(req.pw === ADMIN_PW) return json_({ok:true}); throw new Error('Sai mật khẩu'); }
    if(!TABLES[req.table]) throw new Error('Bảng không hợp lệ');
    check_(req);
    const t = req.table;
    ensureCols_(t);
    switch(req.action){
      case 'insert': insert_(t, [].concat(req.rows)); break;
      case 'update': update_(t, req.id, req.patch || {}); break;
      case 'updateWhere': updateWhere_(t, req.col, req.val, req.patch || {}); break;
      case 'remove': remove_(t, req.id); break;
      case 'clear': clear_(t); break;
      default: throw new Error('Thao tác không hợp lệ');
    }
    SpreadsheetApp.flush();
    cacheClear_();
    if(['dang_ky','khoa','ky_thuat','don_vi','lua_chon'].indexOf(t) >= 0) buildView_();
    return json_({ok:true});
  }catch(err){
    return json_({ok:false, error:String(err.message||err)});
  }finally{ try{ lock.releaseLock(); }catch(x){} }
}

/* ---------- Chạy tay ---------- */
function caiDat(){ Object.keys(TABLES).forEach(ensureCols_); cacheClear_(); buildView_(); }
function onEdit(){ cacheClear_(); } // sửa tay trên Sheet → web thấy ngay
function onOpen(){ SpreadsheetApp.getUi().createMenu('Khảo sát CGKT').addItem('Cập nhật tab "Xem tổng hợp" + "PL5 – Đào tạo"','buildView_').addToUi(); }
