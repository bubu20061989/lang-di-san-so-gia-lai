/**
 * Làng di sản số Gia Lai – máy chủ dữ liệu thực nghiệm (Google Apps Script, miễn phí)
 * 1) Mở Google Sheets mới → Tiện ích mở rộng → Apps Script → dán toàn bộ tệp này → Lưu.
 * 2) Chạy hàm setup() một lần (cấp quyền) → tạo các trang tính Events, FLCAS, FLCAS_CauHoi.
 * 3) Cài đặt dự án → Thuộc tính tập lệnh (Script properties):
 *      TEACHER_KEY   = mật khẩu giáo viên để xem Dashboard (bắt buộc)
 *      Để Mr. David trả lời bằng AI trên trang GitHub, thêm MỘT trong hai khoá (tuỳ chọn):
 *      DEEPSEEK_API_KEY = khoá API DeepSeek (trả phí theo lượt, rất rẻ)   · DEEPSEEK_MODEL (mặc định deepseek-flash)
 *      GEMINI_API_KEY   = khoá API Gemini (có bậc miễn phí)              · GEMINI_MODEL  (mặc định gemini-2.5-flash)
 *      Có cả hai khoá thì dùng DeepSeek trước, lỗi mới chuyển sang Gemini (hoặc đặt AI_PROVIDER = gemini để ưu tiên Gemini).
 * 4) Triển khai → Tùy chọn triển khai mới → Ứng dụng web → Thực thi với: Tôi; Người có quyền truy cập: Bất kỳ ai → Triển khai.
 * 5) Sao chép URL ứng dụng web (…/exec) và dán vào ô "Địa chỉ Apps Script" trong Dashboard của trang.
 */
const EV_HEAD = ['time','student','station','action','detail','value','device','session','received'];
function sheet_(name, head){ const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(name); if(!sh){ sh=ss.insertSheet(name); if(head) sh.appendRow(head); sh.setFrozenRows(1); } return sh; }
function setup(){
  sheet_('Events', EV_HEAD);
  sheet_('FLCAS', ['time','student','phase','total','mean','answers','device','session']);
  const q=sheet_('FLCAS_CauHoi', ['Câu hỏi (dán Phụ lục 2, mỗi dòng một câu)','Đảo điểm? (R)']);
  if(q.getLastRow()<2) q.getRange(2,1,1,2).setValues([['(Để trống: trang sẽ dùng bộ câu hỏi rút gọn có sẵn)','']]);
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function prop_(k){ return PropertiesService.getScriptProperties().getProperty(k)||''; }
function doGet(e){
  const a=(e.parameter.action||'config');
  if(a==='config'){
    const rows=sheet_('FLCAS_CauHoi').getDataRange().getValues().slice(1).filter(r=>String(r[0]).trim() && !/^\(/.test(String(r[0])));
    return json_({ok:true, ai:!!(prop_('GEMINI_API_KEY')||prop_('DEEPSEEK_API_KEY')), flcas: rows.length? rows.map(r=>({t:String(r[0]).trim(), r:/r/i.test(String(r[1]))})) : null});
  }
  if(a==='events'){
    if(!prop_('TEACHER_KEY') || e.parameter.k!==prop_('TEACHER_KEY')) return json_({ok:false, error:'Sai mật khẩu giáo viên'});
    const v=sheet_('Events', EV_HEAD).getDataRange().getValues(); const h=v.shift();
    return json_({ok:true, rows:v.map(r=>{ const o={}; h.forEach((k,i)=>o[k]= r[i] instanceof Date ? r[i].toISOString() : String(r[i])); return o; })});
  }
  return json_({ok:false, error:'unknown action'});
}
function doPost(e){
  let b={}; try{ b=JSON.parse(e.postData.contents); }catch(err){ return json_({ok:false, error:'bad json'}); }
  if(b.action==='events') return json_(saveEvents_(b.rows||[]));
  if(b.action==='chat') return json_(chat_(String(b.prompt||'')));
  return json_({ok:false, error:'unknown action'});
}
function saveEvents_(rows){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const now=new Date(), ev=sheet_('Events', EV_HEAD), fl=sheet_('FLCAS');
    const clean=rows.slice(0,500).map(r=>EV_HEAD.map(k=> k==='received'? now : String(r[k]==null?'':r[k]).slice(0,500)));
    if(clean.length) ev.getRange(ev.getLastRow()+1,1,clean.length,EV_HEAD.length).setValues(clean);
    rows.filter(r=>r.action==='flcas_pre'||r.action==='flcas_post').forEach(r=>{ const ans=String(r.detail||''); const nums=ans.split(',').map(Number).filter(n=>!isNaN(n));
      fl.appendRow([r.time, r.student, r.action==='flcas_pre'?'pre':'post', nums.reduce((a,b)=>a+b,0), r.value, ans, r.device||'', r.session||'']); });
    return {ok:true, n:clean.length};
  } finally { lock.releaseLock(); }
}
function chat_(prompt){
  const ds=prop_('DEEPSEEK_API_KEY'), gm=prop_('GEMINI_API_KEY'); if(!ds && !gm) return {ok:false, error:'no key'};
  const cache=CacheService.getScriptCache(), slot='rate_'+Math.floor(Date.now()/60000), used=+(cache.get(slot)||0);
  if(used>=40) return {ok:false, error:'rate'}; cache.put(slot,String(used+1),120);
  prompt=prompt.slice(0,12000);
  const order=(prop_('AI_PROVIDER').toLowerCase()==='gemini') ? [gm&&'gemini', ds&&'deepseek'] : [ds&&'deepseek', gm&&'gemini'];
  let last='';
  for(const p of order.filter(Boolean)){ try{ const t= p==='deepseek'? deepseek_(ds,prompt) : gemini_(gm,prompt); if(t) return {ok:true, text:t, provider:p}; }catch(err){ last=String(err); } }
  return {ok:false, error:last||'ai'};
}
function deepseek_(key,prompt){
  const res=UrlFetchApp.fetch('https://api.deepseek.com/chat/completions',{method:'post',contentType:'application/json',muteHttpExceptions:true,headers:{Authorization:'Bearer '+key},
    payload:JSON.stringify({model:prop_('DEEPSEEK_MODEL')||'deepseek-flash', messages:[{role:'user',content:prompt}], response_format:{type:'json_object'}, thinking:{type:'disabled'}, temperature:.7, max_tokens:500})});
  if(res.getResponseCode()!==200) throw new Error('deepseek '+res.getResponseCode()+' '+res.getContentText().slice(0,120));
  const j=JSON.parse(res.getContentText()); return String((((j.choices||[])[0]||{}).message||{}).content||'').trim();
}
function gemini_(key,prompt){
  const model=prop_('GEMINI_MODEL')||'gemini-2.5-flash';
  const res=UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+model+':generateContent?key='+key,{method:'post',contentType:'application/json',muteHttpExceptions:true,
    payload:JSON.stringify({contents:[{role:'user',parts:[{text:prompt}]}], generationConfig:{temperature:.7,maxOutputTokens:500,responseMimeType:'application/json'}})});
  if(res.getResponseCode()!==200) throw new Error('gemini '+res.getResponseCode());
  const j=JSON.parse(res.getContentText()); const t=(((j.candidates||[])[0]||{}).content||{}).parts; return t?t.map(x=>x.text||'').join(''):'';
}
