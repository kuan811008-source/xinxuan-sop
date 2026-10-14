/**
 * 金山文檔 AirScript：把「微糖直播出貨進度表」推送到心選SOP看板資料庫
 *
 * 使用方式：
 * 1. 在金山文檔開啟表格 → 選單「效率」→「高級開發」→「AirScript 腳本編輯器」
 * 2. 新建腳本，把整份程式碼貼上
 * 3. 按「執行」測試；成功後可在「定時任務」設定每 30 分鐘自動執行
 *
 * 欄位以「標題文字」辨識（標題列自動偵測），欄位順序改變或有隱藏欄都不影響。
 */

const SUPA_URL = 'https://ijnjfpxthoewpwyuzela.supabase.co';
const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlqbmpmcHh0aG9ld3B3eXV6ZWxhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYyMTc4MDksImV4cCI6MjA5MTc5MzgwOX0.kiyAtaVDUFlaV1pWzAMGyfdMuh4GScO8o47CtjnRYy8';
const SHEET_NAME = '';   // 留空＝使用第一個工作表；要指定頁籤就填名稱，例如 '出貨進度'

// 欄位標題關鍵字（用「包含」比對，寫部分文字即可）
const COLS = {
  closed:      '結案',
  session:     '場次',
  vendor:      '廠商',
  cart:        '購物',
  order_time:  '叫貨時間',
  packing_site:'包裝場',
  checklist:   '提供清點',
  arrival:     '到貨',
  counting:    '清點狀況',
  ship_counts: '出貨單數',
  ship_start:  '開始出貨',
  ship_done:   '出貨完成',
  remark:      '備註'
};

function cellText(v){
  if(v===null||v===undefined)return'';
  return String(v).trim();
}

function main(){
  const sheet = SHEET_NAME ? Application.Sheets.Item(SHEET_NAME) : Application.Sheets.Item(1);
  const used = sheet.UsedRange;
  const values = used.Value2;           // 二維陣列
  if(!values || !values.length){ console.log('找不到資料'); return; }

  // 1) 找標題列：同時含「場次」與「備註」的那一列（掃前 5 列）
  let headerRow = -1, colIdx = {};
  for(let r=0; r<Math.min(5, values.length); r++){
    const row = values[r] || [];
    const idx = {};
    for(let c=0; c<row.length; c++){
      const t = cellText(row[c]);
      if(!t) continue;
      for(const key in COLS){
        if(idx[key]===undefined && t.indexOf(COLS[key])>=0) idx[key]=c;
      }
    }
    if(idx.session!==undefined && idx.remark!==undefined){ headerRow=r; colIdx=idx; break; }
  }
  if(headerRow<0){ console.log('找不到標題列（需含「場次」與「備註」欄）'); return; }
  console.log('標題列：第 '+(headerRow+1)+' 列，欄位對應：'+JSON.stringify(colIdx));

  // 2) 逐列讀取；場次為合併儲存格 → 空白時沿用上一列
  const get=(row,key)=>colIdx[key]===undefined?'':cellText(row[colIdx[key]]);
  const out=[];
  let curSession='';
  const syncedAt=new Date().toISOString();
  for(let r=headerRow+1; r<values.length; r++){
    const row=values[r]||[];
    const s=get(row,'session');
    if(s) curSession=s;
    if(!curSession) continue;
    // 整列（除場次外）皆空就跳過
    const fields=['vendor','cart','order_time','packing_site','checklist','arrival','counting','ship_counts','ship_start','ship_done','remark'];
    const hasData=s||fields.some(k=>get(row,k));
    if(!hasData) continue;
    const km=curSession.match(/\d{4}/);          // 場次開頭的4位數日期，用來對應看板場次
    out.push({
      id: (km?km[0]:'xxxx')+'_'+r,
      session_key: km?km[0]:'',
      session_name: curSession,
      closed: get(row,'closed'),
      vendor: get(row,'vendor'),
      cart: get(row,'cart'),
      order_time: get(row,'order_time'),
      packing_site: get(row,'packing_site'),
      checklist: get(row,'checklist'),
      arrival: get(row,'arrival'),
      counting: get(row,'counting'),
      ship_counts: get(row,'ship_counts'),
      ship_start: get(row,'ship_start'),
      ship_done: get(row,'ship_done'),
      remark: get(row,'remark'),
      synced_at: syncedAt
    });
  }
  console.log('讀到 '+out.length+' 列資料');
  if(!out.length)return;

  // 3) 全量更新：先清空再分批寫入
  const headers={
    'apikey':SUPA_KEY,
    'Authorization':'Bearer '+SUPA_KEY,
    'Content-Type':'application/json',
    'Prefer':'return=minimal'
  };
  let resp=HTTP.fetch(SUPA_URL+'/rest/v1/warehouse_sync?id=neq.__none__',{method:'DELETE',headers:headers});
  console.log('清空舊資料：HTTP '+resp.status);
  for(let i=0;i<out.length;i+=100){
    const batch=out.slice(i,i+100);
    resp=HTTP.fetch(SUPA_URL+'/rest/v1/warehouse_sync',{method:'POST',headers:headers,body:JSON.stringify(batch)});
    console.log('寫入 '+(i+batch.length)+'/'+out.length+'：HTTP '+resp.status);
    if(resp.status>=300){ console.log('寫入失敗：'+resp.text()); return; }
  }
  console.log('✅ 同步完成，共 '+out.length+' 列');
}

main();
