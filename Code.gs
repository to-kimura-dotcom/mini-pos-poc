/** 学習用ミニPOS。公開デモ用。実売上や個人情報を保存しないこと。 */
const REPORT_TO = 'nacchanngogogo@gmail.com';
const ZONE = 'Asia/Tokyo';
const PRODUCT_HEADERS = ['id', 'name', 'price', 'active'];
const SALE_HEADERS = ['id', 'requestId', 'at', 'total', 'cash', 'change'];
const ITEM_HEADERS = ['saleId', 'productId', 'name', 'unitPrice', 'quantity', 'subtotal'];
const SEED = [['ブレンドコーヒー',400],['カフェラテ',480],['紅茶',380],['オレンジジュース',350],['チーズケーキ',520],['チョコケーキ',550],['クロワッサン',320],['サンドイッチ',650],['クッキー',180],['ミネラルウォーター',150]];

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SPREADSHEET_ID')) {
    const book = SpreadsheetApp.create('学習用ミニPOS 売上データ');
    props.setProperty('SPREADSHEET_ID', book.getId());
    const first = book.getSheets()[0]; first.setName('Products');
    first.appendRow(PRODUCT_HEADERS);
    first.getRange(2,1,SEED.length,4).setValues(SEED.map((p,i)=>[i+1,p[0],p[1],true]));
    const sales = book.insertSheet('Sales'); sales.appendRow(SALE_HEADERS);
    const items = book.insertSheet('Items'); items.appendRow(ITEM_HEADERS);
  }
  if (!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='sendReportIfDue')) {
    ScriptApp.newTrigger('sendReportIfDue').timeBased().everyMinutes(1).create();
  }
  Logger.log('Spreadsheet: https://docs.google.com/spreadsheets/d/'+props.getProperty('SPREADSHEET_ID'));
}
function book_() {
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id) throw new Error('管理者が setup を実行してください。');
  return SpreadsheetApp.openById(id);
}
function rows_(sheet) {
  const values=sheet.getDataRange().getValues();
  return values.slice(1);
}
function products_() {
  return rows_(book_().getSheetByName('Products')).map(r=>({id:Number(r[0]),name:String(r[1]),price:Number(r[2]),active:r[3]===true||String(r[3]).toLowerCase()==='true'}));
}
function getState() {
  const b=book_(),products=products_(),items=rows_(b.getSheetByName('Items'));
  const bySale={}; for(const r of items){const id=Number(r[0]);(bySale[id]??=[]).push({productId:Number(r[1]),name:String(r[2]),unitPrice:Number(r[3]),quantity:Number(r[4]),subtotal:Number(r[5])})}
  const sales=rows_(b.getSheetByName('Sales')).map(r=>({id:Number(r[0]),at:String(r[2]),total:Number(r[3]),cash:Number(r[4]),change:Number(r[5]),items:bySale[Number(r[0])]||[]}));
  return {products,sales};
}
function saveSale(requestId, cart, cash) {
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    if(typeof requestId!=='string'||!/^[a-zA-Z0-9-]{12,80}$/.test(requestId))throw new Error('会計識別子が不正です。');
    if(!Array.isArray(cart)||!cart.length||cart.length>50)throw new Error('商品を選んでください。');
    if(!Number.isSafeInteger(Number(cash))||Number(cash)<0)throw new Error('預かり金額が不正です。');
    const b=book_(),saleSheet=b.getSheetByName('Sales'),existing=rows_(saleSheet);
    const duplicated=existing.find(r=>String(r[1])===requestId);
    if(duplicated)return {duplicate:true,state:getState(),saleId:Number(duplicated[0])};
    const products=new Map(products_().map(p=>[p.id,p]));
    const ids=new Set(),items=[];let total=0;
    for(const entry of cart){
      const id=Number(entry.id),quantity=Number(entry.quantity),p=products.get(id);
      if(!p||!p.active||ids.has(id)||!Number.isInteger(quantity)||quantity<1||quantity>99)throw new Error('カートを確認してください。');
      ids.add(id);const subtotal=p.price*quantity;total+=subtotal;
      items.push({productId:id,name:p.name,unitPrice:p.price,quantity,subtotal});
    }
    if(!Number.isSafeInteger(total)||Number(cash)<total)throw new Error('預かり金額が不足しています。');
    const usedIds=existing.map(r=>Number(r[0])).concat(rows_(b.getSheetByName('Items')).map(r=>Number(r[0])));
    const id=usedIds.length?Math.max(...usedIds)+1:1;
    const at=new Date().toISOString();
    // 明細を先に保存し、最後に会計ヘッダーを保存する。ヘッダーがない明細は集計対象外。
    b.getSheetByName('Items').getRange(b.getSheetByName('Items').getLastRow()+1,1,items.length,6).setValues(items.map(i=>[id,i.productId,i.name,i.unitPrice,i.quantity,i.subtotal]));
    saleSheet.appendRow([id,requestId,at,total,Number(cash),Number(cash)-total]);
    return {saleId:id,state:getState()};
  } finally {lock.releaseLock()}
}
function saveProduct(id,name,price) {
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    name=String(name||'').trim();price=Number(price);
    if(/[\x00-\x1F\x7F]/.test(name))throw new Error('商品名に改行や制御文字は使えません。');
    if(!name||name.length>60||!Number.isSafeInteger(price)||price<1||price>999999)throw new Error('商品名と税込価格を確認してください。');
    const sheet=book_().getSheetByName('Products'),all=products_();id=Number(id);
    if(id){const idx=all.findIndex(p=>p.id===id);if(idx<0)throw new Error('商品が見つかりません。');sheet.getRange(idx+2,2,1,2).setValues([[name,price]])}
    else{const nextId=all.length?Math.max(...all.map(p=>p.id))+1:1;sheet.appendRow([nextId,name,price,true])}
    return getState();
  }finally{lock.releaseLock()}
}
function toggleProduct(id) {
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {const all=products_(),idx=all.findIndex(p=>p.id===Number(id));if(idx<0)throw new Error('商品が見つかりません。');book_().getSheetByName('Products').getRange(idx+2,4).setValue(!all[idx].active);return getState()}finally{lock.releaseLock()}
}
function reportData_(day){
 const state=getState();const sales=state.sales.filter(s=>Utilities.formatDate(new Date(s.at),ZONE,'yyyy-MM-dd')===day);
 const total=sales.reduce((n,s)=>n+s.total,0),products=new Map();
 for(const s of sales)for(const i of s.items){const key=i.productId+'|'+i.name+'|'+i.unitPrice;const p=products.get(key)||{name:i.name,units:0,total:0};p.units+=i.quantity;p.total+=i.subtotal;products.set(key,p)}
 return {sales,total,products:[...products.values()].sort((a,b)=>b.total-a.total)};
}
function sendReportIfDue(){
 const now=new Date(),day=Utilities.formatDate(now,ZONE,'yyyy-MM-dd'),hour=Number(Utilities.formatDate(now,ZONE,'H'));
 if(hour!==17)return;
 const lock=LockService.getScriptLock();if(!lock.tryLock(10000))return;
 try{
  const props=PropertiesService.getScriptProperties();if(props.getProperty('LAST_REPORT_DAY')===day)return;
  const d=reportData_(day),lines=[day+' の売上レポート','',`売上合計：${d.total.toLocaleString('ja-JP')}円`,`会計件数：${d.sales.length}件`,`客単価：${d.sales.length?Math.round(d.total/d.sales.length).toLocaleString('ja-JP')+'円':'—'}`,'','商品別販売数と売上額'];
  if(!d.products.length)lines.push('売上はありません。');
  for(const p of d.products)lines.push(`${p.name}：${p.units}点／${p.total.toLocaleString('ja-JP')}円`);
  lines.push('','※17時時点の集計です。17時以降の会計は含まれません。','※学習用PoCのデータです。');
  MailApp.sendEmail({to:REPORT_TO,subject:`【ミニPOS】${day} 17時売上レポート`,body:lines.join('\n')});
  props.setProperty('LAST_REPORT_DAY',day);
 }finally{lock.releaseLock()}
}
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('ミニPOS PoC')}
