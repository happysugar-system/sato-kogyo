/**
 * 佐藤工業 / 協力大工応募受付
 * 1. standalone Apps Script に貼り付ける
 * 2. setupApplication() を一度実行し、権限を承認する
 * 3. ウェブアプリとしてデプロイ（実行ユーザー:自分 / アクセス:全員）
 * 4. /exec URLをGitHubの recruit.html の WEB_APP_URL に設定
 * 個人情報を扱うため、作成されたスプレッドシートは公開しないこと。
 */
const SK = {
  spreadsheetProperty: 'SATO_KOGYO_APPLICATION_SHEET',
  destinationEmail: 'sato.kogyo.chiba@gmail.com',
  sheetName: '協力大工応募',
  fields: ['受付日時','受付ID','氏名・屋号','連絡先','大工経験年数','対応可能エリア','経験のある工事','希望人工代','希望・質問','対応状況'],
  allowedParent: 'https://happysugar-system.github.io'
};

function setupApplication() {
  const props=PropertiesService.getScriptProperties();
  let id=props.getProperty(SK.spreadsheetProperty);
  if(!id){
    const book=SpreadsheetApp.create('佐藤工業_協力大工応募管理');
    id=book.getId();
    props.setProperty(SK.spreadsheetProperty,id);
  }
  const book=SpreadsheetApp.openById(id);
  let sheet=book.getSheetByName(SK.sheetName);
  if(!sheet) {
    sheet=book.getSheets()[0];
    sheet.setName(SK.sheetName);
  }
  sheet.getRange(1,1,1,SK.fields.length).setValues([SK.fields]);
  sheet.getRange(1,1,1,SK.fields.length).setFontWeight('bold').setBackground('#173450').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  Logger.log('メール送信上限（残数）: '+MailApp.getRemainingDailyQuota());
  Logger.log('応募管理シート: '+book.getUrl());
  return book.getUrl();
}

function clean_(value, max) {
  const s=String(value==null?'':value).trim().slice(0,max);
  // Prevent spreadsheet formula injection.
  return /^[=+\-@]/.test(s) ? "'"+s : s;
}

function response_(token, ok, message) {
  // use JSON encoding to safely interpolate into HTML script.
  const data=JSON.stringify({type:'sato-kogyo-recruit-result',token:token,ok:ok,message:message});
  const html='<!doctype html><html><head><meta charset="utf-8"></head><body>'+
    '<script>try{window.top.postMessage('+data+','+JSON.stringify(SK.allowedParent)+')}catch(e){};try{window.parent.postMessage('+data+','+JSON.stringify(SK.allowedParent)+')}catch(e){};<\/script>'+
    '<p>'+ (ok?'受付しました。':'送信を完了できませんでした。') +'</p></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e){
  const p=e&&e.parameter?e.parameter:{};
  const token=String(p.requestToken||'').slice(0,90);
  // If name is filled in the hidden spam-trap input, do not store anything.
  if(p.website) return response_(token,false,'受付できませんでした。');
  const name=clean_(p.name,100),contact=clean_(p.contact,150);
  if(!token||!name||!contact) return response_(token,false,'お名前と連絡先を確認してください。');
  const items=[
    clean_(p.years,30),clean_(p.area,200),clean_(p.works,300),
    clean_(p.price,100),clean_(p.message,2000)
  ];
  if(String(p.consent||'')!=='yes')return response_(token,false,'個人情報の取扱いへの同意が必要です。');
  const props=PropertiesService.getScriptProperties();
  const id=props.getProperty(SK.spreadsheetProperty);
  if(!id) return response_(token,false,'受付システムが未設定です。');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(10000)) return response_(token,false,'受付が混み合っています。しばらくしてから再度お試しください。');
  try {
    const cache=CacheService.getScriptCache();
    const key='req_'+Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,token).map(x=>(x+256)%256).map(x=>x.toString(16).padStart(2,'0')).join('').slice(0,40);
    if(cache.get(key))return response_(token,true,'すでに受付済みです。');
    // Basic spam throttling per contact. Public websites should add Turnstile verification before wider promotion.
    const contactHash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,contact.toLowerCase()).map(x=>(x+256)%256).map(x=>x.toString(16).padStart(2,'0')).join('').slice(0,40);
    const contactKey='contact_'+contactHash;
    if(cache.get(contactKey))return response_(token,false,'短時間に繰り返し応募できません。少し時間をおいてください。');
    const sheet=SpreadsheetApp.openById(id).getSheetByName(SK.sheetName);
    if(!sheet) return response_(token,false,'受付シートが見つかりません。');
    const row=[new Date(),token,name,contact,...items,'未対応'];
    sheet.appendRow(row);
    cache.put(key,'1',21600);
    cache.put(contactKey,'1',90);
    const body=[
      '佐藤工業HPから協力大工の問い合わせがありました。','',
      ...SK.fields.slice(1,9).map((field,i)=>field+': '+String(row[i+1])),
      '', '応募管理シート: '+SpreadsheetApp.openById(id).getUrl()
    ].join('\n');
    try { MailApp.sendEmail(SK.destinationEmail,'【佐藤工業】協力大工の新規応募',body); }
    catch(mailErr) { console.error('受付保存済み・通知メール送信失敗: '+mailErr); }
    return response_(token,true,'受付しました。ご連絡をお待ちください。');
  } catch(err) {
    console.error('応募受付エラー:'+err);
    return response_(token,false,'受付できませんでした。時間をおいて再度お試しください。');
  } finally {lock.releaseLock();}
}
