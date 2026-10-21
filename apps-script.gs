/**
 * 花路穿搭問卷｜資料收集後台
 * 用法：
 * 1. 開一份新的 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上這整段程式碼並儲存
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *    執行身分：我　／　誰可以存取：所有人
 * 4. 複製網頁應用程式網址，貼到兩個問卷 HTML 裡的 ENDPOINT = "" 之中
 *
 * 5. 把下方 SETTINGS 的 IMG_BASE 改成你放圖片的 GitHub Pages 網址（結尾要有 /）
 *    並填上預諮詢、LINE 連結
 * 6. 第一次部署時 Google 會要求授權「以你的名義寄信」，請允許
 *
 * 講座版與官網版會分別寫入「lecture」「website」兩個工作表，
 * 欄位會依收到的資料自動建立，不需要事先設定表頭。
 * 每份問卷送出時，會依結果自動寄出對應的結果信。
 *
 * 注意：一般 Gmail 帳號透過 Apps Script 每天約可寄 100 封；
 * Google Workspace 帳號約 1,500 封。講座人數多時請留意。
 */

const SETTINGS = {
  IMG_BASE: 'https://floralway.github.io/style-quiz/img/',
  SENDER_NAME: 'Evelyn｜質穎形象顏究室',
  BOOKING_URL: 'https://calendly.com/floralway-image/30min',
  IG_URL: 'https://www.instagram.com/go_floralway/',
};

const HEADER_LABELS = {
  time: '填寫時間', version: '版本', result: '結果類型', branch: '成束支線',
  'score_萌芽': '萌芽分數', 'score_舒枝': '舒枝分數', 'score_含苞': '含苞分數', 'score_成束': '成束分數',
  type: '類型代號', 'q9_想法': 'Q9 想法分類', 'q10_狀態': 'Q10 狀態分類', optin: '願意收穿搭分享信', mail_sent: '結果信寄出',
  q1: 'Q1 照鏡子', q2: 'Q2 買衣服', q3: 'Q3 別人怎麼說', q4: 'Q4 看到別人', q5: 'Q5 出門前',
  q6: 'Q6 正式穿搭', q7: 'Q7 換季整理', q8: 'Q8 最想改變', q9: 'Q9 想法', q10: 'Q10 狀態',
  age: '年齡', gender: '性別', role: '身分', region: '居住地區',
  name: '稱呼', email: 'Email'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.version === 'website' ? 'website' : 'lecture';
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // 第 1 列存欄位代號（程式用），第 2 列顯示中文欄名
    let keys = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
      : [];
    if (keys.length === 0) {
      keys = Object.keys(data);
    } else {
      Object.keys(data).forEach(k => { if (keys.indexOf(k) === -1) keys.push(k); });
    }
    const headerRow = keys.map(k => HEADER_LABELS[k] || k);
    sheet.getRange(1, 1, 1, keys.length).setValues([keys]);

    sheet.getRange(2, 1, 1, keys.length).setValues([headerRow]).setFontWeight('bold');
    sheet.setFrozenRows(2);

    let sent = '';
    try {
      if (data.email) { sendResultMail(data); sent = '已寄出'; }
    } catch (mailErr) { sent = '寄送失敗：' + mailErr; }
    data.mail_sent = sent;
    if (keys.indexOf('mail_sent') === -1) {
      keys.push('mail_sent');
      sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
      sheet.getRange(2, 1, 1, keys.length).setValues([keys.map(k => HEADER_LABELS[k] || k)]).setFontWeight('bold');
    }
    const row = keys.map(k => (data[k] === undefined ? '' : data[k]));
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** 講座現場用：在試算表選單快速看各類型人數 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('花路問卷')
    .addItem('統計講座版類型分布', 'showLectureSummary')
    .addToUi();
}

function showLectureSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('lecture');
  if (!sheet || sheet.getLastRow() < 3) {
    SpreadsheetApp.getUi().alert('目前還沒有講座版的資料');
    return;
  }
  const keys = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = keys.indexOf('result') + 1;
  const vals = sheet.getRange(3, col, sheet.getLastRow() - 2, 1).getValues().map(r => r[0]);
  const count = {};
  vals.forEach(v => { count[v] = (count[v] || 0) + 1; });
  const order = ['萌芽期', '舒枝期', '含苞期', '成束期'];
  const msg = order.map(k => k + '：' + (count[k] || 0) + ' 人').join('\n') + '\n\n共 ' + vals.length + ' 人';
  SpreadsheetApp.getUi().alert('講座版類型分布', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


/* ================= 結果信 ================= */

const RESULTS = {
  B: { stage: '萌芽期', name: '向光尋路的你', tag: '你正在找自己想長成的樣子。',
    now: '衣服本身好看，穿在你身上卻不太像你。別人眼中的你，跟你想給人的印象也有一點落差。你不缺衣服，缺的是方向。',
    gap: '一個清楚的「我想被怎麼記住」。有了方向，每一次選擇都會更篤定，也比較不會被流行推著走。',
    tr: '寫下三個你希望別人形容你的詞，例如「溫和、俐落、有想法」。下次買衣服或拍照前，先問自己：這樣穿，像不像這三個詞？',
    dir: '風格定位會結合你的骨架、臉型與色彩，找出「適合你、也像你」的風格範圍，讓方向不再只靠感覺。' },
  S: { stage: '舒枝期', name: '正在舒展姿態的你', tag: '找到撐得起你的線條，整個人就會挺拔起來。',
    now: '尺寸明明對，穿起來卻時好時壞。別人穿很好看的版型，換到你身上就顯得呆板、太拘謹，或像被衣服吃掉。你可能也懷疑過是不是自己太瘦、太壯。',
    gap: '問題通常不在胖瘦，而在身體的線條、質感與比例，跟衣服的版型、材質合不合。合了，同樣的身材會穿出完全不同的樣子。',
    tr: '拿出你穿起來最好看的一件衣服，觀察它的材質是挺還是軟、長度落在哪裡、領口是什麼形狀。這就是你的身體給你的線索，下次買衣服先從相同的特徵找起。',
    dir: '骨架診斷能幫你看懂自己的線條與比例，知道哪些版型和材質，能讓你穿出最好的樣子。' },
  H: { stage: '含苞期', name: '等待上色的你', tag: '底子已經在了，只差讓你發光的那個顏色。',
    now: '某些顏色一上身，整個人就少了精神。別人問你是不是沒睡好，其實你睡得很好。久了乾脆只穿黑白灰，安全，但也少了一點光。',
    gap: '不是更多衣服，而是知道哪些顏色站在你這邊。顏色離臉最近，是別人看你的第一眼，也是最快讓整個人亮起來的地方。',
    tr: '找一個有自然光的地方，拿兩件不同色調的上衣放在臉旁輪流比對，看哪一件讓膚色更透亮、黑眼圈更淡。先從口紅、圍巾、上衣這些靠近臉的單品開始換。',
    dir: '個人色彩診斷可以幫你找出最適合的色彩範圍，也會告訴你，那些「很想穿但不在範圍內」的顏色可以怎麼調整著穿。' },
  C_D1: { stage: '成束期', name: '手握繁花的你', tag: '你身上已經有很多好花，只差把它們束成一束。',
    now: '你其實做過不少功課：可能做過分析、看過很多穿搭內容，也知道自己喜歡什麼。只是這些答案和建議散在各處，回到每天早上的衣櫃前，還是不知道怎麼組合。',
    gap: '一套把「知道」變成「做到」的方法。與其再收集更多建議，不如把手上的花整理成一束。',
    tr: '挑一件你最常穿的單品，用衣櫃裡現有的衣服搭出三套不同的穿法，拍照存起來，當作你自己的穿搭小本。',
    dir: '衣櫃諮詢會陪你把現有的衣服重新整理、配對；需要補單品時，陪同採購可以陪你現場判斷什麼值得買。' },
  C_D2: { stage: '成束期', name: '手握繁花的你', tag: '你身上已經有很多好花，只差把它們束成一束。',
    now: '穿搭對你來說，可能不是不懂，而是常常沒把握：別人一句話就動搖，下定決心要改變，過一陣子又回到原點，或是一直交給家人、朋友幫你決定。',
    gap: '一個能慢慢建立自己判斷力的過程，而且在你動搖的時候，旁邊有人可以穩住你。',
    tr: '這個月，每次覺得「今天穿得不錯」時拍一張照，寫下一句為什麼。一個月後回頭看，你會開始看見屬於自己的規律。',
    dir: '年度陪跑是一整年的陪伴：從分析開始，一起採購、一起調整，在不同季節、不同場合陪你練習，直到你能自己判斷。' },
};

function esc(s) { return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function sendResultMail(data) {
  const type = data.type || 'C';
  const key = type === 'C' ? (data.branch === 'D2' ? 'C_D2' : 'C_D1') : type;
  const t = RESULTS[key];
  const name = esc(data.name || '你');
  const subject = (data.name || '你') + '，你的穿搭可以先從這裡開始';
  GmailApp.sendEmail(data.email, subject,
    t.stage + '｜' + t.name + '\n' + t.tag + '\n\n現在的你：' + t.now + '\n\n你缺的那一塊：' + t.gap + '\n\n這週可以先試試：' + t.tr,
    { htmlBody: buildMailHtml(t, type, data.version, name), name: SETTINGS.SENDER_NAME });
}

function buildMailHtml(t, type, version, name) {
  const IMG = SETTINGS.IMG_BASE;
  const sec = (label, text) => `<tr><td style="padding:0 32px 22px"><div style="font-size:11px;letter-spacing:4px;color:#E8784A;margin-bottom:8px">${label}</div><div style="font-size:15px;line-height:1.9;color:#6B5A52">${text}</div></td></tr>`;
  const intro = version === 'website'
    ? `謝謝你花時間填寫這份測驗。看完你的答案，你現在在「${t.stage}」。`
    : `謝謝你今天參加講座，也謝謝你填寫這份測驗。你現在在「${t.stage}」。`;
  const stages = ['萌芽期', '舒枝期', '含苞期', '成束期'].map(s =>
    `<td align="center" style="padding:0 3px"><span style="display:inline-block;font-size:11px;padding:4px 9px;border-radius:999px;${s === t.stage ? 'background:#E8784A;color:#FFFFFF;border:1px solid #E8784A' : 'color:#A08070;border:1px solid #D6C4B7;background:#FBF6F1'}">${s}</span></td>`).join('');
  const dir = version === 'website' ? sec('如果想走得更確定', t.dir) : '';
  const btn = (href, label, solid) => href ? `<a href="${href}" style="display:inline-block;${solid ? 'background:#E8784A;color:#fff' : 'background:#FFFCFA;color:#E8784A;border:1px solid #F2A882'};text-decoration:none;font-size:15px;letter-spacing:2px;padding:14px 32px;border-radius:999px">${label}</a>` : '';
  const close = version === 'website'
    ? `<tr><td style="padding:4px 32px 8px;font-size:15px;line-height:1.9;color:#6B5A52">如果想先聊聊自己的狀況，歡迎預約預諮詢，我們先認識彼此，再決定下一步。</td></tr>
       <tr><td align="center" style="padding:14px 32px 30px">${btn(SETTINGS.BOOKING_URL, '預約預諮詢', true)}</td></tr>`
    : `<tr><td style="padding:4px 32px 8px;font-size:15px;line-height:1.9;color:#6B5A52">如果想多認識自己的花路，歡迎追蹤 @go_floralway，我會在那裡分享更多穿搭的小方法。</td></tr>
       <tr><td align="center" style="padding:14px 32px 30px">${btn(SETTINGS.IG_URL, '追蹤 @go_floralway', false)}</td></tr>`;
  return `<!doctype html><html><body style="margin:0;padding:0;background:#FAF6F3">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FAF6F3;font-family:Georgia,'Noto Serif TC','Noto Sans TC',sans-serif">
<tr><td align="center" style="padding:28px 12px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#FFFCFA;border-radius:16px;overflow:hidden;border-top:4px solid #E8784A">
<tr><td style="padding:0"><img src="${IMG}banner.png" width="600" alt="" style="display:block;width:100%;height:auto;border:0"></td></tr>
<tr><td style="padding:4px 32px 6px;font-size:11px;letter-spacing:4px;color:#E8784A">FLORALWAY　穿搭方向測驗</td></tr>
<tr><td style="padding:14px 32px 18px;font-size:15px;line-height:1.9;color:#6B5A52">${name} 你好，<br>${intro}</td></tr>
<tr><td style="padding:0 20px 26px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F1E8E0;border-radius:14px;border:1px solid #E8D9CC"><tr><td align="center" style="padding:22px 18px 26px">
    <img src="${IMG}type-${type}.png" width="150" height="150" alt="" style="display:block;margin:0 auto 4px;border:0">
    <div style="font-size:12px;letter-spacing:6px;color:#E8784A">${t.stage}</div>
    <div style="font-size:26px;color:#3A2E28;margin:10px 0 8px">${t.name}</div>
    <div style="font-size:15px;font-style:italic;color:#8C7364;line-height:1.7">${t.tag}</div>
    <table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:18px"><tr>${stages}</tr></table>
  </td></tr></table>
</td></tr>
${sec('現在的你', t.now)}
${sec('你缺的那一塊', t.gap)}
<tr><td style="padding:0 20px 24px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FAF0E8;border-radius:12px"><tr><td style="padding:20px 22px">
  <div style="font-size:11px;letter-spacing:4px;color:#E8784A;margin-bottom:8px">這週可以先試試</div>
  <div style="font-size:15px;line-height:1.9;color:#3A2E28">${t.tr}</div></td></tr></table></td></tr>
${dir}
${close}
<tr><td style="padding:0 32px"><div style="height:1px;background:#EFE2D8"></div></td></tr>
<tr><td align="center" style="padding:22px 32px 30px;font-size:13px;line-height:1.9;color:#A08070;font-style:italic">每個人都在花路上的不同位置，沒有誰比較前面。<br><span style="font-style:normal;color:#3A2E28">Evelyn</span><br><span style="font-style:normal;font-size:12px;color:#C4B0A4;letter-spacing:1px">質穎形象顏究室® Floralway Image Studio</span></td></tr>
</table></td></tr></table></body></html>`;
}

/** 測試用：在 Apps Script 編輯器執行，會寄一封含苞期官網版的測試信給你自己 */
function testMail() {
  sendResultMail({ type: 'H', version: 'website', name: '測試', email: Session.getActiveUser().getEmail() });
}
