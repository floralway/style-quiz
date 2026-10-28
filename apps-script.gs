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
  type: '類型代號', lang: '語言', route: '路線', plan: '方案代號', focus: '主要卡點', companion: '自主／陪伴', strong: '強烈建議預諮詢', 'score_初綻': '初綻分數', q0: 'Q0 做過分析', b1: 'B1 分析後卡關', b2: 'B2 做過哪些分析', 'q9_想法': 'Q9 想法分類', 'q10_狀態': 'Q10 狀態分類', optin: '願意收穿搭分享信', mail_sent: '結果信寄出',
  q1: 'Q1 照鏡子', q2: 'Q2 買衣服', q3: 'Q3 別人怎麼說', q4: 'Q4 看到別人', q5: 'Q5 出門前',
  q6: 'Q6 正式穿搭', q7: 'Q7 換季整理', q8: 'Q8 最想改變', q9: 'Q9 想法', q10: 'Q10 狀態',
  age: '年齡', gender: '性別', role: '身分', region: '居住地區',
  name: '稱呼', email: 'Email', src: '來源', line_name: 'LINE 名稱', line_uid: 'LINE 用戶 ID'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.version === 'lecture' ? 'lecture' : 'website';   // 官網版與 LINE 版同一張表，用「來源」區分
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
  const order = ['萌芽期', '初綻期', '含苞期', '成束期'];
  const msg = order.map(k => k + '：' + (count[k] || 0) + ' 人').join('\n') + '\n\n共 ' + vals.length + ' 人';
  SpreadsheetApp.getUi().alert('講座版類型分布', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


/* ================= 結果信內容 =================
 * key：B 萌芽（方向）／H 初綻（顏色）／S 含苞（線條）／C 成束・路線 A（未做分析）
 *      C_PG 成束・路線 B 陪購／C_PR 成束・路線 B 客製陪跑／C_PGPR 兩者並列
 */
const RESULTS = {
  B: { stage: '萌芽期', name: '向光尋路的你', tag: '你正在找自己想長成的樣子。',
    why: '你的回答裡，最常出現的是「不像自己」「不知道該給人什麼感覺」。這代表你現在最值得先處理的，是方向。',
    now: '衣服本身好看，穿在你身上卻不太像你。別人眼中的你，跟你想給人的印象也有一點落差。你不缺衣服，缺的是方向。',
    gap: '一個清楚的「我想被怎麼記住」。有了方向，每一次選擇都會更篤定，也比較不會被流行推著走。',
    tr: '寫下三個你希望別人形容你的詞，例如「溫和、俐落、有想法」。下次買衣服或拍照前，先問自己：這樣穿，像不像這三個詞？' },
  H: { stage: '初綻期', name: '等待上色的你', tag: '底子已經在了，只差讓你發光的那個顏色。',
    why: '你的回答裡，最常出現的是「氣色差」「只敢穿安全色」。方向和版型目前不是你最大的困擾，先從顏色開始，改變最快、也最看得見。',
    now: '某些顏色一上身，整個人就少了精神。別人問你是不是沒睡好，其實你睡得很好。久了乾脆只穿黑白灰，安全，但也少了一點光。',
    gap: '不是更多衣服，而是知道哪些顏色站在你這邊。顏色離臉最近，是別人看你的第一眼，也是最快讓整個人亮起來的地方。',
    tr: '找一個有自然光的地方，拿兩件不同色調的上衣放在臉旁輪流比對，看哪一件讓膚色更透亮、黑眼圈更淡。先從口紅、圍巾、上衣這些靠近臉的單品開始換。' },
  S: { stage: '含苞期', name: '正在舒展姿態的你', tag: '找到撐得起你的線條，整個人就會挺拔起來。',
    why: '你的回答裡，最常出現的是「比例怪」「撐不起來」「版型難選」。顏色目前不是你最大的困擾，先從線條開始，效果最明顯。',
    now: '尺寸明明對，穿起來卻不對勁。別人穿很好看的版型，換到你身上就顯得呆板、太拘謹，或像被衣服吃掉。你可能也懷疑過是不是自己太瘦、太壯。',
    gap: '問題通常不在胖瘦，而在身體的線條、質感與比例，跟衣服的版型、材質合不合。合了，同樣的身材會穿出完全不同的樣子。',
    tr: '拿出你穿起來最好看的一件衣服，觀察它的材質是挺還是軟、長度落在哪裡、領口是什麼形狀。這就是你的身體給你的線索，下次買衣服先從相同的特徵找起。' },
  C: { stage: '成束期', name: '手握繁花的你', tag: '你身上已經有很多好花，只差把它們束成一束。',
    why: '你的回答裡，最常出現的是「單件好看、搭起來不對」「衣服很多、常穿的只有幾件」。你缺的不是衣服，是組合。',
    now: '你的衣櫃其實不少東西，也大概知道自己喜歡什麼。只是每天早上站在衣櫃前，還是要想很久才搭得出一套，最後又穿回那幾件。',
    gap: '有人在現場陪你判斷：哪些是值得買的關鍵單品，買回來怎麼跟你原本的衣服搭。',
    tr: '挑一件你最常穿的單品，用衣櫃裡現有的衣服搭出三套不同的穿法，拍照存起來，當作你自己的穿搭小本。' },
  C_PG: { stage: '成束期', name: '手握繁花的你', tag: '你身上已經有很多好花，只差把它們束成一束。',
    why: '你已經做過分析，知道自己適合什麼，卡住的是「怎麼挑、怎麼用」。',
    now: '你做過功課，手上也有分析結果，可是一走進店裡，還是不知道哪件該買；回到衣櫃前，也不確定結果要怎麼用在每天的穿搭上。',
    gap: '把分析結果變成現場的判斷。知道「適合什麼」和「看到這件該不該買」，中間差的是有人陪你練習一次。',
    tr: '拿出你的分析報告，對照衣櫃裡最常穿的五件衣服，標出哪些符合、哪些不符合。符合的那幾件，就是你下次挑衣服的參考樣本。' },
  C_PR: { stage: '成束期', name: '手握繁花的你', tag: '你身上已經有很多好花，只差把它們束成一束。',
    why: '你已經做過分析，但結果沒有真正落地到你的衣櫃和生活裡。',
    now: '你做過分析，也很認真想照著做。只是結果跟你現有的衣櫃對不上、太像一個結論，或做完之後覺得好像什麼都沒變，心裡有點空。',
    gap: '一個從你的衣櫃和生活出發的方法，而不是只給一個結論。分析要能回答「為什麼」和「怎麼用」，才會真的留在你身上。',
    tr: '回想分析之後，你最想穿、卻一直沒穿出去的那件衣服。寫下你沒穿它的原因，這通常就是卡住你的地方。' },
};
RESULTS.C_PGPR = Object.assign({}, RESULTS.C_PG, { why: '你已經做過分析，卡住的是「怎麼挑、怎麼用」，而且你也希望有人在旁邊陪你判斷。' });

const STAGE_EN = { '萌芽期': 'Sprouting', '初綻期': 'First Bloom', '含苞期': 'Budding', '成束期': 'Bouquet' };
const RESULTS_EN = {
  B: { name: 'Finding Your Way to the Light', tag: 'You’re discovering who you want to grow into.',
    why: 'Your answers kept coming back to “this isn’t me” and “I don’t know what impression to give.” Direction is the first thing worth working on.',
    now: 'The clothes look good, yet on you they don’t quite feel like you. The way others see you is a little different from the impression you want to give. You’re not short of clothes — you’re missing a direction.',
    gap: 'A clear sense of “how I want to be remembered.” With a direction, every choice feels more certain, and trends stop pushing you around.',
    tr: 'Write down three words you’d like people to use for you — say, “gentle, polished, thoughtful.” Before you buy something or take a photo, ask yourself: does this look like those three words?' },
  H: { name: 'Waiting for Your Color', tag: 'Everything’s already there — you just need the color that makes you glow.',
    why: 'Your answers kept pointing to a tired complexion and sticking to safe colors. Direction and fit aren’t your biggest issue right now — color is the quickest, most visible place to start.',
    now: 'Certain colors drain your energy the moment you put them on. People ask if you slept badly, even when you slept fine. Over time you’ve settled on black, white and gray — safe, but missing a little light.',
    gap: 'Not more clothes, but knowing which colors are on your side. Color sits closest to your face; it’s the first thing people see, and the quickest way to brighten your whole look.',
    tr: 'In natural light, hold two tops in different tones up to your face and compare: which one makes your skin look clearer and your under-eyes lighter? Start with the pieces closest to your face — lipstick, scarves, tops.' },
  S: { name: 'Stretching Into Your Shape', tag: 'Find the lines that suit you, and you’ll stand taller.',
    why: 'Your answers kept pointing to proportions, fit and cuts. Color isn’t your biggest issue right now — starting with lines will make the clearest difference.',
    now: 'The size is right, yet something’s off. Styles that look great on others turn stiff and boxy on you, or seem to swallow you up. You may have wondered whether you’re too thin or too broad.',
    gap: 'The issue is rarely your weight — it’s whether the cut and fabric match your body’s lines, texture and proportions. When they do, the same body looks completely different.',
    tr: 'Take out the piece you look best in. Notice whether the fabric is crisp or soft, where the length falls, and the shape of the neckline. These are your body’s clues — next time you shop, start with the same features.' },
  C: { name: 'Holding a Handful of Blooms', tag: 'You already have plenty of beautiful flowers — they just need to be gathered into a bouquet.',
    why: 'Your answers kept coming back to “each piece is fine but together it’s off” and “lots of clothes, only a few I wear.” You’re not missing clothes — you’re missing combinations.',
    now: 'You have quite a lot in your wardrobe and roughly know what you like. But standing there each morning, it still takes ages to put an outfit together — and you end up in the same few pieces.',
    gap: 'Someone beside you in the store to judge which key pieces are worth buying, and how they’ll work with what you already own.',
    tr: 'Pick the piece you wear most and build three different outfits around it, using only what’s already in your wardrobe. Photograph them — that’s the start of your own style notebook.' },
  C_PG: { name: 'Holding a Handful of Blooms', tag: 'You already have plenty of beautiful flowers — they just need to be gathered into a bouquet.',
    why: 'You’ve had an analysis and know what suits you — what’s stuck is choosing and using it.',
    now: 'You’ve done the homework and have your results, yet in a store you still can’t tell which piece to buy, and at your wardrobe you’re not sure how the results apply to everyday outfits.',
    gap: 'Turning your results into on-the-spot judgment. Between “what suits me” and “should I buy this?” is one round of practice with someone beside you.',
    tr: 'Take your analysis report and compare it with the five pieces you wear most. Mark which match and which don’t — the matches are your reference for what to choose next.' },
  C_PR: { name: 'Holding a Handful of Blooms', tag: 'You already have plenty of beautiful flowers — they just need to be gathered into a bouquet.',
    why: 'You’ve had an analysis, but the results never really landed in your wardrobe and life.',
    now: 'You had an analysis and genuinely wanted to follow it. But the results didn’t match your wardrobe, felt like a verdict, or left you feeling nothing had changed.',
    gap: 'A method that starts from your wardrobe and your life, not just a conclusion. An analysis only stays with you when it explains “why” and “how to use it.”',
    tr: 'Think of the piece you most wanted to wear after your analysis but never did. Write down why — that’s usually where you’re stuck.' },
};
RESULTS_EN.C_PGPR = Object.assign({}, RESULTS_EN.C_PG, { why: 'You’ve had an analysis; what’s stuck is choosing and using it — and you’d like someone beside you to help judge.' });

/* 官網版「如果想走得更確定」：依結果 × 性別（男＝男士，其餘＝女生／多元／不想透露） */
const PLAN = {
  zh: {
    F: {
      B: '建議從<b>全方位分析（色彩＋風格）</b>開始。方向要從了解自己的底子開始：先找出適合你的顏色，再整理出適合你、也像你的風格範圍，之後每一次選擇都有依據。',
      H: '建議從<b>個人色彩診斷</b>開始，這是最快讓整個人亮起來的一步。如果之後也想調整風格與版型，可以再進階到<b>全方位分析</b>。',
      S: '建議從<b>風格定位分析</b>開始，內容包含骨骼 12 分類、顏體平衡、臉診斷與風格限度定位®。一次把適合的版型、材質和風格範圍看清楚。',
      C: '建議從<b>陪同採購</b>開始：我會先依你的預算和需求走過一輪，當天你只要試穿、做決定，買回來的每一件都知道怎麼跟原本的衣服搭。',
      C_PG: '建議從<b>陪同採購</b>開始：帶著你的分析結果一起逛，在現場練習「看到這件，該不該買」，把結果真正用出來。',
      C_PR: '建議從<b>客製陪跑</b>開始。陪跑的分析比一般骨骼分析更細，而且是從你現有的衣櫃和生活出發，陪你在不同季節、不同場合一步步落地。',
      C_PGPR: '你可以從<b>陪同採購</b>開始，先把分析結果用在挑衣服上；如果你希望有人一路陪你判斷，也可以直接選擇<b>客製陪跑</b>。' },
    M: {
      B: '建議參考<b>男士形象陪跑</b>：先透過分析找到適合你的方向，再陪你把方向變成衣櫃裡真正會穿的衣服。',
      H: '如果你願意嘗試多樣的顏色，可以從<b>個人色彩診斷</b>開始，找出讓你氣色變好的顏色。如果你比較害羞、不想一下子穿得太亮，也可以直接做<b>風格定位分析</b>，透過材質和版型，創造出不同的感覺。',
      S: '建議從<b>風格定位分析</b>開始，內容包含骨骼 12 分類、顏體平衡、臉診斷與風格限度定位®，先看懂哪些版型和材質讓你穿起來更挺。',
      C: '建議從<b>陪同採購</b>開始。你其實知道很多，現場陪你挑、陪你試，直接把知道變成做到。',
      C_PG: '建議從<b>陪同採購</b>開始：帶著你的分析結果一起逛，現場陪你挑、陪你試，一次買對。',
      C_PR: '建議參考<b>男士形象陪跑</b>。陪跑的分析比一般骨骼分析更細，而且從你現有的衣櫃和生活出發，在你拿不定主意的時候，有人幫你判斷。',
      C_PGPR: '你可以從<b>陪同採購</b>開始；如果你希望有人一路陪你判斷，也可以直接選擇<b>男士形象陪跑</b>。' } },
  en: {
    F: {
      B: 'We suggest starting with a <b>Full Analysis (Color + Style)</b>. Direction starts with knowing your foundation: first find the colors that suit you, then map out a style range that suits you and feels like you.',
      H: 'We suggest starting with a <b>Personal Color Analysis</b> — the quickest way to brighten your whole look. If you later want to refine your style and cuts too, you can move on to a <b>Full Analysis</b>.',
      S: 'We suggest starting with a <b>Style Positioning Analysis</b>, covering 12-type body analysis, color–body balance, face analysis and Style Limit Positioning®.',
      C: 'We suggest starting with <b>Personal Shopping</b>: I walk the stores ahead based on your budget and needs, so on the day you just try on and decide — and know how every new piece works with what you own.',
      C_PG: 'We suggest starting with <b>Personal Shopping</b>: we bring your analysis results to the store and practice “should I buy this?” on the spot.',
      C_PR: 'We suggest starting with <b>Custom Coaching</b>. The coaching analysis goes deeper than a typical body-type analysis and starts from your current wardrobe and life, helping you apply it step by step across seasons and occasions.',
      C_PGPR: 'You can start with <b>Personal Shopping</b> to put your results to use; or, if you’d like someone beside you all the way, go straight to <b>Custom Coaching</b>.' },
    M: {
      B: 'Take a look at <b>Men’s Image Coaching</b>: we start with analysis to find the direction that suits you, then turn it into clothes you’ll actually wear.',
      H: 'If you’re open to trying a range of colors, start with a <b>Personal Color Analysis</b>. If you’d rather not go bold just yet, go straight to a <b>Style Positioning Analysis</b> and create a different feel through fabrics and cuts.',
      S: 'We suggest starting with a <b>Style Positioning Analysis</b>, covering 12-type body analysis, color–body balance, face analysis and Style Limit Positioning®.',
      C: 'We suggest starting with <b>Personal Shopping</b>. You already know a lot — choosing and trying on together turns knowing into doing.',
      C_PG: 'We suggest starting with <b>Personal Shopping</b>: we bring your analysis results to the store and get it right the first time.',
      C_PR: 'Take a look at <b>Men’s Image Coaching</b>. It goes deeper than a typical body-type analysis and starts from your wardrobe and life — someone to help you decide whenever you’re unsure.',
      C_PGPR: 'You can start with <b>Personal Shopping</b>; or, if you’d like someone beside you all the way, choose <b>Men’s Image Coaching</b>.' } },
};

/* 路線 A 陪伴型：在方案後加推陪跑 */
const ADDON = {
  zh: { F: '如果你希望有人一路陪著，也可以考慮<b>年度陪跑</b>：從分析開始，在不同季節、不同場合陪你練習，直到你能自己判斷。',
        M: '如果你希望有人一路陪著，也可以考慮<b>男士形象陪跑</b>，在你拿不定主意的時候，有人幫你判斷。' },
  en: { F: 'If you’d like someone beside you all the way, consider <b>Year-long Coaching</b>: starting from analysis, practicing across seasons and occasions until you can trust your own judgment.',
        M: 'If you’d like someone beside you all the way, consider <b>Men’s Image Coaching</b> — someone to help you decide whenever you’re unsure.' },
};

/* 強烈建議免費預諮詢的文案（官網版／LINE 版） */
const STRONG = {
  zh: { B: '你已經做過分析，卻還是卡在「怎麼用」。這通常不是你不夠努力，而是分析結果還沒對上你的衣櫃和生活。建議先預約<b>免費 30 分鐘預諮詢</b>，我們一起看看問題出在哪，再決定最適合你的下一步。',
        companion: '你希望有人陪你一起判斷，這很正常。建議先預約<b>免費 30 分鐘預諮詢</b>，我們先聊聊你現在的狀況，再一起決定從哪裡開始最適合你。',
        close: '你的回答裡，不只一個地方讓你卡住。這種情況自己很難判斷先處理哪一塊，建議先預約<b>免費 30 分鐘預諮詢</b>，我們一起排出最適合你的順序。',
        title: '建議你先免費聊 30 分鐘', btn: '預約免費預諮詢' },
  en: { B: 'You’ve already had an analysis, yet you’re still stuck on “how to use it.” That’s rarely about effort — usually the results just haven’t met your wardrobe and life yet. We suggest booking a <b>free 30-minute pre-consultation</b> first, so we can find where it’s stuck and choose the right next step.',
        companion: 'Wanting someone to help you decide is completely normal. We suggest booking a <b>free 30-minute pre-consultation</b> first — we’ll talk through where you are and decide together where to start.',
        close: 'Your answers show more than one thing holding you back, which makes it hard to know where to begin. We suggest booking a <b>free 30-minute pre-consultation</b> so we can map out the right order for you.',
        title: 'Let’s talk for 30 minutes first — free', btn: 'Book a free pre-consultation' },
};

/* 路線 B：從情境題看出的卡點 */
const FOCUS = {
  zh: { '萌芽期': '方向和風格', '初綻期': '顏色的運用', '含苞期': '版型和線條', '成束期': '單品之間的搭配' },
  en: { '萌芽期': 'direction and style', '初綻期': 'using color', '含苞期': 'cuts and lines', '成束期': 'putting pieces together' },
};
const FOCUS_LINE = {
  zh: f => `從你的回答來看，做完分析後你最常卡的還是「${f}」。`,
  en: f => `From your answers, what still trips you up most after your analysis is ${f}.`,
};

/* 男生版「這週可以先試試」：只覆寫有女性用語的時期 */
const TR_M = {
  zh: { H: '找一個有自然光的地方，拿兩件不同色調的上衣放在臉旁輪流比對，看哪一件讓膚色更透亮、黑眼圈更淡。先從襯衫、Polo 衫、外套這些靠近臉的單品開始換。' },
  en: { H: 'In natural light, hold two tops in different tones up to your face and compare: which one makes your skin look clearer and your under-eyes lighter? Start with the pieces closest to your face — shirts, polos, jackets.' },
};

const MAIL_TEXT = {
  zh: { header: 'FLORALWAY　穿搭方向測驗', hi: n => `${n} 你好，`,
    introW: s => `謝謝你花時間填寫這份測驗。看完你的答案，你現在的起點是「${s}」。<br>每個時期都是一個起點，不是等級，它代表你現在最值得先處理的那一塊。`,
    introL: s => `謝謝你今天參加講座，也謝謝你填寫這份測驗。你現在的起點是「${s}」。<br>每個時期都是一個起點，不是等級，它代表你現在最值得先處理的那一塊。`,
    why: '為什麼是這裡', now: '現在的你', gap: '你缺的那一塊', tr: '這週可以先試試', dir: '如果想走得更確定',
    closeW: '如果想先聊聊自己的狀況，歡迎預約預諮詢，我們先認識彼此，再決定下一步。', btnW: '預約預諮詢',
    closeL: '如果想多認識自己的花路，歡迎追蹤 @go_floralway，我會在那裡分享更多穿搭的小方法。', btnL: '追蹤 @go_floralway',
    foot: '每個人都在花路上的不同位置，沒有誰比較前面。', subject: n => n + '，你的穿搭可以先從這裡開始', you: '你' },
  en: { header: 'FLORALWAY　STYLE DIRECTION QUIZ', hi: n => `Hi ${n},`,
    introW: s => `Thank you for taking the quiz. Based on your answers, your starting point is the “${s}” stage.<br>Each stage is a starting point, not a level — it shows what’s most worth working on first.`,
    introL: s => `Thank you for joining today’s workshop and taking the quiz. Your starting point is the “${s}” stage.<br>Each stage is a starting point, not a level — it shows what’s most worth working on first.`,
    why: 'WHY HERE', now: 'WHERE YOU ARE NOW', gap: 'THE MISSING PIECE', tr: 'TRY THIS WEEK', dir: 'IF YOU’D LIKE MORE CERTAINTY',
    closeW: 'If you’d like to talk through your situation first, book a pre-consultation — we’ll get to know each other before deciding the next step.', btnW: 'Book a pre-consultation',
    closeL: 'To keep exploring your own path, follow @go_floralway — I share more simple styling tips there.', btnL: 'Follow @go_floralway',
    foot: 'Everyone is at a different place on their own path — no one is ahead.', subject: n => n + ', here’s where your style can begin', you: 'there' },
};

function MAIL_STRONG_TITLE(lang) { return STRONG[lang || 'zh'].title; }

function esc(s) { return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function sendResultMail(data) {
  const lang = data.lang === 'en' ? 'en' : 'zh';
  const M = MAIL_TEXT[lang];
  const type = data.type || 'C';
  const plan = data.plan || 'A';
  const key = (type === 'C' && plan !== 'A') ? 'C_' + plan : type;
  const g = data.gender === '男' ? 'M' : 'F';
  const zh = RESULTS[key];
  const t = lang === 'en'
    ? Object.assign({}, RESULTS_EN[key], { stage: STAGE_EN[zh.stage], stageZh: zh.stage })
    : Object.assign({}, zh, { stageZh: zh.stage });
  const routeB = String(data.route || '').indexOf('B') === 0;
  const companion = data.companion === '陪伴型';
  if (routeB && data.focus) t.why = FOCUS_LINE[lang](FOCUS[lang][data.focus] || data.focus) + ' ' + t.why;
  t.dir = PLAN[lang][g][key];
  if (!routeB && companion) t.dir += '<br><br>' + ADDON[lang][g];
  if (g === 'M' && TR_M[lang][key]) t.tr = TR_M[lang][key];
  t.strong = data.strong === '是' ? (routeB ? STRONG[lang].B : (companion ? STRONG[lang].companion : STRONG[lang].close)) : '';
  const rawName = data.name || M.you;
  const plain = s => String(s).replace(/<[^>]+>/g, '');
  GmailApp.sendEmail(data.email, M.subject(rawName),
    t.stage + '｜' + t.name + '\n' + t.tag + '\n\n' + M.why + '：' + plain(t.why) + '\n\n' + M.now + '：' + t.now + '\n\n' + M.gap + '：' + t.gap + '\n\n' + M.tr + '：' + t.tr,
    { htmlBody: buildMailHtml(t, type, data.version, esc(rawName), lang), name: SETTINGS.SENDER_NAME });
}

function buildMailHtml(t, type, version, name, lang) {
  const M = MAIL_TEXT[lang || 'zh'];
  const IMG = SETTINGS.IMG_BASE;
  const sp = lang === 'en' ? '2px' : '4px';
  const sec = (label, text) => `<tr><td style="padding:0 32px 22px"><div style="font-size:11px;letter-spacing:${sp};color:#E8784A;margin-bottom:8px">${label}</div><div style="font-size:15px;line-height:1.9;color:#6B5A52">${text}</div></td></tr>`;
  const intro = version !== 'lecture' ? M.introW(t.stage) : M.introL(t.stage);
  const order = ['萌芽期', '初綻期', '含苞期', '成束期'];
  const stages = order.map(s =>
    `<td align="center" style="padding:0 3px"><span style="display:inline-block;font-size:11px;padding:4px 9px;border-radius:999px;${s === t.stageZh ? 'background:#E8784A;color:#FFFFFF;border:1px solid #E8784A' : 'color:#A08070;border:1px solid #D6C4B7;background:#FBF6F1'}">${lang === 'en' ? STAGE_EN[s] : s}</span></td>`).join('');
  const dir = version !== 'lecture' ? sec(M.dir, t.dir) : '';
  const strongBox = (version !== 'lecture' && t.strong) ? `<tr><td style="padding:0 20px 24px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FBEEE6;border-radius:12px;border:1px solid #F2C7AE"><tr><td style="padding:20px 22px">
  <div style="font-size:17px;color:#3A2E28;margin-bottom:8px">${MAIL_STRONG_TITLE(lang)}</div>
  <div style="font-size:15px;line-height:1.9;color:#6B5A52">${t.strong}</div>
  <div style="text-align:center;padding-top:16px"><a href="${SETTINGS.BOOKING_URL}" style="display:inline-block;background:#E8784A;color:#fff;text-decoration:none;font-size:15px;letter-spacing:1px;padding:13px 30px;border-radius:999px">${STRONG[lang || 'zh'].btn}</a></div></td></tr></table></td></tr>` : '';
  const btn = (href, label, solid) => href ? `<a href="${href}" style="display:inline-block;${solid ? 'background:#E8784A;color:#fff' : 'background:#FFFCFA;color:#E8784A;border:1px solid #F2A882'};text-decoration:none;font-size:15px;letter-spacing:1px;padding:14px 32px;border-radius:999px">${label}</a>` : '';
  const close = version !== 'lecture'
    ? `<tr><td style="padding:4px 32px 8px;font-size:15px;line-height:1.9;color:#6B5A52">${M.closeW}</td></tr>
       <tr><td align="center" style="padding:14px 32px 30px">${btn(SETTINGS.BOOKING_URL, M.btnW, true)}</td></tr>`
    : `<tr><td style="padding:4px 32px 8px;font-size:15px;line-height:1.9;color:#6B5A52">${M.closeL}</td></tr>
       <tr><td align="center" style="padding:14px 32px 30px">${btn(SETTINGS.IG_URL, M.btnL, false)}</td></tr>`;
  return `<!doctype html><html><body style="margin:0;padding:0;background:#FAF6F3">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FAF6F3;font-family:Georgia,'Noto Serif TC','Noto Sans TC',sans-serif">
<tr><td align="center" style="padding:28px 12px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#FFFCFA;border-radius:16px;overflow:hidden;border-top:4px solid #E8784A">
<tr><td style="padding:0"><img src="${IMG}banner.png" width="600" alt="" style="display:block;width:100%;height:auto;border:0"></td></tr>
<tr><td style="padding:4px 32px 6px;font-size:11px;letter-spacing:${sp};color:#E8784A">${M.header}</td></tr>
<tr><td style="padding:14px 32px 18px;font-size:15px;line-height:1.9;color:#6B5A52">${M.hi(name)}<br>${intro}</td></tr>
<tr><td style="padding:0 20px 26px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F1E8E0;border-radius:14px;border:1px solid #E8D9CC"><tr><td align="center" style="padding:22px 18px 26px">
    <img src="${IMG}type-${type}.png" width="150" height="150" alt="" style="display:block;margin:0 auto 4px;border:0">
    <div style="font-size:12px;letter-spacing:${lang === 'en' ? '3px' : '6px'};color:#E8784A">${t.stage}</div>
    <div style="font-size:${lang === 'en' ? '23px' : '26px'};color:#3A2E28;margin:10px 0 8px">${t.name}</div>
    <div style="font-size:15px;font-style:italic;color:#8C7364;line-height:1.7">${t.tag}</div>
    <table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:18px"><tr>${stages}</tr></table>
  </td></tr></table>
</td></tr>
${strongBox}
${sec(M.why, t.why)}
${sec(M.now, t.now)}
${sec(M.gap, t.gap)}
<tr><td style="padding:0 20px 24px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FAF0E8;border-radius:12px"><tr><td style="padding:20px 22px">
  <div style="font-size:11px;letter-spacing:${sp};color:#E8784A;margin-bottom:8px">${M.tr}</div>
  <div style="font-size:15px;line-height:1.9;color:#3A2E28">${t.tr}</div></td></tr></table></td></tr>
${dir}
${close}
<tr><td style="padding:0 32px"><div style="height:1px;background:#EFE2D8"></div></td></tr>
<tr><td align="center" style="padding:22px 32px 30px;font-size:13px;line-height:1.9;color:#A08070;font-style:italic">${M.foot}<br><span style="font-style:normal;color:#3A2E28">Evelyn</span><br><span style="font-style:normal;font-size:12px;color:#C4B0A4;letter-spacing:1px">質穎形象顏究室® Floralway Image Studio</span></td></tr>
</table></td></tr></table></body></html>`;
}

/** 測試用：在 Apps Script 編輯器執行，寄測試信給你自己 */
function testMail() {
  sendResultMail({ type: 'S', plan: 'A', route: 'A 未做過分析', companion: '陪伴型', strong: '是', version: 'website', name: '測試', email: Session.getActiveUser().getEmail() });
}
function testMailRouteB() {
  sendResultMail({ type: 'C', plan: 'PR', route: 'B 做過分析', focus: '初綻期', companion: '自主型', strong: '是', version: 'website', name: '測試B', email: Session.getActiveUser().getEmail() });
}
