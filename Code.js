/**
 * ============================================================================
 * 台灣機械業 LINE 智能庫存查詢與自動 PDF 報價單系統 (Code.js)
 * ============================================================================
 * 遵循規範：
 * 1. .agent/rules/quote_safety.md（嚴禁揭露 cost_price 成本底價、驗證業務白名單、台灣機械業繁體中文術語）
 * 2. .agent/skills/quote_generator/SKILL.md（自然語言意圖解析、階梯式折扣計算、庫存交期預警、PDF 報價單與一鍵轉發客戶文案）
 */

// ==================== 1. 系統設定區塊 (Config) ====================

/** LINE Messaging API Channel Access Token（優先讀取 Secrets.js 或 ScriptProperties） */
const LINE_CHANNEL_ACCESS_TOKEN = (typeof SECRET_LINE_CHANNEL_ACCESS_TOKEN !== 'undefined' && SECRET_LINE_CHANNEL_ACCESS_TOKEN)
  ? SECRET_LINE_CHANNEL_ACCESS_TOKEN
  : (PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '請填入您的_LINE_CHANNEL_ACCESS_TOKEN');

/** Google Gemini API Key（優先讀取 Secrets.js 或 ScriptProperties） */
const GEMINI_API_KEY = (typeof SECRET_GEMINI_API_KEY !== 'undefined' && SECRET_GEMINI_API_KEY)
  ? SECRET_GEMINI_API_KEY
  : (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '請填入您的_GEMINI_API_KEY');

/** 使用之 Gemini 模型名稱 */
const GEMINI_MODEL = 'gemini-2.5-flash';

/**
 * 授權業務人員 LINE User ID 白名單 (ALLOWED_SALES_IDS)
 * 僅名列於此陣列中的業務人員可查詢庫存與開立報價單
 */
const ALLOWED_SALES_IDS = [
  'U60f4f5c5c7c193464151ea5e1195d6ab',
  'skipjackwu'
];

/** 庫存試算表分頁名稱（試算表：機台庫存表 / 分頁：工作表1） */
const SHEET_NAME = '工作表1';

/** Google Docs 報價單範本 ID */
const DOC_TEMPLATE_ID = '18oWrNBsOlivqyoGXf9xl9Huljjg8SatK4TC-LFPANcg';

/** 生成之 PDF 報價單存放 Google Drive 資料夾 ID */
const DRIVE_FOLDER_ID = '1r068e1m7NbYSE2WaGKSFrsqzM4NCmbnT';


// ==================== 2. LINE Webhook 主入口 ====================

/**
 * 提供瀏覽器或檢測工具確認 Webhook 服務連線狀態
 */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    system: '台灣機械業 LINE 智能庫存查詢與自動 PDF 報價單系統',
    message: 'Webhook 端點運作正常，請將此網址設定於 LINE Developers Messaging API Webhook URL。'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 處理 LINE Messaging API Webhook POST 請求
 * @param {GoogleAppsScript.Events.DoPost} e
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'no_data' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const payload = JSON.parse(e.postData.contents);
    const events = payload.events || [];

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      if (event.type !== 'message' || !event.message || event.message.type !== 'text') {
        continue;
      }

      const replyToken = event.replyToken;
      const userId = (event.source && event.source.userId) ? event.source.userId : '';
      const userMessage = (event.message.text || '').trim();

      // 【資安規則 2】驗證 ALLOWED_SALES_IDS 業務白名單
      if (!isAuthorizedSales(userId)) {
        const rejectMsg = [
          '⛔ 【系統權限攔截：未授權存取】',
          '抱歉，您的 LINE 帳號尚未列入本公司業務白名單（ALLOWED_SALES_IDS），無法查詢機台庫存或開立報價單。',
          '',
          '📌 您的 LINE User ID 為：',
          userId || '（無法取得 User ID）',
          '',
          '請將上方 User ID 複製並提供予系統管理員加入 Code.js 的 ALLOWED_SALES_IDS 白名單。'
        ].join('\n');

        replyLineMessages(replyToken, [rejectMsg]);
        continue;
      }

      // 執行業務詢價與庫存查詢主流程
      const replyMessages = processSalesInquiry(userMessage);
      replyLineMessages(replyToken, replyMessages);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('doPost 發生例外錯誤：', err);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 檢查發訊者是否位於 ALLOWED_SALES_IDS 業務白名單內
 * @param {string} userId
 * @return {boolean}
 */
function isAuthorizedSales(userId) {
  if (!userId) return false;
  return ALLOWED_SALES_IDS.indexOf(userId) !== -1;
}


// ==================== 3. 安全庫存讀取模組（物理剔除成本欄） ====================

/**
 * 讀取「工作表1」機台庫存資料，並強制剔除「標準成本 (機密欄位)」(cost_price)
 * 試算表欄位順序：
 * A欄(0): 品號 | B欄(1): 品名 | C欄(2): 規格說明 | D欄(3): 標準牌價 | E欄(4): 現有庫存 | F欄(5): 排產天數 | G欄(6): 標準成本 (機密欄位)
 *
 * @return {Array<Object>} 不含任何成本資訊之安全庫存陣列
 */
function getSafeInventoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('找不到指定的庫存工作表：「' + SHEET_NAME + '」，請確認分頁名稱是否正確。');
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return [];
  }

  const headers = rows[0].map(function(h) { return String(h).trim(); });

  // 動態定位各欄位索引（若標題列完全符合預設順序，則對應 0~5；嚴格排除「標準成本」與「cost_price」）
  const idxPartNo = findHeaderIndex(headers, ['品號', 'part_no', 'model_id'], 0);
  const idxName = findHeaderIndex(headers, ['品名', 'machine_name'], 1);
  const idxSpec = findHeaderIndex(headers, ['規格說明', '規格', 'spec'], 2);
  const idxListPrice = findHeaderIndex(headers, ['標準牌價', '牌價', 'list_price'], 3);
  const idxStock = findHeaderIndex(headers, ['現有庫存', '庫存', 'available_stock'], 4);
  const idxLeadTime = findHeaderIndex(headers, ['排產天數', '交期天數', 'lead_time_days'], 5);

  const safeInventory = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const partNo = String(row[idxPartNo] || '').trim();
    const machineName = String(row[idxName] || '').trim();

    if (!partNo && !machineName) {
      continue;
    }

    // 【資安規則 1】僅挑選公開業務欄位，物理剔除第 7 欄「標準成本 (機密欄位)」(cost_price)
    const safeItem = {
      part_no: partNo,
      machine_name: machineName,
      spec: String(row[idxSpec] || '').trim(),
      list_price: parseNumber(row[idxListPrice], 0),
      available_stock: parseNumber(row[idxStock], 0),
      lead_time_days: parseNumber(row[idxLeadTime], 45)
    };

    // 雙重確保絕不含 cost_price 屬性
    delete safeItem.cost_price;
    delete safeItem.standard_cost;

    safeInventory.push(safeItem);
  }

  return safeInventory;
}

/**
 * 輔助函式：尋找標題欄位索引
 */
function findHeaderIndex(headers, candidates, fallbackIndex) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toLowerCase();
    // 嚴格跳過任何包含「成本」或「cost」的欄位
    if (h.indexOf('成本') !== -1 || h.indexOf('cost') !== -1 || h.indexOf('底價') !== -1) {
      continue;
    }
    for (let j = 0; j < candidates.length; j++) {
      if (h === candidates[j].toLowerCase() || h.indexOf(candidates[j].toLowerCase()) !== -1) {
        return i;
      }
    }
  }
  return fallbackIndex;
}


// ==================== 4. 業務詢價主流程與意圖解析 ====================

/**
 * 處理已授權業務的 LINE 訊息，回傳要發送的 LINE 文字訊息陣列
 * @param {string} userMessage
 * @return {Array<string>}
 */
function processSalesInquiry(userMessage) {
  const safeInventory = getSafeInventoryData();
  if (safeInventory.length === 0) {
    return ['⚠️ 目前「' + SHEET_NAME + '」內尚無機台庫存資料，請先確認試算表內容。'];
  }

  // 若業務輸入「庫存」、「庫存清單」、「所有機台」等總覽指令，直接回傳全機台庫存總覽（不含成本）
  if (/^(查詢庫存|庫存清單|目前庫存|所有機台|機台清單|庫存|help|說明)$/i.test(userMessage)) {
    return [formatInventoryOverview(safeInventory)];
  }

  // 【SOP 1】透過 Gemini 解析自然語言意圖
  const intent = parseSalesIntentWithGemini(userMessage, safeInventory);

  // 【資安規則 1】若業務試圖套問進貨成本或底價 (is_asking_cost === true)，立即攔截
  if (intent.is_asking_cost) {
    return [
      [
        '🔒 【報價系統安全規範提醒】',
        '依據公司《報價系統安全規範》，系統已物理隔離並剔除「標準成本（底價）」欄位，嚴禁查詢或揭露機台進貨成本與毛利底價。',
        '',
        '💡 若需開立客戶報價單，請直接輸入客戶名稱、機台品號與需求台數（例如：「請幫台積電報價 2 台 CNC-L200」），系統將自動依授權階梯折數為您計算專案價並產出 PDF 報價單。'
      ].join('\n')
    ];
  }

  // 依據解析出的 keyword 比對庫存機台
  const matchedMachine = findMatchingMachine(intent.keyword, safeInventory);
  if (!matchedMachine) {
    return [
      [
        '🔍 查無符合關鍵字「' + (intent.keyword || userMessage) + '」的機台品號或品名。',
        '',
        formatInventoryOverview(safeInventory)
      ].join('\n')
    ];
  }

  // 【SOP 2 & SOP 3 & SOP 4】計算階梯式折扣、營業稅、庫存交期預警與客戶轉發文案
  const quotePayload = buildQuotePayload(intent, matchedMachine);

  // 產出 Google Docs 套印之 PDF 報價單並存入指定 Google Drive 資料夾
  let pdfInfo = null;
  let pdfErrorMsg = '';
  try {
    pdfInfo = generateQuotePdfFromTemplate(quotePayload.quote_data);
    quotePayload.quote_data.pdf_url = pdfInfo.pdfUrl;
    quotePayload.quote_data.pdf_file_id = pdfInfo.pdfFileId;
    quotePayload.quote_data.quote_no = pdfInfo.quoteNo;
  } catch (pdfErr) {
    console.error('生成 PDF 報價單失敗：', pdfErr);
    pdfErrorMsg = '⚠️ PDF 報價單生成異常：' + pdfErr.message;
  }

  // 第 1 則訊息：給業務人員內部檢視（含折扣明細、庫存交期預警、PDF 報價單連結與結構化摘要）
  const q = quotePayload.quote_data;
  const discountLabel = q.discount_rate === 1 ? '標準牌價 (100%)' : (q.discount_rate * 100) + ' 折';
  const internalSummaryLines = [
    '📋 【內部報價試算與庫存檢核結果】',
    '・報價對象：' + q.customer_name,
    '・機台品號：' + q.part_no,
    '・機台品名：' + q.machine_name,
    '・規格說明：' + q.spec,
    '・需求數量：' + q.qty + ' 台',
    '・標準牌價：NT$ ' + formatCurrency(q.list_price) + ' / 台',
    '・適用折數：' + discountLabel,
    '・專案未稅單價：NT$ ' + formatCurrency(q.unit_price) + ' / 台',
    '・未稅小計：NT$ ' + formatCurrency(q.subtotal),
    '・營業稅 (5%)：NT$ ' + formatCurrency(q.tax),
    '・含稅總金額：NT$ ' + formatCurrency(q.total),
    '・庫存與交期：' + q.stock_note
  ];

  if (pdfInfo) {
    internalSummaryLines.push('');
    internalSummaryLines.push('📄 【正式 PDF 報價單已產出】');
    internalSummaryLines.push('・報價單號：' + pdfInfo.quoteNo);
    internalSummaryLines.push('・雲端連結：' + pdfInfo.pdfUrl);
  } else if (pdfErrorMsg) {
    internalSummaryLines.push('');
    internalSummaryLines.push(pdfErrorMsg);
  }

  internalSummaryLines.push('');
  internalSummaryLines.push('👇 下一則訊息為「一鍵轉發客戶版商務文案」（已隱藏折數與庫存細節），可直接長按轉傳給客戶：');

  // 第 2 則訊息：一鍵轉發客戶版商務文案（若已產出 PDF 則自動附上 PDF 報價單連結）
  let customerCopyText = quotePayload.customer_copy;
  if (pdfInfo && pdfInfo.pdfUrl) {
    customerCopyText += '\n\n📎 正式報價單 PDF 下載連結：\n' + pdfInfo.pdfUrl;
  }

  return [
    internalSummaryLines.join('\n'),
    customerCopyText
  ];
}

/**
 * 【SOP 1】使用 Gemini API 解析業務自然語言訊息意圖
 * 萃取欄位：customer_name, keyword, qty, is_asking_cost
 *
 * @param {string} userMessage
 * @param {Array<Object>} safeInventory
 * @return {{customer_name: string, keyword: string, qty: number, is_asking_cost: boolean}}
 */
function parseSalesIntentWithGemini(userMessage, safeInventory) {
  // 先檢查訊息是否明顯包含套問成本/底價關鍵字（雙重防護）
  const costRegex = /(成本|底價|進價|進貨價|毛利|利潤|底線價|cost_price|cost)/i;
  const localAskingCost = costRegex.test(userMessage);

  // 若尚未設定 GEMINI_API_KEY，自動使用本地規則解析作為 Fallback
  if (!GEMINI_API_KEY || GEMINI_API_KEY.indexOf('請填入') !== -1) {
    return fallbackParseIntent(userMessage, safeInventory, localAskingCost);
  }

  const catalogList = safeInventory.map(function(item) {
    return item.part_no + ' (' + item.machine_name + ' / ' + item.spec + ')';
  }).join('、');

  const systemPrompt = [
    '你是台灣精密機械業的資深業務助理 AI。請從業務人員的 LINE 訊息中萃取詢價意圖，並嚴格只回傳 JSON 物件。',
    '目前公司現有機台清單為：' + catalogList,
    '',
    '請萃取以下 4 個欄位：',
    '1. "customer_name": 字串，客戶名稱（如「台積電」、「鴻海精密」、「宏達機械」）；若訊息未提及客戶名稱，請固定填寫 "貴賓客戶"。',
    '2. "keyword": 字串，業務想查詢或報價的機台品號或關鍵字（若能對應現有機台清單中的品號如 CNC-L200，請優先填寫該品號；否則填寫訊息中的機台關鍵字）。',
    '3. "qty": 整數，客戶需求台數（若訊息未指定台數，請固定填寫 1）。',
    '4. "is_asking_cost": 布林值（true 或 false），判斷業務是否試圖打聽、套問機台的「標準成本」、「進貨成本」、「底價」或「毛利」。'
  ].join('\n');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(GEMINI_API_KEY);
  const requestBody = {
    systemInstruction: {
      parts: [{ text: systemPrompt }]
    },
    contents: [
      {
        role: 'user',
        parts: [{ text: userMessage }]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });

    const resCode = response.getResponseCode();
    const resText = response.getContentText();
    if (resCode !== 200) {
      console.warn('Gemini API 回應非 200 (' + resCode + ')，改用本地解析：', resText);
      return fallbackParseIntent(userMessage, safeInventory, localAskingCost);
    }

    const resJson = JSON.parse(resText);
    const rawContent = resJson.candidates &&
      resJson.candidates[0] &&
      resJson.candidates[0].content &&
      resJson.candidates[0].content.parts &&
      resJson.candidates[0].content.parts[0].text;

    if (!rawContent) {
      return fallbackParseIntent(userMessage, safeInventory, localAskingCost);
    }

    const parsed = JSON.parse(rawContent);
    const qtyNum = parseInt(parsed.qty, 10);

    return {
      customer_name: (parsed.customer_name && String(parsed.customer_name).trim()) ? String(parsed.customer_name).trim() : '貴賓客戶',
      keyword: (parsed.keyword && String(parsed.keyword).trim()) ? String(parsed.keyword).trim() : userMessage,
      qty: (!isNaN(qtyNum) && qtyNum > 0) ? qtyNum : 1,
      is_asking_cost: Boolean(parsed.is_asking_cost) || localAskingCost
    };
  } catch (e) {
    console.warn('Gemini 意圖解析發生錯誤，啟用本地端備援解析：', e);
    return fallbackParseIntent(userMessage, safeInventory, localAskingCost);
  }
}

/**
 * 本地端備援自然語言解析（當 Gemini API Key 尚未設定或連線異常時使用）
 */
function fallbackParseIntent(userMessage, safeInventory, localAskingCost) {
  let matchedKeyword = '';
  const upperMsg = userMessage.toUpperCase();

  for (let i = 0; i < safeInventory.length; i++) {
    const item = safeInventory[i];
    if (item.part_no && upperMsg.indexOf(item.part_no.toUpperCase()) !== -1) {
      matchedKeyword = item.part_no;
      break;
    }
    if (item.machine_name && userMessage.indexOf(item.machine_name) !== -1) {
      matchedKeyword = item.part_no || item.machine_name;
      break;
    }
  }

  // 萃取台數（例如「3台」、「5 台」）
  let qty = 1;
  const qtyMatch = userMessage.match(/(\d+)\s*[台部套架]/);
  if (qtyMatch && qtyMatch[1]) {
    qty = Math.max(1, parseInt(qtyMatch[1], 10));
  }

  // 萃取客戶名稱（例如「幫台積電報價」、「客戶鴻海精密需要」）
  let customerName = '貴賓客戶';
  const custMatch = userMessage.match(/(?:幫|給|替|客戶|致)\s*([^\s，,。！!？?「」『』0-9]{2,15}?)(?:報價|開立|詢問|查詢|需要|買|採購|要)/);
  if (custMatch && custMatch[1]) {
    customerName = custMatch[1].trim();
  }

  return {
    customer_name: customerName,
    keyword: matchedKeyword || userMessage,
    qty: qty,
    is_asking_cost: Boolean(localAskingCost)
  };
}

/**
 * 依據關鍵字於安全庫存清單中比對最相符之機台
 * @param {string} keyword
 * @param {Array<Object>} safeInventory
 * @return {Object|null}
 */
function findMatchingMachine(keyword, safeInventory) {
  if (!keyword) return null;
  const kw = keyword.trim().toLowerCase();

  // 1. 品號完全相符或包含
  for (let i = 0; i < safeInventory.length; i++) {
    const item = safeInventory[i];
    if (item.part_no.toLowerCase() === kw || item.part_no.toLowerCase().indexOf(kw) !== -1 || kw.indexOf(item.part_no.toLowerCase()) !== -1) {
      return item;
    }
  }

  // 2. 品名或規格包含關鍵字
  for (let i = 0; i < safeInventory.length; i++) {
    const item = safeInventory[i];
    const nameLower = item.machine_name.toLowerCase();
    const specLower = item.spec.toLowerCase();
    if (nameLower.indexOf(kw) !== -1 || kw.indexOf(nameLower) !== -1 || specLower.indexOf(kw) !== -1) {
      return item;
    }
  }

  return null;
}


// ==================== 5. 階梯式折扣計算、庫存預警與商務文案生成 ====================

/**
 * 【SOP 2、SOP 3、SOP 4】計算階梯式折扣、5% 營業稅、庫存交期預警與一鍵轉發客戶文案
 *
 * @param {{customer_name: string, keyword: string, qty: number, is_asking_cost: boolean}} intent
 * @param {{part_no: string, machine_name: string, spec: string, list_price: number, available_stock: number, lead_time_days: number}} machine
 * @return {{intent: Object, quote_data: Object, customer_copy: string}}
 */
function buildQuotePayload(intent, machine) {
  const customerName = intent.customer_name || '貴賓客戶';
  const qty = Math.max(1, parseInt(intent.qty, 10) || 1);
  const listPrice = Math.round(Number(machine.list_price) || 0);
  const availableStock = Math.max(0, parseInt(machine.available_stock, 10) || 0);

  // 【SOP 2】階梯式折扣計算規則
  // 購買 1 台：折扣率 1.00（標準牌價 100%）
  // 購買 2 ～ 4 台：折扣率 0.95（95 折）
  // 購買 5 台（含）以上：折扣率 0.90（9 折）
  let discountRate = 1.00;
  if (qty >= 5) {
    discountRate = 0.90;
  } else if (qty >= 2 && qty <= 4) {
    discountRate = 0.95;
  }

  // 單價與總額一律四捨五入至整數，並外加 5% 營業稅計算 tax 與 total
  const unitPrice = Math.round(listPrice * discountRate);
  const subtotal = Math.round(unitPrice * qty);
  const tax = Math.round(subtotal * 0.05);
  const total = subtotal + tax;

  // 【SOP 3】庫存不足與交期預警
  let stockNote = '';
  let customerDeliveryNote = '';
  if (availableStock >= qty) {
    stockNote = '✅ 現貨充足（庫存 ' + availableStock + ' 台），下單後 7 個工作天內安排出機';
    customerDeliveryNote = '確認訂單後 7 個工作天內安排出機';
  } else {
    stockNote = '⚠️ 現貨不足（目前僅剩 ' + availableStock + ' 台），不足台數需排產 45 個工作天';
    customerDeliveryNote = '部分現貨可優先安排出機，預訂機台預計排產交期約 45 個工作天（實際排程依正式訂單確認為準）';
  }

  const quoteData = {
    customer_name: customerName,
    part_no: machine.part_no,
    machine_name: machine.machine_name,
    spec: machine.spec,
    qty: qty,
    list_price: listPrice,
    discount_rate: discountRate,
    unit_price: unitPrice,
    subtotal: subtotal,
    tax: tax,
    total: total,
    available_stock: availableStock,
    stock_note: stockNote
  };

  // 【SOP 4】一鍵轉發客戶版專業商務文案（嚴格隱藏折扣率與實際庫存台數）
  const customerCopy = [
    '【正式報價通知】',
    '致 ' + customerName + ' 採購先進 您好：',
    '',
    '感謝 貴司垂詢本公司精密機械設備，針對「' + machine.machine_name + '（品號：' + machine.part_no + '／規格：' + machine.spec + '）」需求數量 ' + qty + ' 台，謹提供專案優惠報價如下：',
    '',
    '・機台品號：' + machine.part_no,
    '・機台品名：' + machine.machine_name,
    '・規格說明：' + machine.spec,
    '・報價數量：' + qty + ' 台',
    '・專案優惠單價（未稅）：NT$ ' + formatCurrency(unitPrice) + ' / 台',
    '・未稅金額合計：NT$ ' + formatCurrency(subtotal),
    '・營業稅（5%）：NT$ ' + formatCurrency(tax),
    '・含稅總金額：NT$ ' + formatCurrency(total),
    '・預計出機交期：' + customerDeliveryNote,
    '',
    '詳細設備規格與正式報價條款請參閱附件 PDF 報價單。若有任何技術規格或商務問題，歡迎隨時與我們聯繫，誠摯期盼能為 貴司產線升級提供最優質的服務！'
  ].join('\n');

  return {
    intent: {
      customer_name: customerName,
      keyword: intent.keyword || machine.part_no,
      qty: qty,
      is_asking_cost: false
    },
    quote_data: quoteData,
    customer_copy: customerCopy
  };
}


// ==================== 6. Google Docs 範本套印轉 PDF 報價單模組 ====================

/**
 * 複製 Google Docs 報價單範本，替換 {{變數}} 後轉存為 PDF 放入指定 Google Drive 資料夾
 *
 * @param {Object} quoteData
 * @return {{quoteNo: string, pdfFileId: string, pdfUrl: string}}
 */
function generateQuotePdfFromTemplate(quoteData) {
  const now = new Date();
  const quoteDate = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd');
  const timestampStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd-HHmmss');
  const quoteNo = 'QT-' + timestampStr;

  const pdfFileName = '報價單_' + quoteData.customer_name + '_' + quoteData.part_no + '_' + timestampStr;

  const templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID);
  const targetFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // 1. 於目標資料夾建立暫存 Google Docs 複本
  const tempDocFile = templateFile.makeCopy(pdfFileName + '_TEMP', targetFolder);
  const tempDocId = tempDocFile.getId();

  try {
    const doc = DocumentApp.openById(tempDocId);
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();

    const discountText = quoteData.discount_rate === 1
      ? '標準牌價 (100%)'
      : (quoteData.discount_rate * 100) + ' 折';

    // 支援中英文常見 {{變數}} 佔位符，確保與範本完全相容
    const replacements = {
      '{{quote_no}}': quoteNo,
      '{{報價單號}}': quoteNo,
      '{{quote_date}}': quoteDate,
      '{{date}}': quoteDate,
      '{{報價日期}}': quoteDate,
      '{{日期}}': quoteDate,
      '{{customer_name}}': quoteData.customer_name,
      '{{客戶名稱}}': quoteData.customer_name,
      '{{part_no}}': quoteData.part_no,
      '{{model_id}}': quoteData.part_no,
      '{{品號}}': quoteData.part_no,
      '{{machine_name}}': quoteData.machine_name,
      '{{品名}}': quoteData.machine_name,
      '{{spec}}': quoteData.spec,
      '{{規格說明}}': quoteData.spec,
      '{{規格}}': quoteData.spec,
      '{{qty}}': String(quoteData.qty),
      '{{數量}}': String(quoteData.qty),
      '{{台數}}': String(quoteData.qty),
      '{{list_price}}': formatCurrency(quoteData.list_price),
      '{{標準牌價}}': formatCurrency(quoteData.list_price),
      '{{牌價}}': formatCurrency(quoteData.list_price),
      '{{discount_rate}}': discountText,
      '{{折扣率}}': discountText,
      '{{折數}}': discountText,
      '{{unit_price}}': formatCurrency(quoteData.unit_price),
      '{{優惠單價}}': formatCurrency(quoteData.unit_price),
      '{{單價}}': formatCurrency(quoteData.unit_price),
      '{{subtotal}}': formatCurrency(quoteData.subtotal),
      '{{未稅小計}}': formatCurrency(quoteData.subtotal),
      '{{小計}}': formatCurrency(quoteData.subtotal),
      '{{tax}}': formatCurrency(quoteData.tax),
      '{{營業稅}}': formatCurrency(quoteData.tax),
      '{{稅金}}': formatCurrency(quoteData.tax),
      '{{total}}': formatCurrency(quoteData.total),
      '{{含稅總額}}': formatCurrency(quoteData.total),
      '{{總計}}': formatCurrency(quoteData.total),
      '{{總金額}}': formatCurrency(quoteData.total),
      '{{available_stock}}': String(quoteData.available_stock),
      '{{現有庫存}}': String(quoteData.available_stock),
      '{{stock_note}}': quoteData.stock_note,
      '{{庫存備註}}': quoteData.stock_note,
      '{{交期說明}}': quoteData.stock_note,
      '{{備註}}': quoteData.stock_note
    };

    for (const placeholder in replacements) {
      if (Object.prototype.hasOwnProperty.call(replacements, placeholder)) {
        const escapedPattern = escapeRegExp(placeholder);
        const value = String(replacements[placeholder]);
        body.replaceText(escapedPattern, value);
        if (header) header.replaceText(escapedPattern, value);
        if (footer) footer.replaceText(escapedPattern, value);
      }
    }

    doc.saveAndClose();

    // 2. 將套印完成的暫存檔匯出為 PDF 並存入指定 Google Drive 資料夾
    const pdfBlob = tempDocFile.getAs(MimeType.PDF).setName(pdfFileName + '.pdf');
    const pdfFile = targetFolder.createFile(pdfBlob);

    // 設定知道連結者可檢視，方便業務直接轉發連結給客戶下載
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      console.warn('設定 PDF 公開檢視權限時發生警告（可能受網域政策限制）：', shareErr);
    }

    return {
      quoteNo: quoteNo,
      pdfFileId: pdfFile.getId(),
      pdfUrl: pdfFile.getUrl()
    };
  } finally {
    // 3. 刪除暫存 Google Docs 檔案
    try {
      tempDocFile.setTrashed(true);
    } catch (cleanupErr) {
      console.warn('清理暫存 Google Docs 檔案失敗：', cleanupErr);
    }
  }
}


// ==================== 7. LINE Messaging API 與通用輔助函式 ====================

/**
 * 呼叫 LINE Messaging API Reply 端點回覆訊息（最多支援 5 則氣泡訊息）
 * @param {string} replyToken
 * @param {Array<string>} textMessages
 */
function replyLineMessages(replyToken, textMessages) {
  if (!replyToken || !textMessages || textMessages.length === 0) return;

  const messagesPayload = textMessages.slice(0, 5).map(function(txt) {
    return {
      type: 'text',
      text: String(txt)
    };
  });

  const url = 'https://api.line.me/v2/bot/message/reply';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messagesPayload
    }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    console.error('LINE Reply API 發送失敗 (' + code + ')：', response.getContentText());
  }
}

/**
 * 格式化全機台庫存總覽（僅顯示公開欄位，嚴禁顯示標準成本）
 * @param {Array<Object>} safeInventory
 * @return {string}
 */
function formatInventoryOverview(safeInventory) {
  const lines = [
    '🏭 【目前機台庫存與標準牌價總覽】'
  ];

  for (let i = 0; i < safeInventory.length; i++) {
    const item = safeInventory[i];
    const stockBadge = item.available_stock > 0
      ? '✅ 現貨 ' + item.available_stock + ' 台'
      : '⚠️ 現貨 0 台（排產 ' + item.lead_time_days + ' 工作天）';

    lines.push(
      (i + 1) + '. 【' + item.part_no + '】' + item.machine_name + '\n' +
      '   ・規格：' + item.spec + '\n' +
      '   ・標準牌價：NT$ ' + formatCurrency(item.list_price) + '（未稅）\n' +
      '   ・庫存狀態：' + stockBadge
    );
  }

  lines.push('');
  lines.push('💡 開立報價單範例：「請幫台積電報價 2 台 ' + (safeInventory[0] ? safeInventory[0].part_no : 'CNC-L200') + '」');
  return lines.join('\n');
}

/**
 * 將數值轉為千分位字串
 * @param {number} num
 * @return {string}
 */
function formatCurrency(num) {
  const n = Math.round(Number(num) || 0);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * 將儲存格內容安全轉為數值
 */
function parseNumber(val, defaultVal) {
  if (typeof val === 'number' && !isNaN(val)) return val;
  const cleaned = String(val || '').replace(/[,$\sNT元台天]/gi, '');
  const parsed = Number(cleaned);
  return isNaN(parsed) ? defaultVal : parsed;
}

/**
 * 轉義正則表達式特殊字元（供 DocumentApp.Body.replaceText 使用）
 */
function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


// ==================== 8. GAS 編輯器內建測試函式 ====================

/**
 * 【測試函式】在 GAS 編輯器直接選擇 `testQuoteSystem` 點擊「執行」，
 * 即可驗證：
 * 1. 讀取「工作表1」並確認 `cost_price` / `標準成本` 已被完全剔除
 * 2. 驗證階梯式折扣計算與庫存預警
 * 3. 驗證 Google Docs 範本套印並於指定 Drive 資料夾產出 PDF 報價單
 */
function testQuoteSystem() {
  console.log('=== 1. 測試安全庫存讀取（確認剔除標準成本欄位） ===');
  const safeInventory = getSafeInventoryData();
  console.log('讀取筆數：' + safeInventory.length);
  console.log('第一筆安全庫存資料（不含成本）：', JSON.stringify(safeInventory[0], null, 2));

  if (safeInventory.length === 0) {
    console.warn('工作表1 目前無資料，結束測試。');
    return;
  }

  console.log('=== 2. 測試報價試算與 PDF 生成 ===');
  const sampleMachine = safeInventory[0];
  const sampleIntent = {
    customer_name: '台灣積體電路製造股份有限公司',
    keyword: sampleMachine.part_no,
    qty: 3,
    is_asking_cost: false
  };

  const payload = buildQuotePayload(sampleIntent, sampleMachine);
  console.log('結構化報價 JSON：', JSON.stringify(payload, null, 2));

  const pdfResult = generateQuotePdfFromTemplate(payload.quote_data);
  console.log('✅ PDF 報價單產出成功！', JSON.stringify(pdfResult, null, 2));
}
