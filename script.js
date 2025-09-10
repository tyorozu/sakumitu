import { initializeApp } from "https://www.gstatic.com/firebasejs/9.15.0/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/9.15.0/firebase-auth.js";


// login.htmlからコピーしてきたFirebase設定情報
const firebaseConfig = {
  apiKey: "AIzaSyB0fKHkhF9Bbb1EmIrNCcnBRX5LUuCAafQ",
  authDomain: "sakupro-c57b4.firebaseapp.com",
  projectId: "sakupro-c57b4",
  storageBucket: "sakupro-c57b4.firebasestorage.app",
  messagingSenderId: "193298223172",
  appId: "1:193298223172:web:85963a63a13055b7adac11",
  measurementId: "G-7S8LV0GW0E"
};

// Firebaseの初期化
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// ログイン状態を監視する
onAuthStateChanged(auth, (user) => {
    const loginTimestamp = localStorage.getItem('loginTimestamp');
    const threeDaysInMillis = 3 * 24 * 60 * 60 * 1000; // 3日をミリ秒に変換

    if (user && loginTimestamp && (Date.now() - loginTimestamp > threeDaysInMillis)) {
    // ログインしているが、最終ログインから3日以上経過している場合
        console.log('セッションがタイムアウトしました。自動的にログアウトします。');
        localStorage.removeItem('loginTimestamp'); // 記録を削除
        signOut(auth); // ログアウトを実行
        alert('セッションがタイムアウトしました。再度ログインしてください。');
        // signOutが実行されると onAuthStateChanged が再度呼ばれ、
        // userがいない状態になるので自動的にログインページに遷移します。
        return; // これ以降の処理は行わない
    }
    
  if (user) {
    // ログイン済みのユーザーがいる場合
    console.log("ログイン済みです:", user.email);
    // ★元々の初期化処理をここで呼び出します
    initializeAppContent(); 
  } else {
    // ログインしていない場合
    console.log("ログインしていません。ログインページに移動します。");
    // ログインページにリダイレクト
    window.location.href = 'login.html'; 
  }
});

// ★元々の初期化処理を関数にまとめます
async function initializeAppContent() {
  try {
    await loadDataFromCSV();
    init();
    setupEventListeners();
  } catch (error) {
    console.error("アプリケーションの初期化に失敗しました:", error);
    alert(`データの読み込みに失敗しました。CSVファイルが正しく配置されているか確認してください。\n\nエラー詳細: ${error.message}`);
  }
}


// =================================================================================
// グローバル変数
// =================================================================================
let priceData = {}; // CSVから読み込んだ単価データ
let memberList = []; // CSVから読み込んだ担当者リスト
const TRANSPORT_COST_PER_PERSON_DAY = 5000; // 1人日あたりの交通費(仕入)

// =================================================================================
// アプリケーション起動処理
// =================================================================================

async function loadDataFromCSV() {
    const [priceCsvText, memberCsvText] = await Promise.all([
        fetch('prices.csv').then(res => res.ok ? res.text() : Promise.reject(new Error(`prices.csvが見つかりません: ${res.statusText}`))),
        fetch('members.csv').then(res => res.ok ? res.text() : Promise.reject(new Error(`members.csvが見つかりません: ${res.statusText}`)))
    ]);
    priceData = parsePriceCSV(priceCsvText);
    memberList = parseMemberCSV(memberCsvText);
}

function init() {
    setupSalesOfficeDropdown();
    updateMemberDropdowns(Object.keys(priceData)[0] || '');
    addMaterialRow();
    updateAllPrices(); 
    addExpenseRow(false, { item: '夜間割り増し', spec: '', unit: '式', price: 0 });
    addExpenseRow(false, { item: '廃棄費用', spec: '', unit: '式', price: 0 });
    calculateAll();
}

function setupEventListeners() {
    // ★変更: イベントリスナーの対象を body に変更して動的要素に対応
    document.body.addEventListener('input', (e) => {
        if (e.target.matches('input:not(#total-purchase-cost-input), select')) {
            calculateAll();
        }
    });
    document.getElementById('total-purchase-cost-input').addEventListener('change', () => {
        recalculateProfit();
    });
    document.getElementById('sales-office').addEventListener('change', (e) => {
        // 1. CSV由来のデータを更新
        updateAllPrices();
    
        // ★★★ 変更点 ★★★
        // 2. 自動計算される経費行を再追加する
        addExpenseRow(false, { item: '夜間割り増し', spec: '', unit: '式', price: 0 });
        addExpenseRow(false, { item: '廃棄費用', spec: '', unit: '式', price: 0 });

        // 3. 担当者ドロップダウンを更新
        updateMemberDropdowns(e.target.value);
        
        // 4. 全体を再計算
        calculateAll();
    });

    document.getElementById('add-work-row-btn').addEventListener('click', () => addWorkRow(true));
    document.getElementById('add-material-row-btn').addEventListener('click', addMaterialRow);
    document.getElementById('add-expense-row-btn').addEventListener('click', () => addExpenseRow(true));
    document.getElementById('export-excel-btn').addEventListener('click', exportToExcel);

    const logoutBtn = document.getElementById('logout-btn');
    if (logoutBtn) { // ボタンが存在する場合のみリスナーを設定
        logoutBtn.addEventListener('click', () => {
            signOut(auth).then(() => {
                // ログアウト成功
                console.log('ログアウトしました。');
                // login.htmlへリダイレクト（onAuthStateChangedが検知して自動で飛びますが、念のため）
                window.location.href = 'login.html';
            }).catch((error) => {
                // ログアウト失敗
                console.error('Logout Error:', error);
                alert('ログアウトに失敗しました。');
            });
        });
    }
}

// =================================================================================
// UI生成・更新関連の関数
// =================================================================================

function setupSalesOfficeDropdown() {
    const salesOfficeSelect = document.getElementById('sales-office');
    salesOfficeSelect.innerHTML = '';
    Object.keys(priceData).forEach(officeName => {
        const option = document.createElement('option');
        option.value = officeName;
        option.textContent = officeName;
        salesOfficeSelect.appendChild(option);
    });
}

function updateMemberDropdowns(selectedOffice) {
    const myNameSelect = document.getElementById('my-name');
    const approverNameSelect = document.getElementById('approver-name');
    myNameSelect.innerHTML = '';
    approverNameSelect.innerHTML = '';

    const noApproverOption = document.createElement('option');
    noApproverOption.value = "";
    noApproverOption.textContent = "(検印者なし)";
    approverNameSelect.appendChild(noApproverOption);

    const creators = memberList.filter(m => m.office === selectedOffice && m.roles.includes('担当'));
    // ★変更: 検印者は全営業所から選択できるように、営業所の絞り込みを解除
    const approvers = memberList.filter(m => m.roles.includes('検印'));

    if (creators.length === 0) {
        myNameSelect.innerHTML = '<option>該当メンバーなし</option>';
    } else {
        creators.forEach(member => {
            const option = document.createElement('option');
            option.value = member.name;
            option.textContent = member.name;
            myNameSelect.appendChild(option);
        });
    }

    approvers.forEach(member => {
        const option = document.createElement('option');
        option.value = member.name;
        option.textContent = member.name;
        approverNameSelect.appendChild(option);
    });
}

function updateAllPrices() {
    const selectedOffice = document.getElementById('sales-office').value;
    const officeData = priceData[selectedOffice] || { '工事': [], '経費': [] };

    // 工事テーブルをクリアしてCSVデータを再設定
    const workTbody = document.querySelector('#work-items-table tbody');
    workTbody.innerHTML = '';
    officeData['工事'].forEach(item => addWorkRow(false, item));

    // 経費テーブルをクリアしてCSVデータを再設定
    const expenseTbody = document.querySelector('#expense-items-table tbody');
    expenseTbody.innerHTML = '';
    officeData['経費'].forEach(item => addExpenseRow(false, item));
}

// ★変更: 廃棄費用に対応
function addWorkRow(isManual, item = { item: '', spec: '', unit: '', price: 0, disposal_price: 0, assumed_minutes: 0, purchase_price: 0 }) {
    const tbody = document.querySelector('#work-items-table tbody');
    const row = tbody.insertRow();
    const readonlyAttr = !isManual ? 'readonly' : '';
    row.innerHTML = `
        <td data-minutes="${item.assumed_minutes || 0}" data-purchase-price="${item.purchase_price || 0}"><input type="text" class="item-name" value="${item.item}" ${readonlyAttr}></td>
        <td><input type="text" class="spec" value="${item.spec}" ${readonlyAttr}></td>
        <td><input type="text" class="unit" value="${item.unit}" ${readonlyAttr}></td>
        <td><input type="number" class="editable-price" value="${item.price}"></td>
        <td><input type="number" class="quantity" min="0" value="0"></td>
        <td class="total">¥0</td>
        <td><input type="number" class="disposal-price editable-price" value="${item.disposal_price || 0}"></td>
    `;
}

function addMaterialRow() {
    const tbody = document.querySelector('#material-items-table tbody');
    const row = tbody.insertRow();
    row.innerHTML = `
        <td><input type="text" class="name"></td>
        <td><input type="text" class="spec"></td>
        <td><input type="number" class="quantity" min="0" value="1"></td>
        <td><input type="text" class="unit" value="台"></td>
        <td><input type="number" class="unit-price" min="0" value="0"></td>
        <td><input type="number" class="purchase-price" min="0" value="0"></td> <!-- ★これが追加された仕入単価の列です -->
        <td class="total">¥0</td>
        <td><input type="number" class="list-price" min="0" value="0"></td>
    `;
}

function addExpenseRow(isManual, item = { item: '', spec: '', unit: '', price: 0, purchase_price: 0 }) {
    const tbody = document.querySelector('#expense-items-table tbody');

    const isDisposalRow = item.item === '廃棄費用';
    const isNightCostRow = item.item === '夜間割り増し';
    const isAutoCalcRow = isDisposalRow || isNightCostRow;

    const row = isAutoCalcRow ? tbody.insertRow(0) : tbody.insertRow();
    
    if (isDisposalRow) row.id = 'disposal-cost-row';
    if (isNightCostRow) row.id = 'night-cost-row';

    const checkboxState = isManual ? 'checked' : ''; 
    const autoCalcReadonlyAttr = 'readonly';

    if (isAutoCalcRow) {
        const specHtml = isNightCostRow
            ? `工事費の <input type="number" class="rate-input night-rate-input" value="20" min="0" style="width: 60px;"> %`
            : `<input type="text" class="spec" value="${item.spec}" ${autoCalcReadonlyAttr}>`;

        row.innerHTML = `
            <td class="col-check"><input type="checkbox" ${checkboxState}></td>
            <td data-purchase-price="${item.purchase_price || 0}"><input type="text" class="item-name" value="${item.item}" ${autoCalcReadonlyAttr}></td>
            <td>${specHtml}</td>
            <td><input type="text" class="unit" value="${item.unit}" ${autoCalcReadonlyAttr}></td>
            <td><input type="number" class="editable-price" value="${item.price}" ${autoCalcReadonlyAttr}></td>
            <td><input type="number" class="quantity" min="1" value="1" ${autoCalcReadonlyAttr}></td>
            <td class="total">¥0</td>
        `;
    } else {
        const readonlyAttr = isManual ? '' : 'readonly';
        row.innerHTML = `
            <td class="col-check"><input type="checkbox" ${checkboxState}></td>
            <td data-purchase-price="${item.purchase_price || 0}"><input type="text" class="item-name" value="${item.item}" ${readonlyAttr}></td>
            <td><input type="text" class="spec" value="${item.spec}" ${readonlyAttr}></td>
            <td><input type="text" class="unit" value="${item.unit}" ${readonlyAttr}></td>
            <td><input type="number" class="editable-price" value="${item.price}"></td>
            <td><input type="number" class="quantity" min="1" value="1"></td>
            <td class="total">¥0</td>
        `;
    }
}
// =================================================================================
// 計算関連の関数
// =================================================================================

function calculateAll() {
    // --- 1. 変数初期化 ---
    let totalWorkCost = 0, totalMaterialCost = 0, totalExpenseCost = 0;
    let totalWorkPurchaseCost = 0, totalMaterialPurchaseCost = 0, totalExpensePurchaseCost = 0;
    let totalDisposalCost = 0;
    let totalMinutes = 0;

    // --- 2. 工事項目から各種数値を計算 ---
    document.querySelectorAll('#work-items-table tbody tr').forEach(row => {
        const quantity = parseFloat(row.querySelector('.quantity').value) || 0;
        const price = parseFloat(row.querySelector('.editable-price').value) || 0;
        const minutesPerUnit = parseInt(row.querySelector('td[data-minutes]').dataset.minutes) || 0;
        const purchasePrice = parseFloat(row.querySelector('td[data-purchase-price]').dataset.purchasePrice) || 0;
        const total = price * quantity;
        row.querySelector('.total').textContent = formatCurrency(total);
        totalWorkCost += total;
        totalWorkPurchaseCost += purchasePrice * quantity;
        if (quantity > 0) {
            totalMinutes += quantity * minutesPerUnit;
            const disposalPrice = parseFloat(row.querySelector('.disposal-price').value) || 0;
            totalDisposalCost += disposalPrice * quantity;
        }
    });

    // --- 3. 作業日数を計算し、画面に反映 ---
    const manpower = parseFloat(document.getElementById('assumed-manpower').value) || 1;
    const MINUTES_PER_DAY = 480;
    const calculatedDays = totalMinutes > 0 ? Math.ceil((totalMinutes / manpower) / MINUTES_PER_DAY) : 0;
    document.getElementById('work-days').value = calculatedDays;

    // --- 4. 自動計算される経費行の値を更新 ---
    const disposalRow = document.getElementById('disposal-cost-row');
    if (disposalRow) {
        disposalRow.querySelector('.editable-price').value = totalDisposalCost;
        const checkbox = disposalRow.querySelector('input[type="checkbox"]');
        checkbox.disabled = totalDisposalCost <= 0;
        if (totalDisposalCost <= 0) checkbox.checked = false;
    }
    const nightCostRow = document.getElementById('night-cost-row');
    if (nightCostRow) {
        const nightRateValue = parseFloat(nightCostRow.querySelector('.night-rate-input').value) || 0;
        const calculatedNightCost = totalWorkCost * (nightRateValue / 100);
        nightCostRow.querySelector('.editable-price').value = calculatedNightCost;
        const checkbox = nightCostRow.querySelector('input[type="checkbox"]');
        checkbox.disabled = calculatedNightCost <= 0;
        if (calculatedNightCost <= 0) checkbox.checked = false;
    }

    // --- 5. 部材費 & 部材仕入費の計算 ---
    document.querySelectorAll('#material-items-table tbody tr').forEach(row => {
        const quantity = parseFloat(row.querySelector('.quantity').value) || 0;
        const unitPrice = parseFloat(row.querySelector('.unit-price').value) || 0;
        const purchasePrice = parseFloat(row.querySelector('.purchase-price').value) || 0;
        const total = quantity * unitPrice;
        row.querySelector('.total').textContent = formatCurrency(total);
        totalMaterialCost += total;
        totalMaterialPurchaseCost += quantity * purchasePrice;
    });

    // --- 6. 経費 & 経費仕入費の計算 ---
    document.querySelectorAll('#expense-items-table tbody tr').forEach(row => {
        if (row.querySelector('input[type="checkbox"]').checked) {
            const quantity = parseFloat(row.querySelector('.quantity').value) || 1;
            const price = parseFloat(row.querySelector('.editable-price').value) || 0;
            const purchasePrice = parseFloat(row.querySelector('td[data-purchase-price]').dataset.purchasePrice) || price;
            const total = price * quantity;
            row.querySelector('.total').textContent = formatCurrency(total);
            totalExpenseCost += total;
            totalExpensePurchaseCost += purchasePrice * quantity;
        } else {
            row.querySelector('.total').textContent = formatCurrency(0);
        }
    });

    // --- 7. 共通項目 & 見積合計金額の計算 (売値) ---
    const workDays = parseFloat(document.getElementById('work-days').value) || 0;
    const cleaningRate = parseFloat(document.getElementById('cleaning-rate').value) || 0;
    const suppliesRate = parseFloat(document.getElementById('supplies-rate').value) || 0;
    const transportUnitPrice = parseFloat(document.getElementById('transport-unit-price').value) || 0;
    const overheadRate = parseFloat(document.getElementById('overhead-rate').value) || 0;
    const welfareRate = parseFloat(document.getElementById('welfare-rate').value) || 0;
    const cleaningCost = totalWorkCost * (cleaningRate / 100);
    const suppliesCost = totalWorkCost * (suppliesRate / 100);
    const transportCost = workDays * transportUnitPrice;
    const subtotal = totalWorkCost + totalMaterialCost + totalExpenseCost;
    const subtotal2 = subtotal + cleaningCost + suppliesCost + transportCost;
    const overheadCost = subtotal2 * (overheadRate / 100);
    const subtotal3 = subtotal2 + overheadCost;
    const welfareCost = subtotal3 * (welfareRate / 100);
    const grandTotal = subtotal3 + welfareCost;

    // --- 8. 仕入合計の計算 ---
    const transportPurchaseCost = manpower * workDays * TRANSPORT_COST_PER_PERSON_DAY;
    const totalPurchaseCost = totalWorkPurchaseCost + totalMaterialPurchaseCost + totalExpensePurchaseCost + transportPurchaseCost;

    // --- 9. 全ての計算結果を画面に表示 ---
    // (中略)...
    document.getElementById('total-work-cost').textContent = formatCurrency(totalWorkCost);
    document.getElementById('total-material-cost').textContent = formatCurrency(totalMaterialCost);
    document.getElementById('total-expense-cost').textContent = formatCurrency(totalExpenseCost);
    document.getElementById('subtotal').textContent = formatCurrency(subtotal);
    document.getElementById('cleaning-cost').textContent = formatCurrency(cleaningCost);
    document.getElementById('supplies-cost').textContent = formatCurrency(suppliesCost);
    document.getElementById('transport-cost').textContent = formatCurrency(transportCost);
    document.getElementById('subtotal2').textContent = formatCurrency(subtotal2);
    document.getElementById('overhead-cost').textContent = formatCurrency(overheadCost);
    document.getElementById('subtotal3').textContent = formatCurrency(subtotal3);
    document.getElementById('welfare-cost').textContent = formatCurrency(welfareCost);
    document.getElementById('grand-total').textContent = formatCurrency(grandTotal);
    
    // 仕入関連
    document.getElementById('total-work-purchase-cost').textContent = formatCurrency(totalWorkPurchaseCost);
    document.getElementById('total-material-purchase-cost').textContent = formatCurrency(totalMaterialPurchaseCost);
    document.getElementById('total-expense-purchase-cost').textContent = formatCurrency(totalExpensePurchaseCost);
    document.getElementById('transport-purchase-cost').textContent = formatCurrency(transportPurchaseCost);
    document.getElementById('total-purchase-cost-input').value = totalPurchaseCost;

    // 粗利を再計算
    recalculateProfit();
}

/**
 * 粗利金額と粗利率だけを再計算して表示する関数
 * 合計仕入が手動で変更された時に呼び出される
 */
function recalculateProfit() {
    const grandTotal = parseCurrency(document.getElementById('grand-total').textContent);
    const totalPurchaseCost = parseFloat(document.getElementById('total-purchase-cost-input').value) || 0;
    const grossProfitAmount = grandTotal - totalPurchaseCost;
    const grossProfitMargin = grandTotal > 0 ? (grossProfitAmount / grandTotal) * 100 : 0;
    document.getElementById('gross-profit-amount').textContent = formatCurrency(grossProfitAmount);
    document.getElementById('gross-profit-margin').textContent = `${grossProfitMargin.toFixed(2)}%`;
}

// =================================================================================
// ヘルパー関数
// =================================================================================

// ★変更: 廃棄費用 `disposal_price` も読み込む
function parsePriceCSV(csvText) {
    const data = {};
    const lines = csvText.trim().split('\n');
    const headers = lines.shift().trim().split(',');
    lines.forEach(line => {
        const values = line.trim().split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
        if (values.length < headers.length) return;
        const entry = headers.reduce((obj, h, i) => ({ ...obj, [h.trim()]: values[i] ? values[i].trim() : '' }), {});
        const { office, type, item, spec, unit, price, disposal_price, assumed_minutes, purchase_price } = entry;
        if (!office || !type) return;
        if (!data[office]) data[office] = {};
        if (!data[office][type]) data[office][type] = [];
        data[office][type].push({ item, spec, unit, price: parseFloat(price) || 0, disposal_price: parseFloat(disposal_price) || 0, assumed_minutes: parseInt(assumed_minutes) || 0, purchase_price: parseFloat(purchase_price) || 0 });
    });
    return data;
}

function parseMemberCSV(csvText) {
    const members = [];
    const lines = csvText.trim().split('\n');
    const headers = lines.shift().trim().split(',').map(h => h.trim());
    lines.forEach(line => {
        if (!line.trim()) return;
        const values = line.split(',');
        const entry = headers.reduce((obj, h, i) => ({ ...obj, [h]: values[i] ? values[i].trim() : '' }), {});
        if(entry.name && entry.office && entry.role) {
            entry.roles = entry.role.split(';').map(r => r.trim());
            members.push(entry);
        }
    });
    return members;
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' }).format(amount || 0);
}

// =================================================================================
// Excel出力関連の関数 (★★問題解決・完全版★★)
// =================================================================================

/**
 * 通貨形式の文字列 ("¥123,456") を数値 (123456) に変換します。
 * この関数はExcelに数値を正しく書き込むために必須です。
 * @param {string} text - 通貨形式の文字列
 * @returns {number} - 変換された数値
 */
function parseCurrency(text) {
    if (typeof text !== 'string') return 0;
    return parseFloat(text.replace(/[^0-9.-]+/g,"")) || 0;
}

/**
 * 指定されたシートに画像を追加するヘルパー関数です。
 * 画像のURL、位置、サイズ、画像がない場合の代替テキストを指定できます。
 * @param {object} options - 画像設定のオブジェクト
 * @param {ExcelJS.Workbook} options.workbook - ExcelJSのワークブックオブジェクト
 * @param {ExcelJS.Worksheet} options.sheet - 画像を追加するワークシートオブジェクト
 * @param {string | null} options.imageUrl - 画像のURL (例: 'hanko/logo.png')
 * @param {object} options.position - 画像の位置とサイズ (例: { tl: { col: 4, row: 3 }, ext: { width: 150, height: 50 } } または 'E4:H8')
 * @param {string} [options.fallbackText] - 画像が見つからない場合に表示するテキスト
 */
const addImageToSheet = async (options) => {
    const { workbook, sheet, imageUrl, position, fallbackText } = options;
    if (!imageUrl && !fallbackText) return;

    // 画像がある場合のみfetchを試みる
    if (imageUrl) {
        try {
            const response = await fetch(imageUrl);
            if (response.ok) {
                const imageBuffer = await response.arrayBuffer();
                // ファイル拡張子を取得
                const imageExt = imageUrl.split('.').pop().toLowerCase() || 'png';
                const imageId = workbook.addImage({ buffer: imageBuffer, extension: imageExt });
                sheet.addImage(imageId, position);
                return; // 画像を追加できたら処理を終了
            } else {
                console.warn(`画像が見つかりませんでした: ${imageUrl}`);
            }
        } catch (err) {
            console.error(`画像の処理中にエラーが発生しました: ${imageUrl}`, err);
        }
    }
    
    // 画像を追加できなかった場合、または最初から画像URLがない場合に代替テキストを書き込む
    if (fallbackText) {
        // positionが 'A1:B2' のようなレンジ形式の場合、左上のセルを特定
        const cellAddress = typeof position === 'string' ? position.split(':')[0] : sheet.getCell(position.tl.row + 1, position.tl.col + 1).address;
        const cell = sheet.getCell(cellAddress);
        cell.value = fallbackText;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
};

/**
 * フォームの入力内容をExcelファイルに出力します。
 */
async function exportToExcel() {
    const exportBtn = document.getElementById('export-excel-btn');
    exportBtn.textContent = 'Excel生成中...';
    exportBtn.disabled = true;

    const officeImageMap = {
        "札幌営業所": "sapporo.png",
        "仙台営業所": "sendai.png",
        "北関東営業所": "kitakanto.png",
        "首都圏営業所": "tokyo.png",
        "金沢営業所": "kanazawa.png",
        "名古屋営業所": "nagoya.png",
        "大阪営業所": "osaka.png",
        "岡山営業所": "okayama.png",
        "福岡営業所": "hukuoka.png", 
        "沖縄営業所": "okinawa.png"
    };

    try {
        // 1. テンプレートファイルを読み込む
        const templateUrl = 'template.xlsx';
        const response = await fetch(templateUrl);
        if (!response.ok) { throw new Error(`template.xlsx が見つかりません (HTTPステータス: ${response.status})`); }
        const templateData = await response.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(templateData);

        // 2. フォームから全てのデータを取得する
        const salesOfficeName = document.getElementById('sales-office').value; 
        const useTokyoAddressCheckbox = document.getElementById('use-tokyo-address');
        const isTokyoAddressForced = useTokyoAddressCheckbox.checked; 
        const companyName = document.getElementById('company-name').value;
        const contactName = document.getElementById('company-contact').value;
        const projectName = document.getElementById('project-name').value;
        const projectContent = document.getElementById('project-content').value;
        const myName = document.getElementById('my-name').value;
        const approverName = document.getElementById('approver-name').value;
        const grandTotal = parseCurrency(document.getElementById('grand-total').textContent);
        const totalWorkCost = parseCurrency(document.getElementById('total-work-cost').textContent);
        const totalMaterialCost = parseCurrency(document.getElementById('total-material-cost').textContent);
        const workDays = document.getElementById('work-days').value; 
        const workItems = Array.from(document.querySelectorAll('#work-items-table tbody tr')).map(row => ({ item: row.querySelector('.item-name').value, spec: row.querySelector('.spec').value, quantity: parseFloat(row.querySelector('.quantity').value) || 0, unit: row.querySelector('.unit').value, price: parseFloat(row.querySelector('.editable-price').value) || 0, total: parseCurrency(row.querySelector('.total').textContent) })).filter(d => d.quantity > 0);
        const materialItems = Array.from(document.querySelectorAll('#material-items-table tbody tr')).map(row => ({ name: row.querySelector('.name').value, spec: row.querySelector('.spec').value, quantity: parseFloat(row.querySelector('.quantity').value) || 0, unit: row.querySelector('.unit').value, price: parseFloat(row.querySelector('.unit-price').value) || 0, total: parseCurrency(row.querySelector('.total').textContent), listPrice: parseFloat(row.querySelector('.list-price').value) || 0 })).filter(d => d.name);
        const expenseItems = Array.from(document.querySelectorAll('#expense-items-table tbody tr')).filter(row => row.querySelector('input[type="checkbox"]').checked).map(row => ({ item: row.querySelector('.item-name').value, spec: row.querySelector('.spec').value, quantity: parseFloat(row.querySelector('.quantity').value) || 0, unit: row.querySelector('.unit').value, price: parseFloat(row.querySelector('.editable-price').value) || 0, total: parseCurrency(row.querySelector('.total').textContent) }));
        const commonCosts = [ 
            { item: '養生清掃費', total: parseCurrency(document.getElementById('cleaning-cost').textContent) }, 
            { item: '雑材消耗品費', total: parseCurrency(document.getElementById('supplies-cost').textContent) }, 
            { item: '運搬交通費', total: parseCurrency(document.getElementById('transport-cost').textContent) }, 
            { item: '諸経費', total: parseCurrency(document.getElementById('overhead-cost').textContent) }, 
            { item: '法定福利費', total: parseCurrency(document.getElementById('welfare-cost').textContent) } 
        ].filter(cost => cost.total > 0);

        // 3. 各シートを取得 (変更なし)
        const quoteSheet = workbook.getWorksheet("見積書");
        const workSheet = workbook.getWorksheet("内訳");
        const materialSheet = workbook.getWorksheet("内訳部材"); // ★部材シート名を確認 "部材内訳" かも？
        
        let companyImageFile;
        let fallbackText;

        if (isTokyoAddressForced) {
            // ★ チェックボックスがONの場合、強制的に東京の画像と営業所名を使う
            companyImageFile = "tokyo.png";
            fallbackText = "東京営業所";
        } else {
            // ★ チェックボックスがOFFの場合、通常通り選択された営業所の情報を使う
            companyImageFile = officeImageMap[salesOfficeName];
            fallbackText = salesOfficeName;
        }

        // 画像は 'hanko' フォルダにあると仮定します
        const companyImageUrl = companyImageFile ? `hanko/${companyImageFile}` : null; 

        // 4. 画像を挿入する (変更なし、ただしハンコの位置は調整しました)
        await Promise.all([
            addImageToSheet({ workbook, sheet: quoteSheet, imageUrl: 'hanko/logoname.png', position: 'P3:AH4' }),
            addImageToSheet({
                workbook,
                sheet: quoteSheet,
                imageUrl: companyImageUrl, // 動的なURLを渡す
                position: 'U5:AF12',
                fallbackText: salesOfficeName // 画像が見つからない場合は営業所名をテキストで表示
            }),
            addImageToSheet({ workbook, sheet: quoteSheet, imageUrl: 'hanko/kaku.png', position: 'Z6:AD10' }),
            addImageToSheet({ workbook, sheet: quoteSheet, imageUrl: 'hanko/qrcode.png', position: 'AE9:AH12' })
        ]);
        const hankoSize = { width: 50, height: 50 };
        
        // 担当者ハンコ (または名前テキスト)
        const myNameCell = quoteSheet.getCell('AD16');
        myNameCell.alignment = { vertical: 'middle', horizontal: 'center' };
        // 画像が見つからなかった場合に備えて、先にテキストとスタイルを設定
        if (myName) {
            myNameCell.value = myName;
            myNameCell.font = { color: { argb: 'FFFF0000' }, name: 'ＭＳ ゴシック', size: 12, bold: true };
        }
        await addImageToSheet({
            workbook, sheet: quoteSheet,
            imageUrl: myName ? `hanko/${encodeURIComponent(myName)}.png` : null,
            position: { tl: { col: 29.2, row: 15.3 }, ext: hankoSize },
            // fallbackTextは不要になったので削除
        });

        // 検印者ハンコ (または名前テキスト)
        const approverNameCell = quoteSheet.getCell('AA16');
        approverNameCell.alignment = { vertical: 'middle', horizontal: 'center' };
        // 画像が見つからなかった場合に備えて、先にテキストとスタイルを設定
        if (approverName) {
            approverNameCell.value = approverName;
            approverNameCell.font = { color: { argb: 'FFFF0000' }, name: 'ＭＳ ゴシック', size: 12, bold: true };
        }
        await addImageToSheet({
            workbook, sheet: quoteSheet,
            imageUrl: approverName ? `hanko/${encodeURIComponent(approverName)}.png` : null,
            position: { tl: { col: 26.2, row: 15.3 }, ext: hankoSize },
            // fallbackTextは不要になったので削除
        });

        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        // ★★★ ここから、セルへのデータ転記ロジックを全面的に修正します ★★★
        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

        // 5. セルにデータを転記する
        // (1) 見積書シートの共通項目 (結合セルの左上のセルを指定)
        quoteSheet.getCell("A3").value = companyName;      // A3:M3
        quoteSheet.getCell("A4").value = contactName;      // A4:K4
        quoteSheet.getCell("E14").value = projectName;     // E14:R14
        quoteSheet.getCell("E15").value = projectContent;  // E15:R15
        quoteSheet.getCell("F16").value = grandTotal;      // F16:O16

        // (2) 内訳シート (仕様通り)
        workItems.forEach((item, index) => {
            const row = workSheet.getRow(7 + index);
            row.getCell('B').value = item.item;
            row.getCell('C').value = item.spec;
            row.getCell('D').value = item.quantity;
            row.getCell('E').value = item.unit;
            row.getCell('F').value = item.price;
            row.getCell('G').value = item.total;
        });

        // (3) 部材内訳シート (仕様通り)
        if (materialSheet) {
            materialItems.forEach((item, index) => {
                const row = materialSheet.getRow(7 + index);
                row.getCell('B').value = item.name;
                row.getCell('C').value = item.spec;
                row.getCell('D').value = item.quantity;
                row.getCell('E').value = item.unit;
                row.getCell('F').value = item.price;
                row.getCell('G').value = item.total;
                row.getCell('H').value = item.listPrice;
            });
        }

        // (4) 見積書シートのサマリー項目 (仕様に合わせて書き方を完全に変更)
        const quoteSummaryItems = [];
        // ① 工事費は必ず追加
        quoteSummaryItems.push({ item: '照明工事費用一式', quantity: 1, unit: '式', total: totalWorkCost });
        // ② 部材があれば追加
        if (materialItems.length > 0) { 
            quoteSummaryItems.push({ item: '部材費用一式', quantity: 1, unit: '式', total: totalMaterialCost });
        }
        // ③ チェックされた経費を追加
        expenseItems.forEach(exp => { 
            quoteSummaryItems.push({ item: exp.item, spec: exp.spec, quantity: exp.quantity, unit: exp.unit, price: exp.price, total: exp.total });
        });
        // ④ その他の共通経費を追加
        commonCosts.forEach(cost => { 
            quoteSummaryItems.push({ item: cost.item, quantity: 1, unit: '式', total: cost.total });
        });

        // 取得したサマリーデータをExcelに転記
        let currentRowIndex = 20; // 開始行は20行目
        
        // 最初の行（工事費）だけ特別処理
        const firstItem = quoteSummaryItems.shift(); // 配列から最初の要素を取り出す
        if (firstItem) {
            const firstRow = quoteSheet.getRow(currentRowIndex);
            firstRow.getCell('C').value = firstItem.item;   // C20:K20
            firstRow.getCell('U').value = firstItem.quantity; // U20:W20
            firstRow.getCell('X').value = firstItem.unit;   // X20:Z20
            firstRow.getCell('AE').value = firstItem.total; // AE20:AI20
            currentRowIndex++;
        }
        
        // 21行目以降の残りの項目を転記
        quoteSummaryItems.forEach(item => {
            const row = quoteSheet.getRow(currentRowIndex);
            row.getCell('C').value = item.item;           // C:K
            row.getCell('L').value = item.spec || '';       // L:T
            row.getCell('U').value = item.quantity;       // U:W
            row.getCell('X').value = item.unit;           // X:Z
            if (item.price) {
                row.getCell('AA').value = item.price;     // AA:AD
            }
            row.getCell('AE').value = item.total;         // AE:AI
            currentRowIndex++;
        });

        // (5-2) 備考欄への転記
        // 工事日
        const selectedDays = Array.from(document.querySelectorAll('input[name="work-day"]:checked')).map(cb => cb.value);
        if (selectedDays.length > 0) {
            const workDaysText = `${selectedDays.join('/')}の工事で見積しております。`;
            quoteSheet.getCell('C43').value = workDaysText; // C43:AI43
        }
            // 工期
        if (workDays && parseFloat(workDays) > 0) {
            const workPeriodText = `工期は${workDays}日を想定しております。`;
            quoteSheet.getCell('C44').value = workPeriodText; // C44:AI44
        }

        // 5-EX. 各シートの印刷設定を行う
        const printOptions = {
            fitToPage: true, // ページに合わせる
            fitToWidth: 1,   // 横を1ページに収める
            fitToHeight: 1   // 縦を1ページに収める
        };
        
        if(quoteSheet) {
            quoteSheet.pageSetup = printOptions;
        }
        if(workSheet) {
            workSheet.pageSetup = printOptions;
        }
        if (materialSheet) {
            materialSheet.pageSetup = printOptions;
        }


        // 6. Excelファイルを生成してダウンロード (変更なし)
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        document.body.appendChild(a);
        a.href = url;
        a.download = `【見積書】${projectName}.xlsx`;
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        console.error(err);
        alert("Excelファイルの生成に失敗しました。\n" + err.message);
    } finally {
        exportBtn.textContent = 'Excel出力';
        exportBtn.disabled = false;
    }
}