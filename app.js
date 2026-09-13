const sampleRows = [
  ["4/03", "株式会社 東光", "クラウドサービス", "佐藤", "¥1,240,000", "+18.2%"],
  ["4/05", "みなと商事", "コンサルティング", "鈴木", "¥980,000", "+5.4%"],
  ["4/08", "青山デザイン", "ソフトウェア", "高橋", "¥1,560,000", "+22.1%"],
  ["4/12", "北辰工業", "クラウドサービス", "伊藤", "¥2,120,000", "+9.8%"],
  ["4/15", "LIFE株式会社", "保守・サポート", "渡辺", "¥760,000", "−2.3%"],
  ["4/19", "三葉ホールディングス", "ソフトウェア", "山本", "¥1,890,000", "+15.6%"],
  ["4/23", "成和パートナーズ", "コンサルティング", "中村", "¥1,340,000", "+7.9%"],
  ["4/27", "株式会社ノース", "クラウドサービス", "小林", "¥1,670,000", "+11.2%"]
];

const $ = (selector) => document.querySelector(selector);
const tableBody = $("#sheetTable tbody");
let zoom = 75;

function renderRows(rows) {
  tableBody.innerHTML = rows.slice(0, 12).map((row) => `<tr>${row.map((cell, index) => `<td class="${index > 3 ? "numeric" : ""}">${escapeHtml(cell ?? "")}</td>`).join("")}</tr>`).join("");
}

function escapeHtml(value) {
  const div = document.createElement("div");
  div.textContent = String(value);
  return div.innerHTML;
}

function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2400);
}

function updateLayout() {
  const orientation = document.querySelector('input[name="orientation"]:checked').value;
  const paper = $("#paperSize").selectedOptions[0].text;
  $("#page").classList.toggle("landscape", orientation === "landscape");
  $("#statusSub").textContent = `${paper}・${orientation === "portrait" ? "縦向き" : "横向き"}・1ページ`;
}

async function loadFile(file) {
  if (!file) return;
  if (file.size > 20 * 1024 * 1024) return showToast("20MB以下のファイルを選択してください");
  $("#fileLabel").textContent = file.name;
  $("#statusText").textContent = "読み込み中…";
  try {
    if (!window.XLSX) throw new Error("reader unavailable");
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rows.length > 1) {
      const headings = rows[0].slice(0, 6);
      $("#sheetTable thead tr").innerHTML = headings.map((h, i) => `<th class="${i > 3 ? "numeric" : ""}">${escapeHtml(h)}</th>`).join("");
      renderRows(rows.slice(1));
    }
    $("#sheetName").textContent = workbook.SheetNames[0];
    $("#statusText").textContent = "変換準備完了";
    showToast("Excelファイルを読み込みました");
  } catch (error) {
    $("#statusText").textContent = "読み込みエラー";
    showToast("ファイルを読み込めませんでした");
  }
}

renderRows(sampleRows);
$("#fileInput").addEventListener("change", (event) => loadFile(event.target.files[0]));
const dropzone = $("#dropzone");
["dragenter", "dragover"].forEach((name) => dropzone.addEventListener(name, (event) => { event.preventDefault(); dropzone.classList.add("dragging"); }));
["dragleave", "drop"].forEach((name) => dropzone.addEventListener(name, (event) => { event.preventDefault(); dropzone.classList.remove("dragging"); }));
dropzone.addEventListener("drop", (event) => loadFile(event.dataTransfer.files[0]));
dropzone.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") $("#fileInput").click(); });
document.querySelectorAll('input[name="orientation"], #paperSize, #scaleMode').forEach((control) => control.addEventListener("change", updateLayout));
$("#marginRange").addEventListener("input", (event) => { $("#marginOutput").textContent = `${event.target.value == 15 ? "標準（" : ""}${event.target.value}mm${event.target.value == 15 ? "）" : ""}`; $("#page").style.padding = `${Math.max(20, event.target.value * 2.2)}px`; });
$("#gridToggle").addEventListener("change", (event) => document.body.classList.toggle("hide-grid", !event.target.checked));
$("#headerToggle").addEventListener("change", (event) => $("#sheetTable thead").style.display = event.target.checked ? "table-header-group" : "none");
function setZoom(next) { zoom = Math.max(50, Math.min(100, next)); $("#zoomValue").textContent = `${zoom}%`; $("#page").style.transform = `scale(${zoom / 75})`; }
$("#zoomOut").addEventListener("click", () => setZoom(zoom - 5));
$("#zoomIn").addEventListener("click", () => setZoom(zoom + 5));
$("#exportButton").addEventListener("click", () => { $("#statusText").textContent = "PDF出力画面を開きます"; showToast("印刷画面で「PDFに保存」を選択してください"); setTimeout(() => window.print(), 450); });
updateLayout();
