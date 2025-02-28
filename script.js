/* script.js */

// Global Variables
let ROWS = 20;
let COLS = 10;
let sheetData = {}; // Format: { "A1": { formula: "=...", display: "...", format: { ... } } }
let selectedCell = null;
let isDragging = false;
let dragStartCell = null;
let dragEndCell = null;
let myChart = null; // For Chart.js visualization

document.addEventListener("DOMContentLoaded", () => {
  // Try to load saved data; if none, generate a blank grid.
  loadSheet();

  initToolbar();
  initFormulaPalette();

  // Attach listener for Excel upload if the file input exists.
  const excelUpload = document.getElementById("excelUpload");
  if (excelUpload) {
    excelUpload.addEventListener("change", uploadExcel);
  }
});

/* -------------------- Grid Generation -------------------- */
function generateGrid() {
  const sheetBody = document.getElementById("sheetBody");
  const columnHeader = document.getElementById("columnHeader");

  // Clear previous headers and body
  columnHeader.innerHTML = '<th class="bg-light"></th>';
  for (let c = 0; c < COLS; c++) {
    let th = document.createElement("th");
    th.classList.add("bg-light");
    th.innerText = String.fromCharCode(65 + c);
    initColumnResizing(th);
    columnHeader.appendChild(th);
  }

  sheetBody.innerHTML = "";
  for (let r = 1; r <= ROWS; r++) {
    let tr = document.createElement("tr");
    // Fixed row height to prevent auto-expanding
    tr.style.height = "30px"; 
    let thRow = document.createElement("th");
    thRow.classList.add("bg-light");
    thRow.innerText = r;
    tr.appendChild(thRow);

    for (let c = 0; c < COLS; c++) {
      let td = document.createElement("td");
      td.style.height = "30px";
      td.style.whiteSpace = "nowrap";
      td.style.overflow = "hidden";

      let cellId = String.fromCharCode(65 + c) + r;
      td.setAttribute("data-cell", cellId);
      td.contentEditable = true;
      td.classList.add("p-2");

      // Flag to clear cell on first key press
      td.dataset.clearOnType = "false";

      // Populate cell if data exists
      if (sheetData[cellId]) {
        td.innerText = sheetData[cellId].display || "";
        applySavedFormat(td, sheetData[cellId].format);
      }

      // Event listeners for cell interaction
      td.addEventListener("focus", (e) => selectCell(e.target));
      td.addEventListener("keydown", (e) => {
        if (td.dataset.clearOnType === "true" && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
          td.innerText = "";
          td.dataset.clearOnType = "false";
        }
      });
      td.addEventListener("input", (e) => updateCellData(e.target));
      td.addEventListener("blur", () => {
        recalcSheet();
        delete td.dataset.clearOnType;
      });

      // Append fill handle (small square at bottom-right)
      let fillHandle = document.createElement("div");
      fillHandle.className = "fill-handle";
      fillHandle.addEventListener("mousedown", (e) => {
        e.stopPropagation();
        startDragFill(e, td);
      });
      td.appendChild(fillHandle);

      tr.appendChild(td);
    }
    sheetBody.appendChild(tr);
  }
}

/* -------------------- Column Resizing -------------------- */
function initColumnResizing(thElement) {
  let startX, startWidth;
  const resizeHandle = document.createElement("span");
  resizeHandle.className = "col-resize";
  thElement.style.position = "relative";
  thElement.appendChild(resizeHandle);
  resizeHandle.addEventListener("mousedown", function (e) {
    startX = e.pageX;
    startWidth = thElement.offsetWidth;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
    e.preventDefault();
  });
  function onMouseMove(e) {
    let newWidth = startWidth + (e.pageX - startX);
    if (newWidth > 30) {
      thElement.style.width = newWidth + "px";
    }
  }
  function onMouseUp() {
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  }
}

/* -------------------- Toolbar Initialization -------------------- */
function initToolbar() {
  // Formatting buttons
  document.getElementById("boldBtn").addEventListener("click", () => {
    if (selectedCell) toggleFormat(selectedCell, "bold");
  });
  document.getElementById("italicBtn").addEventListener("click", () => {
    if (selectedCell) toggleFormat(selectedCell, "italic");
  });
  document.getElementById("underlineBtn").addEventListener("click", () => {
    if (selectedCell) toggleFormat(selectedCell, "underline");
  });
  // Font size
  document.getElementById("fontSizeSelect").addEventListener("change", (e) => {
    if (selectedCell) {
      selectedCell.style.fontSize = e.target.value;
      saveCellFormat(selectedCell);
    }
  });
  // Font color
  document.getElementById("fontColorPicker").addEventListener("change", (e) => {
    if (selectedCell) {
      selectedCell.style.color = e.target.value;
      saveCellFormat(selectedCell);
    }
  });
  // Formula bar functionality
  const formulaBar = document.getElementById("formulaBar");
  formulaBar.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && selectedCell) {
      e.preventDefault();
      let formula = formulaBar.value;
      setCellFormula(selectedCell, formula);
      recalcSheet();
      selectedCell.focus();
    }
  });
}

/* -------------------- Formula Palette Initialization -------------------- */
function initFormulaPalette() {
  document.getElementById("formulaPaletteBtn").addEventListener("click", () => {
    let formulaModal = new bootstrap.Modal(document.getElementById("formulaModal"));
    formulaModal.show();
  });
  document.querySelectorAll(".formula-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      let formulaTemplate = e.target.dataset.formula;
      document.getElementById("formulaBar").value = formulaTemplate;
      let modalEl = document.getElementById("formulaModal");
      let modal = bootstrap.Modal.getInstance(modalEl);
      modal.hide();
    });
  });
}

/* -------------------- Cell Selection & Data Update -------------------- */
function selectCell(cell) {
  if (selectedCell) selectedCell.classList.remove("selected");
  selectedCell = cell;
  cell.classList.add("selected");
  cell.dataset.clearOnType = "true";
  let cellId = cell.getAttribute("data-cell");
  document.getElementById("formulaBar").value =
    (sheetData[cellId] && sheetData[cellId].formula) || cell.innerText;
}

function updateCellData(cell) {
  let cellId = cell.getAttribute("data-cell");
  let content = cell.innerText;
  if (!sheetData[cellId]) sheetData[cellId] = {};
  if (!content.startsWith("=") && !isNaN(content) && content.trim() !== "") {
    sheetData[cellId].display = parseFloat(content);
    sheetData[cellId].formula = null;
  } else {
    sheetData[cellId].formula = content.startsWith("=") ? content : null;
    sheetData[cellId].display = content;
  }
}

/* -------------------- Formula Handling -------------------- */
function setCellFormula(cell, formula) {
  let cellId = cell.getAttribute("data-cell");
  if (!sheetData[cellId]) sheetData[cellId] = {};
  sheetData[cellId].formula = formula;
  
  if (formula.startsWith("=")) {
    let result = evaluateFormula(formula.slice(1));
    if (result === "ERR") {
      Swal.fire("Error", "Invalid formula entered.", "error");
      return;
    }
    sheetData[cellId].display = result;
    cell.innerText = result;
  } else {
    sheetData[cellId].display = formula;
    cell.innerText = formula;
  }
  reapplyFillHandle(cell);
}

function recalcSheet() {
  for (let cellId in sheetData) {
    if (sheetData[cellId].formula && sheetData[cellId].formula.startsWith("=")) {
      let cellElement = document.querySelector(`[data-cell="${cellId}"]`);
      let result = evaluateFormula(sheetData[cellId].formula.slice(1));
      sheetData[cellId].display = result;
      if (cellElement) {
        cellElement.innerText = result;
        reapplyFillHandle(cellElement);
      }
    }
  }
  // Update Chart visualization (if applicable)
  updateChart();
}

function evaluateFormula(formula) {
  try {
    let match = formula.match(/([A-Z_]+)\(([^)]*)\)/i);
    if (!match) {
      // Supports references with optional "$" signs (absolute references)
      let expr = formula.replace(/(\$?[A-Z]+\$?[0-9]+)/g, (ref) => {
        let normalized = ref.replace(/\$/g, "");
        if (sheetData[normalized]) return parseFloat(sheetData[normalized].display) || 0;
        return 0;
      });
      return eval(expr);
    }
    let func = match[1].toUpperCase();
    let arg = match[2];
    let cells = parseRange(arg);

    if (["SUM", "AVERAGE", "MAX", "MIN", "COUNT", "MEDIAN"].includes(func)) {
      if (func === "COUNT") {
        let count = 0;
        cells.forEach(cellId => {
          let value = sheetData[cellId] ? sheetData[cellId].display : "";
          if (value !== "" && !isNaN(value)) count++;
        });
        return count;
      } else {
        let values = cells.map(cellId => {
          let val = sheetData[cellId] ? parseFloat(sheetData[cellId].display) : NaN;
          return !isNaN(val) ? val : 0;
        });
        if (func === "SUM") return values.reduce((a, b) => a + b, 0);
        if (func === "AVERAGE") return values.reduce((a, b) => a + b, 0) / values.length;
        if (func === "MAX") return Math.max(...values);
        if (func === "MIN") return Math.min(...values);
        if (func === "MEDIAN") {
          values.sort((a, b) => a - b);
          let mid = Math.floor(values.length / 2);
          return values.length % 2 === 0 ? (values[mid - 1] + values[mid]) / 2 : values[mid];
        }
      }
    }
    else if (["TRIM", "UPPER", "LOWER"].includes(func)) {
      let cellRef = arg.trim().replace(/\$/g, "");
      let text = sheetData[cellRef]?.display || "";
      if (func === "TRIM") return text.trim();
      if (func === "UPPER") return text.toUpperCase();
      if (func === "LOWER") return text.toLowerCase();
    }
    else if (func === "CLEAR") {
      return "";
    }
    return "ERR";
  } catch (err) {
    return "ERR";
  }
}

function parseRange(rangeStr) {
  rangeStr = rangeStr.trim();
  if (!rangeStr.includes(":")) return [rangeStr];
  let parts = rangeStr.split(":");
  let start = parts[0].trim();
  let end = parts[1].trim();
  let startCol = start.charCodeAt(0);
  let startRow = parseInt(start.substring(1));
  let endCol = end.charCodeAt(0);
  let endRow = parseInt(end.substring(1));
  let cells = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      cells.push(String.fromCharCode(c) + r);
    }
  }
  return cells;
}

/* -------------------- Drag-Fill Functionality -------------------- */
function startDragFill(e, cell) {
  isDragging = true;
  dragStartCell = cell;
  document.addEventListener("mousemove", dragFill);
  document.addEventListener("mouseup", stopDragFill);
}

function dragFill(e) {
  if (!isDragging || !dragStartCell) return;
  let element = document.elementFromPoint(e.clientX, e.clientY);
  if (element && element.tagName === "TD") {
    if (dragEndCell) dragEndCell.classList.remove("selected");
    dragEndCell = element;
    dragEndCell.classList.add("selected");
  }
}

function stopDragFill() {
  if (isDragging && dragStartCell && dragEndCell) {
    let startCellId = dragStartCell.getAttribute("data-cell");
    let endCellId = dragEndCell.getAttribute("data-cell");
    let startCol = startCellId.charCodeAt(0);
    let startRow = parseInt(startCellId.substring(1));
    let endCol = endCellId.charCodeAt(0);
    let endRow = parseInt(endCellId.substring(1));
    let minCol = Math.min(startCol, endCol);
    let maxCol = Math.max(startCol, endCol);
    let minRow = Math.min(startRow, endRow);
    let maxRow = Math.max(startRow, endRow);
    let fillValue = sheetData[startCellId]?.display || "";
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        let cellId = String.fromCharCode(c) + r;
        let cellElement = document.querySelector(`[data-cell="${cellId}"]`);
        if (cellElement) {
          cellElement.innerText = fillValue;
          if (!sheetData[cellId]) sheetData[cellId] = {};
          sheetData[cellId].display = fillValue;
          reapplyFillHandle(cellElement);
        }
      }
    }
  }
  isDragging = false;
  dragStartCell = null;
  if (dragEndCell) {
    dragEndCell.classList.remove("selected");
    dragEndCell = null;
  }
  document.removeEventListener("mousemove", dragFill);
  document.removeEventListener("mouseup", stopDragFill);
}

/* -------------------- Formatting Functions -------------------- */
function toggleFormat(cell, formatType) {
  if (formatType === "bold") {
    cell.style.fontWeight = cell.style.fontWeight === "bold" ? "normal" : "bold";
  } else if (formatType === "italic") {
    cell.style.fontStyle = cell.style.fontStyle === "italic" ? "normal" : "italic";
  } else if (formatType === "underline") {
    cell.style.textDecoration = cell.style.textDecoration === "underline" ? "none" : "underline";
  }
  saveCellFormat(cell);
}

function saveCellFormat(cell) {
  let cellId = cell.getAttribute("data-cell");
  if (!sheetData[cellId]) sheetData[cellId] = {};
  sheetData[cellId].format = {
    fontWeight: cell.style.fontWeight,
    fontStyle: cell.style.fontStyle,
    textDecoration: cell.style.textDecoration,
    fontSize: cell.style.fontSize,
    color: cell.style.color,
  };
}

function applySavedFormat(cell, format) {
  if (format) {
    cell.style.fontWeight = format.fontWeight || "normal";
    cell.style.fontStyle = format.fontStyle || "normal";
    cell.style.textDecoration = format.textDecoration || "none";
    cell.style.fontSize = format.fontSize || "12px";
    cell.style.color = format.color || "#000";
  }
}

function reapplyFillHandle(cell) {
  let handle = cell.querySelector(".fill-handle");
  if (handle) cell.removeChild(handle);
  let newHandle = document.createElement("div");
  newHandle.className = "fill-handle";
  newHandle.addEventListener("mousedown", (e) => {
    e.stopPropagation();
    startDragFill(e, cell);
  });
  cell.appendChild(newHandle);
}

/* -------------------- Add/Delete Rows & Columns -------------------- */
function addRow() {
  ROWS++;
  generateGrid();
  recalcSheet();
}

function addColumn() {
  COLS++;
  generateGrid();
  recalcSheet();
}

function confirmDeleteRow() {
  if (!selectedCell) {
    Swal.fire("Error", "Select a cell in the row to delete.", "error");
    return;
  }
  Swal.fire({
    title: "Are you sure?",
    text: "This will delete the entire row.",
    icon: "warning",
    showCancelButton: true,
    confirmButtonText: "Yes, delete it!"
  }).then((result) => {
    if (result.isConfirmed) deleteRow();
  });
}

function deleteRow() {
  let cellId = selectedCell.getAttribute("data-cell");
  let rowToDelete = parseInt(cellId.substring(1));
  let newSheetData = {};
  for (let key in sheetData) {
    let rowNum = parseInt(key.substring(1));
    if (rowNum === rowToDelete) continue;
    else if (rowNum > rowToDelete) {
      let newKey = key.charAt(0) + (rowNum - 1);
      newSheetData[newKey] = sheetData[key];
    } else {
      newSheetData[key] = sheetData[key];
    }
  }
  sheetData = newSheetData;
  ROWS = Math.max(ROWS - 1, 1);
  generateGrid();
  recalcSheet();
  Swal.fire("Deleted", "Row deleted successfully.", "success");
}

function confirmDeleteColumn() {
  if (!selectedCell) {
    Swal.fire("Error", "Select a cell in the column to delete.", "error");
    return;
  }
  Swal.fire({
    title: "Are you sure?",
    text: "This will delete the entire column.",
    icon: "warning",
    showCancelButton: true,
    confirmButtonText: "Yes, delete it!"
  }).then((result) => {
    if (result.isConfirmed) deleteColumn();
  });
}

function deleteColumn() {
  let cellId = selectedCell.getAttribute("data-cell");
  let colToDelete = cellId.charAt(0);
  let colIndex = colToDelete.charCodeAt(0);
  let newSheetData = {};
  for (let key in sheetData) {
    let currentCol = key.charAt(0);
    let currentRow = key.substring(1);
    let currentColIndex = currentCol.charCodeAt(0);
    if (currentColIndex === colIndex) continue;
    else if (currentColIndex > colIndex) {
      let newCol = String.fromCharCode(currentColIndex - 1);
      let newKey = newCol + currentRow;
      newSheetData[newKey] = sheetData[key];
    } else {
      newSheetData[key] = sheetData[key];
    }
  }
  sheetData = newSheetData;
  COLS = Math.max(COLS - 1, 1);
  generateGrid();
  recalcSheet();
  Swal.fire("Deleted", "Column deleted successfully.", "success");
}

/* -------------------- Remove Duplicates -------------------- */
function removeDuplicates() {
  let newSheetData = {};
  let seenRows = new Set();
  let newRowCount = 0;
  for (let r = 1; r <= ROWS; r++) {
    let rowString = "";
    for (let c = 0; c < COLS; c++) {
      let cellId = String.fromCharCode(65 + c) + r;
      let cellVal = sheetData[cellId] ? sheetData[cellId].display || "" : "";
      rowString += cellVal + "|";
    }
    if (!seenRows.has(rowString)) {
      seenRows.add(rowString);
      newRowCount++;
      for (let c = 0; c < COLS; c++) {
        let oldCellId = String.fromCharCode(65 + c) + r;
        let newCellId = String.fromCharCode(65 + c) + newRowCount;
        newSheetData[newCellId] = sheetData[oldCellId] || {};
      }
    }
  }
  sheetData = newSheetData;
  ROWS = newRowCount;
  generateGrid();
  recalcSheet();
  Swal.fire("Completed", "Duplicate rows removed.", "success");
}

/* -------------------- Find and Replace -------------------- */
function findAndReplace() {
  let findText = document.getElementById("findText").value.trim();
  let replaceText = document.getElementById("replaceText").value;
  if (!findText) {
    Swal.fire("Info", "Please enter text to find.", "info");
    return;
  }
  let replaced = false;
  for (let cellId in sheetData) {
    if (
      sheetData[cellId].display &&
      sheetData[cellId].display.toString().includes(findText)
    ) {
      sheetData[cellId].display = sheetData[cellId]
        .display.toString()
        .replace(new RegExp(findText, "g"), replaceText);
      let cellElement = document.querySelector(`[data-cell="${cellId}"]`);
      if (cellElement) {
        cellElement.innerText = sheetData[cellId].display;
        reapplyFillHandle(cellElement);
      }
      replaced = true;
    }
  }
  if (replaced)
    Swal.fire("Completed", "Find and Replace completed.", "success");
  else
    Swal.fire("Info", "No matching text found.", "info");
}

/* -------------------- Save & Load -------------------- */
function saveSheet() {
  try {
    localStorage.setItem("sheetData", JSON.stringify({ ROWS, COLS, sheetData }));
    Swal.fire("Saved", "Sheet saved successfully!", "success");
  } catch (err) {
    Swal.fire("Error", "Failed to save sheet data.", "error");
  }
}

function loadSheet() {
  try {
    let data = localStorage.getItem("sheetData");
    if (data) {
      let parsed = JSON.parse(data);
      ROWS = parsed.ROWS;
      COLS = parsed.COLS;
      sheetData = parsed.sheetData;
      generateGrid();
      recalcSheet();
      Swal.fire("Loaded", "Sheet loaded successfully!", "success");
    } else {
      // No saved data found – generate a new blank grid without an alert.
      generateGrid();
    }
  } catch (err) {
    console.error("Error loading sheet data:", err);
    generateGrid();
  }
}

/* -------------------- Excel Upload -------------------- */
// Uses SheetJS (XLSX) to load an Excel file and convert it to sheetData.
function uploadExcel(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = e.target.result;
    const workbook = XLSX.read(data, { type: 'binary' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    // Convert worksheet to an array of arrays (each representing a row)
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Update ROWS and COLS based on uploaded data.
    ROWS = jsonData.length;
    COLS = Math.max(...jsonData.map(row => row.length));
    sheetData = {};
    for (let r = 0; r < jsonData.length; r++) {
      for (let c = 0; c < jsonData[r].length; c++) {
        const cellId = String.fromCharCode(65 + c) + (r + 1);
        sheetData[cellId] = { display: jsonData[r][c], formula: null, format: {} };
      }
    }
    generateGrid();
    recalcSheet();
    Swal.fire("Loaded", "Excel data loaded successfully!", "success");
  };
  reader.readAsBinaryString(file);
}

/* -------------------- Additional Features & Bonus -------------------- */
// [BONUS] Data Visualization using Chart.js
function updateChart() {
  // Example: Visualize data in column A (cells A1 through A{ROWS})
  let labels = [];
  let data = [];
  for (let r = 1; r <= ROWS; r++) {
    let cellId = "A" + r;
    labels.push(cellId);
    let value = sheetData[cellId] ? parseFloat(sheetData[cellId].display) : 0;
    data.push(value);
  }
  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (myChart) myChart.destroy();
  myChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Column A Data',
        data: data,
        borderColor: 'rgba(75, 192, 192, 1)',
        backgroundColor: 'rgba(75, 192, 192, 0.2)',
      }]
    },
    options: {
      responsive: true,
      scales: {
        y: { beginAtZero: true }
      }
    }
  });
}

// [BONUS] Additional mathematical or data quality functions can be added here.


document.getElementById('darkModeToggle').addEventListener('click', () => {
  document.body.classList.toggle('dark-mode');
  const btn = document.getElementById('darkModeToggle');
  if (document.body.classList.contains('dark-mode')) {
    btn.textContent = 'Light Mode';
  } else {
    btn.textContent = 'Dark Mode';
  }
});