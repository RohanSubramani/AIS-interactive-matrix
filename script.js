const table = document.getElementById("matrix-table");
const output = document.getElementById("output");
const toggleIntroBtn = document.getElementById("toggle-intro-btn");
const introductionBox = document.getElementById("introduction-box");

// Function to Read Excel File
async function loadExcel(filePath) {
  try {
    const response = await fetch(filePath);
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  } catch (error) {
    console.error("Error loading Excel file:", error);
    output.innerHTML = `<p style="color: red;">Failed to load data. Please try again later.</p>`;
    return [];
  }
}

// Function to Remove Empty Rows and Columns
function cleanData(data) {
  return data.filter(row => row.some(cell => cell.trim() !== ""));
}

// Function to Build Table
function buildTable(headers, rows) {
  let tableHTML = "<tr><th></th>";
  headers.slice(1).forEach((header, index) => {
    tableHTML += `<th data-col="${index}">${header}</th>`;
  });
  tableHTML += "</tr>";

  rows.forEach((row, rowIndex) => {
    tableHTML += `<tr><th data-row="${rowIndex}">${row[0]}</th>`;
    row.slice(1).forEach(cell => {
      tableHTML += `<td>${cell || ""}</td>`;
    });
    tableHTML += "</tr>";
  });

  table.innerHTML = tableHTML;
}

// Function to Display Row Details
function displayRow(headers, row) {
  const title = row[0];
  let displayHTML = `<h2>${title}</h2>`;
  headers.slice(1).forEach((header, i) => {
    if (row[i + 1].trim()) {
      displayHTML += `<p><strong>${header}:</strong> ${row[i + 1]}</p>`;
    }
  });
  output.innerHTML = displayHTML;
  scrollToOutput();
}

// Function to Display Column Details
function displayColumn(headers, rows, colIndex) {
  const title = headers[colIndex + 1];
  let displayHTML = `<h2>${title}</h2>`;
  rows.forEach(row => {
    if (row[colIndex + 1].trim()) {
      displayHTML += `<p><strong>${row[0]}:</strong> ${row[colIndex + 1]}</p>`;
    }
  });
  output.innerHTML = displayHTML;
  scrollToOutput();
}

// Function to Scroll to the Output Section
function scrollToOutput() {
  document.getElementById("display-section").scrollIntoView({ behavior: "smooth" });
}

// Add Event Listeners to Table Headers
function addClickListeners(headers, rows) {
  table.addEventListener("click", (e) => {
    const rowIndex = e.target.getAttribute("data-row");
    const colIndex = e.target.getAttribute("data-col");

    if (rowIndex !== null) displayRow(headers, rows[Number(rowIndex)]);
    if (colIndex !== null) displayColumn(headers, rows, Number(colIndex));
  });
}

// Toggle Introduction Visibility
toggleIntroBtn.addEventListener("click", () => {
  if (introductionBox.classList.contains("hidden")) {
    introductionBox.classList.remove("hidden");
    toggleIntroBtn.textContent = "Click to hide introduction (recommended)";
  } else {
    introductionBox.classList.add("hidden");
    toggleIntroBtn.textContent = "Click to expand introduction";
  }
});

// Initialize the Table
async function init() {
  const filePath = "matrix.xlsx"; // Replace with your file path or use "matrix.json" if converted
  let data = await loadExcel(filePath);
  if (data.length === 0) return; // Stop if data failed to load
  data = cleanData(data); // Remove empty rows
  const headers = data[0];
  const rows = data.slice(1);

  buildTable(headers, rows);
  addClickListeners(headers, rows);
}

init();
