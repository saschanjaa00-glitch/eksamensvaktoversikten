let currentTeachersData = [];
let filterDate = null;
let currentWorkbook = null;
let currentSheetName = '';
let currentFileName = 'edited_schedule.xlsx';

const tabScheduleBtn = document.getElementById('tabScheduleBtn');
const tabEditorBtn = document.getElementById('tabEditorBtn');
const scheduleTab = document.getElementById('scheduleTab');
const editorTab = document.getElementById('editorTab');

tabScheduleBtn.addEventListener('click', () => setActiveTab('schedule'));
tabEditorBtn.addEventListener('click', () => setActiveTab('editor'));
document.getElementById('downloadEditedBtn').addEventListener('click', downloadEditedWorkbook);

document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();

  const fileInput = document.getElementById('file');
  const sheetNameInput = document.getElementById('sheetName');
  const filterDateInput = document.getElementById('filterDate');
  const file = fileInput.files[0];
  const sheetName = sheetNameInput.value || 'Lærervakter';

  if (!file) {
    showError('Please select a file');
    return;
  }

  // Set filter date if provided
  filterDate = filterDateInput.value ? new Date(filterDateInput.value) : null;
  currentFileName = file.name;

  showLoading(true);
  hideError();
  hideResults();

  try {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('sheetName', sheetName);

    const response = await fetch('/api/upload', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || 'Failed to process file');
    }

    const data = await response.json();
    displayResults(data);
    await loadEditableSheet(file, sheetName);
    setActiveTab('schedule');
  } catch (error) {
    showError(error.message);
  } finally {
    showLoading(false);
  }
});

function showLoading(show) {
  const loading = document.getElementById('loading');
  loading.style.display = show ? 'block' : 'none';
}

function showError(message) {
  const errorDiv = document.getElementById('error');
  errorDiv.textContent = message;
  errorDiv.style.display = 'block';
}

function hideError() {
  document.getElementById('error').style.display = 'none';
}

function hideResults() {
  document.getElementById('results').style.display = 'none';
}

function setActiveTab(tabName) {
  const isSchedule = tabName === 'schedule';
  tabScheduleBtn.classList.toggle('active', isSchedule);
  tabEditorBtn.classList.toggle('active', !isSchedule);
  scheduleTab.style.display = isSchedule ? 'block' : 'none';
  editorTab.style.display = isSchedule ? 'none' : 'block';
}

function displayResults(data) {
  const resultsDiv = document.getElementById('results');
  const resultsContent = document.getElementById('resultsContent');

  let html = '';

  if (data.teachers.length === 0) {
    html = '<p>No teacher data found in the sheet.</p>';
  } else {
    data.teachers.forEach(teacher => {
      let filteredDates = teacher.dates;
      
      // Apply date filter if set
      if (filterDate) {
        filteredDates = teacher.dates.filter(date => isDateOnOrAfter(date, filterDate));
      }

      html += `
        <div class="teacher-card">
          <div class="teacher-name">👨‍🏫 ${escapeHtml(teacher.name)}</div>
          <div class="teacher-dates">
            ${filteredDates.length > 0
              ? filteredDates
                  .map(date => {
                    const dateStr = escapeHtml(date.toString());
                    const formatted = dateStr.replace(/\(([^)]+)\)/, '<span class="date-time">($1)</span>');
                    return `<span class="date-badge">${formatted}</span>`;
                  })
                  .join('')
              : '<span class="date-badge empty">No scheduled dates</span>'
            }
          </div>
        </div>
      `;
    });
  }

  resultsContent.innerHTML = html;
  currentTeachersData = data.teachers;
  resultsDiv.style.display = 'block';

  // Scroll to results
  resultsDiv.scrollIntoView({ behavior: 'smooth' });
}

function isDateOnOrAfter(dateString, filterDate) {
  // Extract DD.MM from the date string (e.g., "04.05" from "04.05 (08.30-11.45)")
  const match = dateString.toString().match(/(\d{2})\.(\d{2})/);
  if (!match) return true;
  
  const day = parseInt(match[1]);
  const month = parseInt(match[2]);
  
  // Get filter date components
  const filterDay = filterDate.getDate();
  const filterMonth = filterDate.getMonth() + 1; // getMonth() returns 0-11
  
  // Compare month first, then day
  if (month !== filterMonth) {
    return month > filterMonth;
  }
  return day >= filterDay;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

async function loadEditableSheet(file, preferredSheetName) {
  if (!window.XLSX) {
    throw new Error('Excel editor library failed to load. Please refresh and try again.');
  }

  const buffer = await file.arrayBuffer();
  currentWorkbook = window.XLSX.read(buffer, { type: 'array' });

  currentSheetName = currentWorkbook.SheetNames.includes(preferredSheetName)
    ? preferredSheetName
    : currentWorkbook.SheetNames[0];

  const worksheet = currentWorkbook.Sheets[currentSheetName];
  const sheetData = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  renderSheetEditor(sheetData);
}

function renderSheetEditor(sheetData) {
  const editor = document.getElementById('sheetEditorContainer');

  if (!sheetData || sheetData.length === 0) {
    editor.innerHTML = '<p style="padding:12px; color:#666;">No editable data found in this sheet.</p>';
    return;
  }

  const maxCols = Math.max(...sheetData.map((row) => row.length), 1);
  const normalized = sheetData.map((row) => {
    const cells = [...row];
    while (cells.length < maxCols) cells.push('');
    return cells;
  });

  const DATE_START_COL = 4;
  
  // Find Maks column (look for header containing "Maks")
  let maksCol = -1;
  if (normalized.length > 0) {
    maksCol = normalized[0].findIndex(cell => String(cell).toLowerCase().includes('maks'));
  }

  const tableRows = normalized
    .map((row, rowIndex) => {
      const tds = row
        .map((cell, colIndex) => {
          const cellValue = String(cell ?? '');
          const isDateColumn = colIndex >= DATE_START_COL && colIndex !== maksCol;
          const isMaksColumn = colIndex === maksCol;

          if (rowIndex > 0 && isDateColumn) {
            const checked = cellValue.trim() === '1' ? 'checked' : '';
            return `<td class="toggle-cell"><input type="checkbox" data-type="toggle" data-row="${rowIndex}" data-col="${colIndex}" ${checked} /></td>`;
          }

          const colClass = isMaksColumn ? 'maks-col' : '';
          return `<td class="${colClass}"><input data-type="text" data-row="${rowIndex}" data-col="${colIndex}" value="${escapeHtml(cellValue)}" /></td>`;
        })
        .join('');
      
      // Add Sum and % columns
      let sumCell = '';
      let percentCell = '';
      
      if (rowIndex === 0) {
        // Header row
        sumCell = '<td class="calc-col sum-col"><input readonly value="Sum" /></td>';
        percentCell = '<td class="calc-col percent-col"><input readonly value="%" /></td>';
      } else {
        // Data rows - calculate sum and percentage (exclude Maks column from count)
        const sum = row.filter((cell, colIndex) => {
          return colIndex >= DATE_START_COL && colIndex !== maksCol && String(cell).trim() === '1';
        }).length;
        const maks = maksCol >= 0 ? parseFloat(row[maksCol]) || 0 : 0;
        const percent = maks > 0 ? Math.round((sum / maks) * 100) : 0;
        
        sumCell = `<td class="calc-col sum-col"><input readonly value="${sum}" data-calc-type="sum" data-row="${rowIndex}" /></td>`;
        percentCell = `<td class="calc-col percent-col"><input readonly value="${percent}%" data-calc-type="percent" data-row="${rowIndex}" /></td>`;
      }
      
      return `<tr>${tds}${sumCell}${percentCell}</tr>`;
    })
    .join('');

  editor.innerHTML = `<table class="sheet-editor-table"><tbody>${tableRows}</tbody></table>`;

  // Store maksCol for use in the update function
  const storedMaksCol = maksCol;

  // Update calculations when checkboxes change
  editor.querySelectorAll('input[data-type="toggle"]').forEach((checkbox) => {
    checkbox.addEventListener('change', () => {
      updateRowCalculations();
    });
  });
  
  // Also update when Maks value changes
  if (storedMaksCol >= 0) {
    editor.querySelectorAll(`input[data-col="${storedMaksCol}"]`).forEach((input) => {
      input.addEventListener('input', () => {
        updateRowCalculations();
      });
    });
  }
  
  function updateRowCalculations() {
    const table = editor.querySelector('.sheet-editor-table');
    const rows = table.querySelectorAll('tr');
    
    rows.forEach((tr, rowIndex) => {
      if (rowIndex === 0) return; // Skip header
      
      const checkboxes = tr.querySelectorAll('input[data-type="toggle"]');
      const sum = Array.from(checkboxes).filter(cb => cb.checked).length;
      
      const maksInput = storedMaksCol >= 0 ? tr.querySelector(`input[data-col="${storedMaksCol}"]`) : null;
      const maks = maksInput ? parseFloat(maksInput.value) || 0 : 0;
      const percent = maks > 0 ? Math.round((sum / maks) * 100) : 0;
      
      const sumInput = tr.querySelector('input[data-calc-type="sum"]');
      const percentInput = tr.querySelector('input[data-calc-type="percent"]');
      
      if (sumInput) sumInput.value = sum;
      if (percentInput) percentInput.value = percent + '%';
    });
  }
}

function downloadEditedWorkbook() {
  if (!currentWorkbook || !currentSheetName) {
    alert('No sheet loaded. Please upload and process a file first.');
    return;
  }

  const inputs = Array.from(document.querySelectorAll('#sheetEditorContainer input')).filter(
    el => !el.hasAttribute('data-calc-type') // Exclude calculated columns
  );
  if (inputs.length === 0) {
    alert('No editable sheet data found.');
    return;
  }

  const maxRow = Math.max(...inputs.map((el) => Number(el.dataset.row)), 0);
  const maxCol = Math.max(...inputs.map((el) => Number(el.dataset.col)), 0);
  const updatedData = Array.from({ length: maxRow + 1 }, () => Array(maxCol + 1).fill(''));

  inputs.forEach((el) => {
    const row = Number(el.dataset.row);
    const col = Number(el.dataset.col);
    if (el.dataset.type === 'toggle') {
      updatedData[row][col] = el.checked ? 1 : 0;
    } else {
      updatedData[row][col] = el.value;
    }
  });

  const updatedSheet = window.XLSX.utils.aoa_to_sheet(updatedData);
  currentWorkbook.Sheets[currentSheetName] = updatedSheet;

  const outputName = currentFileName.toLowerCase().endsWith('.xlsx')
    ? currentFileName.replace(/\.xlsx$/i, '_edited.xlsx')
    : 'edited_schedule.xlsx';

  window.XLSX.writeFile(currentWorkbook, outputName);
}

document.getElementById('copyOneNoteBtn').addEventListener('click', async () => {
  if (!currentTeachersData || currentTeachersData.length === 0) {
    alert('No data to copy. Please upload a file first.');
    return;
  }

  try {
    // Build HTML table
    let html = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
    html += '<tr style="background-color: #667eea; color: white;"><th style="font-weight: bold;">Teacher Name</th><th style="font-weight: bold;">Scheduled Dates</th></tr>';
    
    currentTeachersData.forEach((teacher, idx) => {
      let dates = teacher.dates;
      if (filterDate) {
        dates = dates.filter(date => isDateOnOrAfter(date, filterDate));
      }
      
      const bgColor = idx % 2 === 0 ? 'white' : '#f5f5f5';
      let datesText = dates.length > 0 
        ? dates.map((date, dateIdx) => {
          const isBold = (dateIdx + 1) % 2 === 0;
          return isBold ? `<strong>${escapeHtml(date)}</strong>` : escapeHtml(date);
        }).join(' | ')
        : 'No schedules';
      
      html += `<tr style="background-color: ${bgColor};"><td>${escapeHtml(teacher.name)}</td><td>${datesText}</td></tr>`;
    });
    
    html += '</table>';

    // Copy as HTML to clipboard
    const blob = new Blob([html], { type: 'text/html' });
    const data = [new ClipboardItem({ 'text/html': blob })];
    await navigator.clipboard.write(data);
    alert('Table copied to clipboard! Paste it into OneNote using Ctrl+V');
  } catch (error) {
    console.error('Error copying to clipboard:', error);
    // Fallback: try copying as text if HTML copy fails
    try {
      let textTable = 'Teacher Name\tScheduled Dates\n';
      currentTeachersData.forEach(teacher => {
        let dates = teacher.dates;
        if (filterDate) {
          dates = dates.filter(date => isDateOnOrAfter(date, filterDate));
        }
        const datesText = dates.length > 0 ? dates.join(' | ') : 'No schedules';
        textTable += `${teacher.name}\t${datesText}\n`;
      });
      await navigator.clipboard.writeText(textTable);
      alert('Table copied to clipboard (as text)! Paste it into OneNote using Ctrl+V');
    } catch (fallbackError) {
      alert('Error copying to clipboard: ' + error.message);
    }
  }
});

document.getElementById('exportPdfBtn').addEventListener('click', async () => {
  if (!currentTeachersData || currentTeachersData.length === 0) {
    alert('No data to export. Please upload a file first.');
    return;
  }

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Add title
    doc.setFontSize(16);
    doc.text('Teacher Schedule', 15, 15);

    // Prepare table data with filtered dates
    const tableData = currentTeachersData.map(teacher => {
      let dates = teacher.dates;
      if (filterDate) {
        dates = dates.filter(date => isDateOnOrAfter(date, filterDate));
      }
      
      let datesText = dates.length > 0 ? dates.join(' | ') : 'No schedules';
      
      return [
        teacher.name,
        datesText
      ];
    });

    // Generate table
    doc.autoTable({
      head: [['Teacher Name', 'Scheduled Dates']],
      body: tableData,
      startY: 25,
      columnStyles: {
        0: { cellWidth: 60 },
        1: { cellWidth: 120 }
      },
      styles: {
        fontSize: 10,
        cellPadding: 6,
        halign: 'left',
        valign: 'top'
      },
      headStyles: {
        fillColor: [102, 126, 234],
        textColor: 255,
        fontStyle: 'bold'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      },
      bodyStyles: {
        valign: 'middle'
      },
      didDrawCell: (data) => {
        // Only process the dates column (index 1) in body cells
        if (data.column.index === 1 && data.row.section === 'body') {
          const text = data.cell.text[0];
          const dates = text.split(' | ');
          
          // Clear the default text
          const startX = data.cell.x + data.cell.padding('left');
          const startY = data.cell.y + data.cell.padding('top') + 7;
          
          doc.setFontSize(10);
          let currentX = startX;
          
          dates.forEach((date, idx) => {
            // Alternate between bold and normal - 2nd, 4th, 6th etc are bold
            const isBold = (idx + 1) % 2 === 0;
            doc.setFont(undefined, isBold ? 'bold' : 'normal');
            
            doc.text(date, currentX, startY);
            currentX += doc.getTextWidth(date) + 2;
            
            // Add separator if not last date
            if (idx < dates.length - 1) {
              doc.text(' | ', currentX, startY);
              currentX += doc.getTextWidth(' | ') + 2;
            }
          });
        }
      }
    });

    // Save the PDF
    doc.save('teacher_schedule.pdf');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('Error generating PDF: ' + error.message);
  }
});
