let currentTeachersData = [];
let filterDate = null;
let disabledDates = new Set();
let currentWorkbook = null;
let allSheetNames = [];

// Initialize Flatpickr date picker
window.addEventListener('load', () => {
  flatpickr('#filterDate', {
    mode: 'single',
    dateFormat: 'd.m.Y',
    defaultDate: new Date(),
    locale: 'de'
  });
  
  // Tab button listeners
  document.getElementById('tabTeacherBtn').addEventListener('click', () => switchTab('teacher'));
  document.getElementById('tabDateBtn').addEventListener('click', () => switchTab('date'));
  
  // Sheet selector listener
  document.getElementById('sheetSelector').addEventListener('change', (e) => {
    if (e.target.value) {
      loadDateScheduleSheet(e.target.value);
    }
  });
});

function switchTab(tab) {
  const teacherBtn = document.getElementById('tabTeacherBtn');
  const dateBtn = document.getElementById('tabDateBtn');
  const teacherTab = document.getElementById('teacherTab');
  const dateTab = document.getElementById('dateTab');
  const exportPdf = document.getElementById('exportPdfBtn');
  const copyOneNote = document.getElementById('copyOneNoteBtn');
  const copyDateOneNote = document.getElementById('copyDateOneNoteBtn');
  
  if (tab === 'teacher') {
    teacherBtn.classList.add('active');
    dateBtn.classList.remove('active');
    teacherTab.style.display = 'block';
    dateTab.style.display = 'none';
    exportPdf.style.display = 'block';
    copyOneNote.style.display = 'block';
    copyDateOneNote.style.display = 'none';
  } else {
    teacherBtn.classList.remove('active');
    dateBtn.classList.add('active');
    teacherTab.style.display = 'none';
    dateTab.style.display = 'block';
    exportPdf.style.display = 'none';
    copyOneNote.style.display = 'none';
    copyDateOneNote.style.display = 'block';
  }
}

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
  // Parse DD.MM.YYYY format for date filter
  filterDate = null;
  if (filterDateInput.value.trim()) {
    const match = filterDateInput.value.trim().match(/(\d{2})\.(\d{2})\.(\d{4})/);
    if (match) {
      const day = parseInt(match[1]);
      const month = parseInt(match[2]);
      filterDate = { day, month };
    }
  }

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
    
    // Load workbook for sheet selection
    const buffer = await file.arrayBuffer();
    currentWorkbook = window.XLSX.read(buffer, { type: 'array' });
    allSheetNames = currentWorkbook.SheetNames;
    
    // Populate sheet selector dropdown
    const selector = document.getElementById('sheetSelector');
    selector.innerHTML = '<option value="">-- Select a sheet --</option>';
    allSheetNames.forEach(sheetName => {
      if (sheetName !== sheetNameInput.value) { // Exclude the teacher schedule sheet
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        selector.appendChild(option);
      }
    });
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
                    const formatted = dateStr.replace(/(\([^)]+\))/, '<span class="date-time">$1</span>');
                    const isDisabled = disabledDates.has(dateStr);
                    return `<span class="date-badge ${isDisabled ? 'disabled' : ''}" data-date="${dateStr}">${formatted}</span>`;
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
  
  // Add click handlers to date badges
  document.querySelectorAll('.date-badge[data-date]').forEach(badge => {
    badge.addEventListener('click', (e) => {
      const dateStr = badge.getAttribute('data-date');
      if (disabledDates.has(dateStr)) {
        disabledDates.delete(dateStr);
        badge.classList.remove('disabled');
      } else {
        disabledDates.add(dateStr);
        badge.classList.add('disabled');
      }
    });
  });
}

function isDateOnOrAfter(dateString, filterDate) {
  // If no filter date set, include all dates
  if (!filterDate) return true;
  
  // Extract DD.MM from the date string (e.g., "04.05" from "04.05 (08.30-11.45)")
  const match = dateString.toString().match(/(\d{2})\.(\d{2})/);
  if (!match) return true;
  
  const day = parseInt(match[1]);
  const month = parseInt(match[2]);
  
  // Compare month first, then day
  if (month !== filterDate.month) {
    return month > filterDate.month;
  }
  return day >= filterDate.day;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function loadDateScheduleSheet(sheetName) {
  if (!currentWorkbook) {
    alert('No workbook loaded.');
    return;
  }
  
  const worksheet = currentWorkbook.Sheets[sheetName];
  if (!worksheet) {
    alert('Sheet not found.');
    return;
  }
  
  // Read the data starting from B6
  // Get raw data to preserve values
  const range = worksheet['!ref'];
  const data = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  
  // Extract data from B6 to H6 and down
  // Row 6 is index 5 (0-based), columns B-H are indices 1-7
  const tableData = [];
  
  if (data.length > 5) {
    // Get headers from row 6 (B6 to H6)
    const headerRow = data[5];
    const headers = headerRow.slice(1, 8); // B6 to H6 (indices 1-7)
    tableData.push(headers);
    
    // Get data rows starting from row 7
    for (let i = 6; i < data.length; i++) {
      const row = data[i];
      if (row && row.length > 1) {
        const firstCell = row[0]; // Column A
        if (!firstCell || String(firstCell).trim() === '') break; // Stop at empty row
        
        // Get cells B to H
        const rowData = row.slice(1, 8);
        tableData.push(rowData);
      } else {
        break;
      }
    }
  }
  
  // Render the table
  const table = document.getElementById('dateScheduleTable');
  let html = '';
  
  tableData.forEach((row, rowIdx) => {
    const isHeader = rowIdx === 0;
    const bgColor = !isHeader && rowIdx % 2 === 0 ? 'white' : '#f5f5f5';
    const cellTag = isHeader ? 'th' : 'td';
    const style = isHeader 
      ? 'style="background-color: #667eea; color: white; font-weight: bold; padding: 8px; border: 1px solid #ddd;"'
      : `style="padding: 8px; border: 1px solid #ddd; background-color: ${bgColor};"`;
    
    html += '<tr>';
    row.forEach(cell => {
      html += `<${cellTag} ${style}>${escapeHtml(String(cell || ''))}</${cellTag}>`;
    });
    html += '</tr>';
  });
  
  table.innerHTML = html;
  document.getElementById('dateTabContent').style.display = 'block';
  
  // Store the table data for copying
  window.currentDateTableData = tableData;
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
      // Filter out disabled dates
      dates = dates.filter(date => !disabledDates.has(date));
      
      // Skip teachers with no scheduled dates
      if (dates.length === 0) return;
      
      const bgColor = idx % 2 === 0 ? 'white' : '#f5f5f5';
      let datesText = dates.map((date, dateIdx) => {
        const isBold = (dateIdx + 1) % 2 === 0;
        return isBold ? `<strong>${escapeHtml(date)}</strong>` : escapeHtml(date);
      }).join(' | ');
      
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
        // Filter out disabled dates
        dates = dates.filter(date => !disabledDates.has(date));
        
        // Skip teachers with no scheduled dates
        if (dates.length === 0) return;
        
        const datesText = dates.join(' | ');
        textTable += `${teacher.name}\t${datesText}\n`;
      });
      await navigator.clipboard.writeText(textTable);
      alert('Table copied to clipboard (as text)! Paste it into OneNote using Ctrl+V');
    } catch (fallbackError) {
      alert('Error copying to clipboard: ' + error.message);
    }
  }
});

document.getElementById('copyDateOneNoteBtn').addEventListener('click', async () => {
  if (!window.currentDateTableData || window.currentDateTableData.length === 0) {
    alert('No data to copy. Please select a sheet first.');
    return;
  }

  try {
    // Build HTML table
    let html = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
    
    window.currentDateTableData.forEach((row, rowIdx) => {
      const isHeader = rowIdx === 0;
      const bgColor = isHeader ? '#667eea' : (rowIdx % 2 === 0 ? 'white' : '#f5f5f5');
      const textColor = isHeader ? 'white' : 'black';
      const fontWeight = isHeader ? 'bold' : 'normal';
      
      html += `<tr style="background-color: ${bgColor}; color: ${textColor};">`;
      row.forEach(cell => {
        const cellTag = isHeader ? 'th' : 'td';
        html += `<${cellTag} style="font-weight: ${fontWeight}; padding: 8px; border: 1px solid #ddd;">${escapeHtml(String(cell || ''))}</${cellTag}>`;
      });
      html += '</tr>';
    });
    
    html += '</table>';

    // Copy as HTML to clipboard
    const blob = new Blob([html], { type: 'text/html' });
    const data = [new ClipboardItem({ 'text/html': blob })];
    await navigator.clipboard.write(data);
    alert('Table copied to clipboard! Paste it into OneNote using Ctrl+V');
  } catch (error) {
    console.error('Error copying to clipboard:', error);
    // Fallback: try copying as text
    try {
      let textTable = window.currentDateTableData.map(row => row.join('\t')).join('\n');
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
    const tableData = currentTeachersData
      .map(teacher => {
        let dates = teacher.dates;
        if (filterDate) {
          dates = dates.filter(date => isDateOnOrAfter(date, filterDate));
        }
        // Filter out disabled dates
        dates = dates.filter(date => !disabledDates.has(date));
        
        // Only include teachers with scheduled dates
        if (dates.length === 0) return null;
        
        let datesText = dates.join(' | ');
        
        return [
          teacher.name,
          datesText
        ];
      })
      .filter(row => row !== null);

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
      }
    });

    // Save the PDF
    doc.save('teacher_schedule.pdf');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('Error generating PDF: ' + error.message);
  }
});
