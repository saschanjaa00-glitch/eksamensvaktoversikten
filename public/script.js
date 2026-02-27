let currentTeachersData = [];
let filterDate = null;
let disabledDates = new Set();
let currentWorkbook = null;
let allSheetNames = [];
let showHiddenColsDate = false;
let showHiddenColsMulti = false;

// Check if a sheet name is formatted as a date (DD.MM)
function isDateFormattedSheet(sheetName) {
  return /^\d{2}\.\d{2}$/.test(String(sheetName).trim());
}

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
  
  // Tab button listeners
  document.getElementById('tabTeacherBtn').addEventListener('click', () => switchTab('teacher'));
  document.getElementById('tabDateBtn').addEventListener('click', () => switchTab('date'));
  document.getElementById('tabMultiBtn').addEventListener('click', () => switchTab('multi'));
  
  // Sheet selector listener
  document.getElementById('sheetSelector').addEventListener('change', (e) => {
    if (e.target.value) {
      loadDateScheduleSheet(e.target.value);
    }
  });
  
  // Multi-sheet loader listener
  document.getElementById('loadMultiSheetsBtn').addEventListener('click', loadMultipleSheets);
  
  // Column visibility toggles
  document.getElementById('dateShowHiddenCols').addEventListener('change', () => {
    showHiddenColsDate = document.getElementById('dateShowHiddenCols').checked;
    if (window.currentDateTableDataFull) {
      renderDateScheduleTable(window.currentDateTableDataFull);
    }
  });
  
  document.getElementById('multiShowHiddenCols').addEventListener('change', () => {
    showHiddenColsMulti = document.getElementById('multiShowHiddenCols').checked;
    if (window.currentMultiTableDataFull) {
      renderMultiScheduleTable(window.currentMultiTableDataFull);
    }
  });

  // Upload form listener
  const uploadForm = document.getElementById('uploadForm');
  if (uploadForm) {
    uploadForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      console.log('Form submitted');

      const fileInput = document.getElementById('file');
      const sheetNameInput = document.getElementById('sheetName');
      const filterDateInput = document.getElementById('filterDate');
      const file = fileInput.files[0];
      const sheetName = sheetNameInput.value || 'Lærervakter';

      if (!file) {
        showError('Please select a file');
        return;
      }

      console.log('File selected:', file.name);

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
        console.log('Sending file to server');
        console.log('File object:', file);
        console.log('File name:', file.name);
        console.log('File size:', file.size);
        console.log('File type:', file.type);
        
        const formData = new FormData();
        formData.append('file', file);
        formData.append('sheetName', sheetName);
        
        console.log('FormData created, keys:', Array.from(formData.keys()));

        const response = await fetch('/api/upload', {
          method: 'POST',
          body: formData
        });

        console.log('Response status:', response.status);

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Failed to process file');
        }

        const data = await response.json();
        console.log('Data received:', data);
        displayResults(data);
      
      // Load workbook for sheet selection
      const buffer = await file.arrayBuffer();
      currentWorkbook = window.XLSX.read(buffer, { type: 'array' });
      allSheetNames = currentWorkbook.SheetNames;
      
      // Populate sheet selector dropdown
      const selector = document.getElementById('sheetSelector');
      selector.innerHTML = '<option value="">-- Select a sheet --</option>';
      allSheetNames.forEach(sheetName => {
        if (sheetName !== sheetNameInput.value && isDateFormattedSheet(sheetName)) { // Exclude teacher sheet and only show date-formatted sheets
          const option = document.createElement('option');
          option.value = sheetName;
          option.textContent = sheetName;
          selector.appendChild(option);
        }
      });
      
      // Populate multi-sheet checkboxes
      const checkboxContainer = document.getElementById('multiSheetCheckboxes');
      checkboxContainer.innerHTML = '';
      const dateFormattedSheets = allSheetNames.filter(sheetName => 
        sheetName !== sheetNameInput.value && isDateFormattedSheet(sheetName)
      );
      
      // Hide select all label if no date sheets
      const selectAllLabel = document.getElementById('selectAllLabel');
      selectAllLabel.style.display = dateFormattedSheets.length > 0 ? 'flex' : 'none';
      
      // Reset select all checkbox
      const selectAllCheckbox = document.getElementById('selectAllCheckbox');
      selectAllCheckbox.checked = false;
      
      dateFormattedSheets.forEach(sheetName => {
        const label = document.createElement('label');
        label.style.display = 'flex';
        label.style.alignItems = 'center';
        label.style.gap = '8px';
        label.style.cursor = 'pointer';
        label.style.margin = '0';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = sheetName;
        checkbox.className = 'sheet-checkbox';
        checkbox.style.cursor = 'pointer';
        
        const text = document.createElement('span');
        text.textContent = sheetName;
        
        label.appendChild(checkbox);
        label.appendChild(text);
        checkboxContainer.appendChild(label);
      });
      
      // Add select all functionality
      selectAllCheckbox.addEventListener('change', (e) => {
        const allCheckboxes = document.querySelectorAll('.sheet-checkbox');
        allCheckboxes.forEach(cb => cb.checked = e.target.checked);
      });
      
      // Update select all checkbox state when individual checkboxes change
      document.querySelectorAll('.sheet-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', () => {
          const allCheckboxes = document.querySelectorAll('.sheet-checkbox');
          const allChecked = Array.from(allCheckboxes).every(cb => cb.checked);
          const someChecked = Array.from(allCheckboxes).some(cb => cb.checked);
          selectAllCheckbox.checked = allChecked;
          selectAllCheckbox.indeterminate = someChecked && !allChecked;
        });
      });
    } catch (error) {
      console.error('Error uploading file:', error);
      showError(error.message);
    } finally {
      showLoading(false);
    }
  }) || console.warn('Upload form not found');
  } else {
    console.warn('Window load event: uploadForm element not found');
  }
});

function switchTab(tab) {
  const teacherBtn = document.getElementById('tabTeacherBtn');
  const dateBtn = document.getElementById('tabDateBtn');
  const multiBtn = document.getElementById('tabMultiBtn');
  const teacherTab = document.getElementById('teacherTab');
  const dateTab = document.getElementById('dateTab');
  const multiTab = document.getElementById('multiTab');
  const exportPdf = document.getElementById('exportPdfBtn');
  const copyOneNote = document.getElementById('copyOneNoteBtn');
  const copyDateOneNote = document.getElementById('copyDateOneNoteBtn');
  const saveDatePng = document.getElementById('saveDatePngBtn');
  const copyMultiOneNote = document.getElementById('copyMultiOneNoteBtn');
  const saveMultiPng = document.getElementById('saveMultiPngBtn');
  
  // Reset all
  teacherBtn.classList.remove('active');
  dateBtn.classList.remove('active');
  multiBtn.classList.remove('active');
  teacherTab.style.display = 'none';
  dateTab.style.display = 'none';
  multiTab.style.display = 'none';
  exportPdf.style.display = 'none';
  copyOneNote.style.display = 'none';
  copyDateOneNote.style.display = 'none';
  saveDatePng.style.display = 'none';
  copyMultiOneNote.style.display = 'none';
  saveMultiPng.style.display = 'none';
  
  // Set active tab
  if (tab === 'teacher') {
    teacherBtn.classList.add('active');
    teacherTab.style.display = 'block';
    exportPdf.style.display = 'block';
    copyOneNote.style.display = 'block';
  } else if (tab === 'date') {
    dateBtn.classList.add('active');
    dateTab.style.display = 'block';
    copyDateOneNote.style.display = 'block';
    saveDatePng.style.display = 'block';
  } else if (tab === 'multi') {
    multiBtn.classList.add('active');
    multiTab.style.display = 'block';
    copyMultiOneNote.style.display = 'block';
    saveMultiPng.style.display = 'block';
  }
}

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
  
  // Extract data from A6 to H6 and down (with all columns)
  // Row 6 is index 5 (0-based), columns A-H are indices 0-7
  const fullTableData = [];
  
  if (data.length > 5) {
    // Get headers from row 6 (A6 to H6) - all columns
    const headerRow = data[5];
    const headers = headerRow.slice(0, 8); // A to H (indices 0-7)
    fullTableData.push(headers);
    
    // Get data rows starting from row 7
    // Collect rows but look ahead for content after empty rows
    let emptyRowCount = 0;
    for (let i = 6; i < data.length; i++) {
      const row = data[i];
      const firstCell = row && row.length > 0 ? row[0] : '';
      const hasContent = firstCell && String(firstCell).trim() !== '';
      
      if (hasContent) {
        // Found content, reset empty row counter
        emptyRowCount = 0;
        // Get all cells A to H
        const rowData = row.slice(0, 8);
        fullTableData.push(rowData);
      } else {
        // Empty row - increment counter but keep looking
        emptyRowCount++;
        // Only stop if we've had 5 consecutive empty rows with no more content ahead
        if (emptyRowCount >= 5) {
          // Check if there's any content in the next few rows
          let foundContentAhead = false;
          for (let j = i + 1; j < Math.min(i + 6, data.length); j++) {
            const aheadRow = data[j];
            const aheadFirstCell = aheadRow && aheadRow.length > 0 ? aheadRow[0] : '';
            if (aheadFirstCell && String(aheadFirstCell).trim() !== '') {
              foundContentAhead = true;
              break;
            }
          }
          if (!foundContentAhead) break;
        }
      }
    }
    
    // Remove trailing empty rows
    while (fullTableData.length > 1) {
      const lastRow = fullTableData[fullTableData.length - 1];
      const lastRowFirstCell = lastRow[0];
      if (!lastRowFirstCell || String(lastRowFirstCell).trim() === '') {
        fullTableData.pop();
      } else {
        break;
      }
    }
  }
  
  // Store full data and render
  window.currentDateTableDataFull = fullTableData;
  showHiddenColsDate = false;
  document.getElementById('dateShowHiddenCols').checked = false;
  renderDateScheduleTable(fullTableData);
  document.getElementById('dateTabContent').style.display = 'block';
}

function renderDateScheduleTable(fullTableData) {
  // Filter columns B and C if toggle is off
  const filteredData = fullTableData.map((row, rowIdx) => {
    if (showHiddenColsDate) {
      return row; // Show all columns
    } else {
      // Hide B and C columns - only show A and D-H
      return row.slice(0, 1).concat(row.slice(3, 8));
    }
  });
  
  // Render the table
  const table = document.getElementById('dateScheduleTable');
  let html = '';
  
  filteredData.forEach((row, rowIdx) => {
    const isHeader = rowIdx === 0;
    const cellTag = isHeader ? 'th' : 'td';
    const style = isHeader 
      ? 'style="font-weight: bold; padding: 8px; border: 1px solid #000; white-space: nowrap;"'
      : 'style="padding: 8px; border: 1px solid #000; white-space: nowrap;"';
    
    html += '<tr>';
    row.forEach(cell => {
      html += `<${cellTag} ${style}>${escapeHtml(String(cell || ''))}</${cellTag}>`;
    });
    html += '</tr>';
  });
  
  table.innerHTML = html;
  
  // Store the filtered table data for copying
  window.currentDateTableData = filteredData;
}

function loadMultipleSheets() {
  if (!currentWorkbook) {
    alert('No workbook loaded.');
    return;
  }
  
  // Get all checked sheets
  const checkedBoxes = Array.from(document.querySelectorAll('#multiSheetCheckboxes input[type="checkbox"]:checked'));
  if (checkedBoxes.length === 0) {
    alert('Please select at least one sheet.');
    return;
  }
  
  const sheetNames = checkedBoxes.map(checkbox => checkbox.value);
  const combinedData = [];
  
  // Load data from each selected sheet
  sheetNames.forEach(sheetName => {
    const worksheet = currentWorkbook.Sheets[sheetName];
    if (worksheet) {
      const data = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      
      // Extract data from A6 to H6 and down
      if (data.length > 5) {
        const headerRow = data[5];
        const headers = headerRow.slice(0, 8); // A to H (all columns)
        
        // Add sheet name as a header separator (8 columns for full data)
        combinedData.push([sheetName, '', '', '', '', '', '', '']);
        combinedData.push(headers);
        
        // Get data rows
        let emptyRowCount = 0;
        for (let i = 6; i < data.length; i++) {
          const row = data[i];
          const firstCell = row && row.length > 0 ? row[0] : '';
          const hasContent = firstCell && String(firstCell).trim() !== '';
          
          if (hasContent) {
            emptyRowCount = 0;
            const rowData = row.slice(0, 8); // All columns for full data
            combinedData.push(rowData);
          } else {
            emptyRowCount++;
            if (emptyRowCount >= 5) {
              let foundContentAhead = false;
              for (let j = i + 1; j < Math.min(i + 6, data.length); j++) {
                const aheadRow = data[j];
                const aheadFirstCell = aheadRow && aheadRow.length > 0 ? aheadRow[0] : '';
                if (aheadFirstCell && String(aheadFirstCell).trim() !== '') {
                  foundContentAhead = true;
                  break;
                }
              }
              if (!foundContentAhead) break;
            }
          }
        }
      }
    }
  });
  
  // Remove trailing empty rows
  while (combinedData.length > 0) {
    const lastRow = combinedData[combinedData.length - 1];
    const lastRowFirstCell = lastRow[0];
    if (!lastRowFirstCell || String(lastRowFirstCell).trim() === '') {
      combinedData.pop();
    } else {
      break;
    }
  }
  
  // Render the table
  const table = document.getElementById('multiScheduleTable');
  window.currentMultiTableDataFull = combinedData;
  showHiddenColsMulti = false;
  document.getElementById('multiShowHiddenCols').checked = false;
  renderMultiScheduleTable(combinedData);
  document.getElementById('multiTabContent').style.display = 'block';
}

function renderMultiScheduleTable(fullTableData) {
  // Filter columns B and C if toggle is off
  const filteredData = fullTableData.map((row, rowIdx) => {
    const isSheetHeader = row[0] && allSheetNames.includes(String(row[0]));
    
    if (isSheetHeader || showHiddenColsMulti) {
      return row; // Show all columns for sheet headers or if toggle is on
    } else {
      // Hide B and C columns - only show A and D-H
      return row.slice(0, 1).concat(row.slice(3, 8));
    }
  });
  
  // Render the table
  const table = document.getElementById('multiScheduleTable');
  let html = '';
  
  filteredData.forEach((row, rowIdx) => {
    const isSheetHeader = row[0] && allSheetNames.includes(String(row[0]));
    const isColumnHeader = !isSheetHeader && filteredData.length > rowIdx - 1 && 
                          filteredData[rowIdx - 1] && allSheetNames.includes(String(filteredData[rowIdx - 1][0]));
    
    const cellTag = (isSheetHeader || isColumnHeader) ? 'th' : 'td';
    const style = (isSheetHeader || isColumnHeader)
      ? 'style="font-weight: bold; padding: 8px; border: 1px solid #000; white-space: nowrap;"'
      : 'style="padding: 8px; border: 1px solid #000; white-space: nowrap;"';
    
    const trStyle = isSheetHeader ? 'style="background-color: #e0f0ff;"' : '';
    html += `<tr ${trStyle}>`;
    row.forEach(cell => {
      html += `<${cellTag} ${style}>${escapeHtml(String(cell || ''))}</${cellTag}>`;
    });
    html += '</tr>';
  });
  
  table.innerHTML = html;
  
  // Store the filtered table data for copying
  window.currentMultiTableData = filteredData;
}

document.getElementById('copyOneNoteBtn').addEventListener('click', async () => {
  if (!currentTeachersData || currentTeachersData.length === 0) {
    alert('No data to copy. Please upload a file first.');
    return;
  }

  try {
    // Build HTML table
    let html = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
    html += '<tr style="background-color: #e0f0ff;"><th style="font-weight: bold; padding: 8px; border: 1px solid #000;">Teacher Name</th><th style="font-weight: bold; padding: 8px; border: 1px solid #000;">Scheduled Dates</th></tr>';
    
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
      const bgColor = isHeader ? '#e0f0ff' : 'white';
      const fontWeight = isHeader ? 'bold' : 'normal';
      
      html += `<tr style="background-color: ${bgColor};">`;
      row.forEach(cell => {
        const cellTag = isHeader ? 'th' : 'td';
        html += `<${cellTag} style="font-weight: ${fontWeight}; padding: 8px; border: 1px solid #000;">${escapeHtml(String(cell || ''))}</${cellTag}>`;
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

document.getElementById('copyMultiOneNoteBtn').addEventListener('click', async () => {
  if (!window.currentMultiTableData || window.currentMultiTableData.length === 0) {
    alert('No data to copy. Please load sheets first.');
    return;
  }

  try {
    // Build HTML table
    let html = '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
    
    window.currentMultiTableData.forEach((row, rowIdx) => {
      const isSheetHeader = row[0] && allSheetNames.includes(String(row[0]));
      const isColumnHeader = !isSheetHeader && window.currentMultiTableData.length > rowIdx - 1 && 
                            window.currentMultiTableData[rowIdx - 1] && 
                            allSheetNames.includes(String(window.currentMultiTableData[rowIdx - 1][0]));
      
      const bgColor = (isSheetHeader || isColumnHeader) ? '#e0f0ff' : 'white';
      const fontWeight = (isSheetHeader || isColumnHeader) ? 'bold' : 'normal';
      
      html += `<tr style="background-color: ${bgColor};">`;
      row.forEach(cell => {
        const cellTag = (isSheetHeader || isColumnHeader) ? 'th' : 'td';
        html += `<${cellTag} style="font-weight: ${fontWeight}; padding: 8px; border: 1px solid #000;">${escapeHtml(String(cell || ''))}</${cellTag}>`;
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
      let textTable = window.currentMultiTableData.map(row => row.join('\t')).join('\n');
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

document.getElementById('saveDatePngBtn').addEventListener('click', async () => {
  if (!window.currentDateTableData || window.currentDateTableData.length === 0) {
    alert('No data to save. Please select a sheet first.');
    return;
  }

  try {
    const element = document.getElementById('dateScheduleTable');
    if (!element) {
      alert('Table not found.');
      return;
    }

    const canvas = await html2canvas(element, {
      scale: 2,
      backgroundColor: '#ffffff'
    });

    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/png');
    link.download = 'schedule_by_date.png';
    link.click();
  } catch (error) {
    console.error('Error saving PNG:', error);
    alert('Error saving PNG: ' + error.message);
  }
});

document.getElementById('saveMultiPngBtn').addEventListener('click', async () => {
  if (!window.currentMultiTableData || window.currentMultiTableData.length === 0) {
    alert('No data to save. Please load sheets first.');
    return;
  }

  try {
    const element = document.getElementById('multiScheduleTable');
    if (!element) {
      alert('Table not found.');
      return;
    }

    const canvas = await html2canvas(element, {
      scale: 2,
      backgroundColor: '#ffffff'
    });

    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/png');
    link.download = 'multi_sheet_export.png';
    link.click();
  } catch (error) {
    console.error('Error saving PNG:', error);
    alert('Error saving PNG: ' + error.message);
  }
});
