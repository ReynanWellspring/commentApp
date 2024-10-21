document.addEventListener('DOMContentLoaded', function () {
  const sheetUrls = {
    'EnglishSheet': 'https://script.google.com/macros/s/AKfycbxVb6AfquBhjNOMXYRy2914DZu_08x-u1-4hbKnYe49HrUYDxkPNb7ltUxx3IAdgDdRiw/exec',
    'MathSheet': 'https://script.google.com/macros/s/AKfycbyQoaYAruRHG8mJvpygzOAUSdWp9bnZAZ7958uFifW6GhlBZ4K_VIDMhJ2IMxsKAgxR/exec',
    'ScienceSheet': 'https://script.google.com/macros/s/AKfycbyhyWjEqnbXDfbXCMX5U4saJ8qSrpoqUk4eixDC7dboc2dF7PZbg0X91-ewbo0vHZZPKw/exec',
    'ICTSheet': 'https://script.google.com/macros/s/AKfycbza8qM9Brr9vl8qpeomrApgc8XbsF7qUu0inqUV3HQza30F45lpNNmKrDfbiFiGQSeT/exec',
  };

  const googleSheetId = '1GgdYi_m8trTQnkL0kp4P2xb_PBuqinmsvn1hZN-G1pM';

  let selectedSheet = localStorage.getItem('selectedSheet') || 'Sheet1';
  let selectedSubject = localStorage.getItem('selectedSubject') || 'ICTSheet';
  let sheetData = JSON.parse(localStorage.getItem('sheetData')) || {};

  const classOptions = {
    'Sheet1': ['1A1', '1A2', '1A3', '1A4', '1A5'],
    'Sheet2': ['2A1', '2A2', '2A3', '2A4', '2A5', '2A6'],
    'Sheet3': ['3A1', '3A2', '3A3', '3A4', '3A5', '3A6'],
    'Sheet4': ['4A2', '4A3', '4A4', '4A5', '4A6', '4A7'],
    'Sheet5': ['5A1', '5A2', '5A3', '5A4', '5A5', '5A6']
  };

  const updateClassDropdown = (grade) => {
    const classSelect = document.getElementById('classSelect');
    classSelect.innerHTML = '<option value="" disabled selected>Select a class</option>';

    const classes = classOptions[grade] || [];
    classes.forEach(className => {
      const option = document.createElement('option');
      option.value = className;
      option.textContent = className;
      classSelect.appendChild(option);
    });

    classSelect.disabled = false;
  };

  const fetchDataFromGoogleSheet = async (sheetName) => {
    const baseUrl = `https://docs.google.com/spreadsheets/d/${googleSheetId}/gviz/tq?sheet=${sheetName}&tqx=out:json`;
    try {
      const response = await fetch(baseUrl);
      const text = await response.text();
      const json = JSON.parse(text.substr(47).slice(0, -2));
      return json.table.rows.map(row => row.c.map(cell => (cell ? cell.v : '')));
    } catch (error) {
      console.error('Error fetching data from Google Sheet:', error);
    }
  };

  const updateTableData = async (selectedClass) => {
    const tableBody = document.getElementById('commentTable').querySelector('tbody');
    tableBody.innerHTML = '';

    const data = await fetchDataFromGoogleSheet(selectedClass);
    
    if (data && data.length > 0) {
      data.forEach((rowData, rowIndex) => {
        const newRow = tableBody.insertRow();

        rowData.slice(0, 10).forEach((cellData, index) => {
          const cell = newRow.insertCell(index);
          cell.textContent = cellData;
        });

        addCommentColumns(newRow, rowData[4], rowData[3], rowIndex);
      });
    }

    // Add event listeners to enforce dropdown width
    enforceDropdownWidth();

    // Restore saved data if available
    restoreSavedData();
  };

  const addCommentColumns = (newRow, gender, firstName, rowIndex) => {
    const selectedComments = [];

    ["introduction", "behavior", "classwork", "participation", "improvements"].forEach((category, index) => {
      const cell = newRow.insertCell();
      
      // Create container for the dropdown
      const container = document.createElement('div');
      container.classList.add('comment-container');

      const select = document.createElement('select');
      select.dataset.category = category;
      select.dataset.rowIndex = rowIndex;
      select.dataset.columnIndex = index;
      populateOptions(select, category);
      
      // Append select to container, and container to the cell
      container.appendChild(select);
      cell.appendChild(container);
      
      selectedComments.push(select);
    });

    const summaryCell = newRow.insertCell();
    summaryCell.classList.add('summary-column');
    const summaryDiv = document.createElement('div');
    summaryDiv.contentEditable = true;
    summaryCell.appendChild(summaryDiv);

    selectedComments.forEach((element) => {
      element.addEventListener('change', () => {
        updateSummary(firstName, gender, selectedComments, summaryDiv);
        saveData(rowIndex, element.dataset.columnIndex, element.value);
      });

      if (element.tagName === 'INPUT') {
        element.addEventListener('input', () => {
          updateSummary(firstName, gender, selectedComments, summaryDiv);
          saveData(rowIndex, element.dataset.columnIndex, element.value);
        });
      }
    });
  };

  const populateOptions = (select, category) => {
    select.innerHTML = '';
    const data = sheetData[category];
    if (data && Array.isArray(data)) {
      data.forEach(item => {
        const option = document.createElement('option');
        option.value = item;
        option.textContent = item;
        select.appendChild(option);
      });
    }
  };

  const updateSummary = (firstName, gender, selectedComments, summaryCell) => {
    const pronouns = {
      'M': { subject: 'He', object: 'him', possessive: 'his' },
      'F': { subject: 'She', object: 'her', possessive: 'her' }
    };

    const currentPronoun = pronouns[gender] || pronouns['M'];

    let summaryText = `${firstName} `;

    const introduction = selectedComments[0].value;
    const behavior = `${currentPronoun.subject} ${selectedComments[1].value.toLowerCase()}`;
    const classwork = `${currentPronoun.possessive.charAt(0).toUpperCase() + currentPronoun.possessive.slice(1)} classwork ${selectedComments[2].value.toLowerCase()}`;
    const participation = `${currentPronoun.subject} ${selectedComments[3].value.toLowerCase()}`;
    const improvements = selectedComments[4].value;

    summaryText += `${introduction}. ${behavior}. ${classwork}. ${participation}. ${improvements}.`;

    summaryText = summaryText.replace(/\s+/g, ' ').replace(/(\.\s*)+/g, '. ').trim();

    summaryCell.textContent = summaryText;
  };

  const enforceDropdownWidth = () => {
    document.querySelectorAll('.comment-container select').forEach(select => {
      select.addEventListener('mousedown', function () {
        this.style.width = `${this.parentElement.clientWidth}px`;
      });

      select.addEventListener('blur', function () {
        this.style.width = '100%';
      });
    });
  };

  const saveData = (rowIndex, columnIndex, value) => {
    let savedData = JSON.parse(localStorage.getItem('savedComments')) || {};
    if (!savedData[rowIndex]) {
      savedData[rowIndex] = {};
    }
    savedData[rowIndex][columnIndex] = value;
    localStorage.setItem('savedComments', JSON.stringify(savedData));
  };

  const restoreSavedData = () => {
    let savedData = JSON.parse(localStorage.getItem('savedComments')) || {};
    for (let rowIndex in savedData) {
      for (let columnIndex in savedData[rowIndex]) {
        let select = document.querySelector(`select[data-row-index="${rowIndex}"][data-column-index="${columnIndex}"]`);
        if (select) {
          select.value = savedData[rowIndex][columnIndex];
        }
      }
    }
  };

  const resetAll = () => {
    document.getElementById('commentTable').querySelector('tbody').innerHTML = '';
    localStorage.removeItem('savedComments');
  };

  document.getElementById('gradeSelect').addEventListener('change', function () {
    selectedSheet = this.value;
    localStorage.setItem('selectedSheet', selectedSheet);
    updateClassDropdown(selectedSheet);
  });

  document.getElementById('subjectSelect').addEventListener('change', function () {
    selectedSubject = this.value;
    localStorage.setItem('selectedSubject', selectedSubject);
  });

  document.getElementById('classSelect').addEventListener('change', function () {
    document.getElementById('goBtn').disabled = false;
  });

  document.getElementById('goBtn').addEventListener('click', async function() {
    const selectedClass = document.getElementById('classSelect').value;
    await updateTableData(selectedClass);
    populateDropdowns();
  });

  document.getElementById('saveBtn').addEventListener('click', function() {
    // Save data is handled dynamically upon dropdown changes
    alert("Data saved successfully!");
  });

  document.getElementById('createNewBtn').addEventListener('click', function() {
    resetAll();
  });
});
