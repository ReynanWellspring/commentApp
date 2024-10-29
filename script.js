const maxLengthDisplay = 200; // Limit the number of displayed characters to 200

// Define the Google Apps Script URLs to retrieve data for each subject
// 'EnglishSheet': 'https://script.google.com/macros/s/AKfycbxVb6AfquBhjNOMXYRy2914DZu_08x-u1-4hbKnYe49HrUYDxkPNb7ltUxx3IAdgDdRiw/exec',
// 'MathSheet': 'https://script.google.com/macros/s/AKfycbyQoaYAruRHG8mJvpygzOAUSdWp9bnZAZ7958uFifW6GhlBZ4K_VIDMhJ2IMxsKAgxR/exec',
// 'ScienceSheet': 'https://script.google.com/macros/s/AKfycbyhyWjEqnbXDfbXCMX5U4saJ8qSrpoqUk4eixDC7dboc2dF7PZbg0X91-ewbo0vHZZPKw/exec',
// 'ICTSheet': 'https://script.google.com/macros/s/AKfycbza8qM9Brr9vl8qpeomrApgc8XbsF7qUu0inqUV3HQza30F45lpNNmKrDfbiFiGQSeT/exec',

document.addEventListener('DOMContentLoaded', function () {

  const googleSheetClassesId = '1GgdYi_m8trTQnkL0kp4P2xb_PBuqinmsvn1hZN-G1pM'; // ID of Google Sheet Classes
  const googleSheetICTId = '1utKoEwDu208dlHLs0-Pfa_m4Nm06dTkpoqhqeSn2YNE'; // ID of Google Sheet ICT

  // Store selected spreadsheet and subject information in localStorage for easy restoration later
  let gradeSelected = localStorage.getItem('selectedSheet') || '';
  let selectedSubject = localStorage.getItem('selectedSubject') || 'ICTSheet';
  let sheetData = JSON.parse(localStorage.getItem('sheetData')) || {};

  // Class options corresponding to the selected spreadsheet
  const classOptions = {
    'Sheet1': ['1A1', '1A2', '1A3', '1A4', '1A5'],
    'Sheet2': ['2A1', '2A2', '2A3', '2A4', '2A5', '2A6'],
    'Sheet3': ['3A1', '3A2', '3A3', '3A4', '3A5', '3A6'],
    'Sheet4': ['4A2', '4A3', '4A4', '4A5', '4A6', '4A7'],
    'Sheet5': ['5A1', '5A2', '5A3', '5A4', '5A5', '5A6']
  };

  // Update the class dropdown when the spreadsheet changes
  const updateClassDropdown = (grade) => {
    const classSelect = document.getElementById('classSelect');
    classSelect.innerHTML = '<option value="" disabled selected>Select a class</option>';

    // Create class options based on the selected grade
    const classes = classOptions[grade] || [];
    classes.forEach(className => {
      const option = document.createElement('option');
      option.value = className;
      option.textContent = className;
      classSelect.appendChild(option);
    });

    classSelect.disabled = false; // Enable class selection when a grade is selected
  };


  const fetchDataFromGoogleSheetICT = async (gradeSelect) => {
    const baseUrl = `https://docs.google.com/spreadsheets/d/${googleSheetICTId}/gviz/tq?sheet=${gradeSelect}&tqx=out:json`;
    try {
      const response = await fetch(baseUrl);
      const text = await response.text();
      // Parse JSON from the returned string
      const json = JSON.parse(text.substr(47).slice(0, -2));
      const rows = json.table.rows;

      // Create a sheetData object to store data by category
      sheetData = {
        introduction: [],
        behavior: [],
        classwork: [],
        participation: [],
        improvements: []
      };

      // Iterate through each row and store data into corresponding categories
      rows.forEach((row, index) => {
        if (index > 0) { // Skip the first row, which contains the headers
          sheetData.introduction.push(row.c[0]?.v || '');
          sheetData.behavior.push(row.c[1]?.v || '');
          sheetData.classwork.push(row.c[2]?.v || '');
          sheetData.participation.push(row.c[3]?.v || '');
          sheetData.improvements.push(row.c[4]?.v || '');
        }
      });

      // Save new data to localStorage
      localStorage.setItem('sheetData', JSON.stringify(sheetData));
      return sheetData;
    } catch (error) {
      console.error('Error fetching data from Google Sheet:', error);
    }
  };


  const fetchDataFromGoogleSheetClasses = async (sheetName) => {
    const baseUrl = `https://docs.google.com/spreadsheets/d/${googleSheetClassesId}/gviz/tq?sheet=${sheetName}&tqx=out:json`;
    try {
      const response = await fetch(baseUrl);
      const text = await response.text();
      // Parse JSON from the returned string
      const json = JSON.parse(text.substr(47).slice(0, -2));
      const newData = json.table.rows.map(row => row.c.map(cell => (cell ? cell.v : '')));
      // Save new data to localStorage
      sheetData = newData;
      localStorage.setItem('sheetData', JSON.stringify(sheetData));
      return newData;
    } catch (error) {
      console.error('Error fetching data from Google Sheet:', error);
    }
  };


  // Update table data based on the selected class
  const updateTableData = async (selectedClass) => {

    const tableBody = document.getElementById('commentTable').querySelector('tbody');
    tableBody.innerHTML = ''; // Clear current table content

    // Fetch data from Google Sheets based on the selected class
    const data = await fetchDataFromGoogleSheetClasses(selectedClass);

    // Fetch data from Google Sheets ICT
    const dataGrade = await fetchDataFromGoogleSheetICT(gradeSelected);

    if (data && data.length > 0) {
      // Iterate through each row of data and add it to the table
      data.forEach((rowData, rowIndex) => {
        const newRow = tableBody.insertRow();
        // Add data to corresponding columns
        rowData.slice(0, 10).forEach((cellData, index) => {
          const cell = newRow.insertCell(index);
          cell.textContent = cellData;
        });

        // Call the function to add comment dropdowns for each student's row
        addCommentColumns(newRow, rowData[4], rowData[3], rowIndex);
      });
    }

    // Set up events for the Summary column and enforce dropdown width
    setupSummaryEvents();
    enforceDropdownWidth();
    restoreSavedData();
  };

  // Function to add comment dropdowns for each row in the table
  const addCommentColumns = (newRow, gender, firstName, rowIndex) => {
    const selectedComments = [];

    // Iterate through comment categories and add dropdowns
    ["introduction", "behavior", "classwork", "participation", "improvements"].forEach((category, index) => {
      const cell = newRow.insertCell();
      const container = document.createElement('div');
      container.classList.add('comment-container');

      const select = document.createElement('select');
      select.dataset.category = category;
      select.dataset.rowIndex = rowIndex;
      select.dataset.columnIndex = index;
      populateOptions(select, category); // Populate options into dropdown

      container.appendChild(select);
      cell.appendChild(container);

      selectedComments.push(select);
    });

    // Add Summary column
    const summaryCell = newRow.insertCell();
    summaryCell.classList.add('summary-column');
    const summaryDiv = document.createElement('div');
    summaryDiv.contentEditable = true;
    summaryCell.appendChild(summaryDiv);

    // Set up change events for dropdowns to update summary
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

  // Populate options into dropdown based on the category
  const populateOptions = (select, category) => {
    // Clear previous content in the dropdown
    select.innerHTML = '';

    // Add an empty option to start with nothing selected
    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = '';
    select.appendChild(emptyOption);

    // Retrieve data from sheetData based on the category
    const data = sheetData[category];

    if (data && Array.isArray(data)) {
      // Iterate through each item in the category and add it to the dropdown
      data.forEach(item => {
        if (item) {
          const option = document.createElement('option');
          option.value = item;
          option.textContent = item;
          select.appendChild(option);
        }
      });
    }
  };

  // Update summary content when the user changes the dropdown
  const updateSummary = (firstName, gender, selectedComments, summaryCell) => {
    const pronouns = {
      'M': { subject: 'He', object: 'him', possessive: 'his' },
      'F': { subject: 'She', object: 'her', possessive: 'her' }
    };

    const currentPronoun = pronouns[gender] || pronouns['M'];

    // Create summary content from the selected comment values
    let summaryText = `${firstName} `;

    const introduction = selectedComments[0].value;
    const behavior = `${currentPronoun.subject} ${selectedComments[1].value.toLowerCase()}`;
    const classwork = `${currentPronoun.possessive.charAt(0).toUpperCase() + currentPronoun.possessive.slice(1)} classwork ${selectedComments[2].value.toLowerCase()}`;
    const participation = `${currentPronoun.subject} ${selectedComments[3].value.toLowerCase()}`;
    const improvements = `${currentPronoun.subject} ${selectedComments[4].value.toLowerCase()}`;

    summaryText += `${introduction}. ${behavior}. ${classwork}. ${participation}. ${improvements}.`;

    // Handle formatting and character limit
    summaryText = summaryText.replace(/\s+/g, ' ').replace(/(\.\s*)+/g, '. ').trim();

    // Limit displayed characters
    if (summaryText.length > maxLengthDisplay) {
      summaryCell.dataset.fullText = summaryText;
      summaryText = summaryText.slice(0, maxLengthDisplay) + "...";
    }
    summaryCell.textContent = summaryText;
  };

  // Set up expand/collapse events for the Summary column
  const setupSummaryEvents = () => {
    document.querySelectorAll('.summary-column div').forEach(div => {
      const originalText = div.dataset.fullText || div.textContent.trim();
      if (originalText.length > maxLengthDisplay) {
        div.textContent = originalText.slice(0, maxLengthDisplay) + "...";
        div.dataset.fullText = originalText;
      }

      div.addEventListener('click', function () {
        if (this.classList.contains('expanded')) {
          this.classList.remove('expanded');
          this.style.width = ''; // Reset width to default when collapsed
          this.textContent = this.dataset.fullText.slice(0, maxLengthDisplay) + "...";
        } else {
          this.classList.add('expanded');
          this.style.width = '100%'; // Expand width when opened
          this.textContent = this.dataset.fullText;
        }
      });
    });
  };

  // Enforce dropdown width to fit container when clicked
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

  // Save selected data to localStorage
  const saveData = (rowIndex, columnIndex, value) => {
    let savedData = JSON.parse(localStorage.getItem('savedComments')) || {};
    if (!savedData[rowIndex]) {
      savedData[rowIndex] = {};
    }
    savedData[rowIndex][columnIndex] = value;
    localStorage.setItem('savedComments', JSON.stringify(savedData));
  };

  // Restore saved data from localStorage
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

  // Reset the table and clear saved data
  const resetAll = () => {
    document.getElementById('commentTable').querySelector('tbody').innerHTML = '';
    localStorage.removeItem('savedComments');
  };

  document.getElementById('gradeSelect').addEventListener('change', function () {
    gradeSelected = this.value;
    localStorage.setItem('selectedSheet', gradeSelected);
    // Clear old sheet data in localStorage when selecting a new grade
    localStorage.removeItem('sheetData');
    sheetData = {}; // Reset sheetData to avoid retaining old data
    updateClassDropdown(gradeSelected);
  });

  // Enable "Go" button when a class is selected
  document.getElementById('classSelect').addEventListener('change', function () {
    document.getElementById('goBtn').disabled = false;
  });

  // Event when pressing "Go" button to update table data
  document.getElementById('goBtn').addEventListener('click', async function () {
    const selectedClass = document.getElementById('classSelect').value;
    await updateTableData(selectedClass);
    populateDropdowns();
  });

  // Event when pressing "Save" button to notify saving data
  document.getElementById('saveBtn').addEventListener('click', function () {
    // Data is saved automatically when dropdown changes
    alert("Data saved successfully!");
  });

  // Event when pressing "Create New" button to reset the table
  document.getElementById('createNewBtn').addEventListener('click', function () {
    resetAll();
  });

  // Set up summary events for the Summary column
  setupSummaryEvents();
});

// Function to download selected table columns as an Excel file
const downloadExcel = () => {
  // Reference the table element in the HTML
  const table = document.getElementById('commentTable');

  // Create a new workbook and add a worksheet
  const workbook = XLSX.utils.book_new();
  const worksheetData = [];

  // Add headers to the worksheet (only the specified columns)
  const headers = [
    "No", "Student Code", "Student Name", 
    "Participation (15)", "Classwork (15)", "Assessment 1 (100)",
    "Assessment 2 (100)", "Final Exam (100)", "Summary"
  ];
  worksheetData.push(headers);

  // Iterate over table rows to get the cell values
  const rows = table.querySelectorAll('tbody tr');
  rows.forEach((row) => {
    const rowData = [];
    const cells = row.querySelectorAll('td');

    // Add specific cell values based on the column indices of interest
    // Indices correspond to: No (0), Student Code (1), Student Name (2),
    // Participation (3), Classwork (4), Assessment 1 (5), Assessment 2 (6),
    // Final Exam (7), and Summary (8)
    [0, 1, 2, 5, 6, 7 , 8, 9, 15].forEach((index) => {
      if (index === 14) {
        // For the Summary column (index 10), extract content from the inner div
        const summaryDiv = cells[index]?.querySelector('div');
        rowData.push(summaryDiv ? summaryDiv.textContent.trim() : '');
      } else {
        rowData.push(cells[index]?.textContent.trim() || '');
      }
    });

    // Add row data to the worksheet data
    worksheetData.push(rowData);
  });

  // Create a worksheet with the extracted data
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

  // Add the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, "Student Comments");

  // Create a binary string of the workbook and prompt download
  XLSX.writeFile(workbook, "Student_Comments.xlsx");
};

// Add event listener to the "Download" button
document.getElementById('downloadBtn').addEventListener('click', downloadExcel);