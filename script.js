document.addEventListener('DOMContentLoaded', function () {
  const sheetUrls = {
    'EnglishSheet': 'https://script.google.com/macros/s/AKfycbyXvhZFuFL1gllrdOjgnNfnqFfIkGiDAotHob2pZeFvU1dfWW0zgNzRLmCMDUekXVUuFA/exec',
    'MathSheet': 'https://script.google.com/macros/s/AKfycby3yb7vXEqdqNLU9z3xUeybkYqL2cygrsDUvtmKi5t0Defxf-BR-JeBJn4BPMVJDPYo/exec',
    'ScienceSheet': 'https://script.google.com/macros/s/AKfycbyfY-ZB1F-9dcChFe4vWKC23x78O9LpBSEbKtuCCLTEkpFYk8bWDW9KdbZ5jKU9yF-Riw/exec',
    'ICTSheet': 'https://script.google.com/macros/s/AKfycbzaMRyBSWO8e1pRcSryCLMNyqCI9fQhfpqu1m14XTCBZRnx-wDPcdyljHPdfNZ5iNGg/exec',
  };

  let selectedSheet = localStorage.getItem('selectedSheet') || 'Sheet1'; // Default sheet
  let selectedSubject = localStorage.getItem('selectedSubject') || 'ICTSheet'; // Default subject sheet
  let sheetData = JSON.parse(localStorage.getItem('sheetData')) || {}; // To store the fetched data

  // Class options based on grade selection
  const classOptions = {
    'Sheet1': ['1A1', '1A2', '1A3', '1A4', '1A5'],
    'Sheet2': ['2A1', '2A2', '2A3', '2A4', '2A5', '2A6'],
    'Sheet3': ['3A1', '3A2', '3A3', '3A4', '3A5', '3A6'],
    'Sheet4': ['4A1', '4A2', '4A3', '4A4', '4A5', '4A6', '4A7'],
    'Sheet5': ['5A1', '5A2', '5A3', '5A4', '5A5', '5A6']
  };

  // Function to fetch data from the Google Apps Script
  const fetchData = () => {
    const url = sheetUrls[selectedSubject] + '?sheet=' + selectedSheet;
    fetch(url)
      .then(response => response.json())
      .then(data => {
        sheetData = data;
        localStorage.setItem('sheetData', JSON.stringify(sheetData)); // Save data to localStorage
        populateDropdowns(); // Populate dropdowns after data is fetched
      })
      .catch(error => console.error('Error fetching data:', error));
  };

  // Function to update the class dropdown based on the selected grade
  const updateClassDropdown = (grade) => {
    const classSelect = document.getElementById('classSelect');
    classSelect.innerHTML = ''; // Clear existing options

    const classes = classOptions[grade] || [];
    classes.forEach(className => {
      const option = document.createElement('option');
      option.value = className;
      option.textContent = className;
      classSelect.appendChild(option);
    });
    localStorage.setItem('selectedClass', classSelect.value); // Save selected class to localStorage
  };

  // Populate dropdowns with data from Google Sheet
  const populateDropdowns = () => {
    ["introduction", "behavior", "classwork", "participation", "improvements"].forEach(category => {
      const data = sheetData[category];
      if (data && Array.isArray(data)) {
        const dropdowns = document.querySelectorAll(`select[data-category="${category}"]`);
        dropdowns.forEach(dropdown => {
          dropdown.innerHTML = ''; // Clear existing options
          data.forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            dropdown.appendChild(option);
          });
        });
      }
    });
  };

  // Function to replace gender pronouns based on the gender input
  const applyGenderPronouns = (text, gender) => {
    if (gender === 'f') {
      return text
        .replace(/\bHe\b/g, 'She')
        .replace(/\bhe\b/g, 'she')
        .replace(/\bHis\b/g, 'Her')
        .replace(/\bhis\b/g, 'her')
        .replace(/\bHim\b/g, 'Her')
        .replace(/\bhim\b/g, 'her');
    } else if (gender === 'm') {
      return text
        .replace(/\bShe\b/g, 'He')
        .replace(/\bshe\b/g, 'he')
        .replace(/\bHer\b/g, 'His')
        .replace(/\bher\b/g, 'his')
        .replace(/\bHim\b/g, 'Him')
        .replace(/\bhim\b/g, 'him');
    }
    return text; // No replacement if gender is not specified
  };

  // Function to clean up the summary by fixing encoding, removing extra punctuation, ensuring correct spacing, and correcting grammar
  const cleanUpSummary = (summary) => {
    let decodedSummary = summary.replace(/&#039;/g, "'");

    decodedSummary = decodedSummary.replace(/\. *,/g, '. ').replace(/\s+/g, ' ').replace(/,+/g, ',').replace(/,\s*$/, '').trim();

    decodedSummary = decodedSummary.replace(/(^\w|\.\s*\w)/g, letter => letter.toUpperCase());

    return decodedSummary;
  };

  // Function to update the summary whenever any input changes
  const updateSummary = (name, genderInput, selectedComments, summaryCell) => {
    let gender = genderInput.value.trim().toLowerCase();
    let summaryGenerated = false;

    let fullSummary = `${name} ` + selectedComments.slice(0, 5).map((e, i) => {
      let comment = e.value; // Only use selected value, not the textContent
      if (comment) {
        summaryGenerated = true;
        if (i >= 1 && i <= 4) {
          return applyGenderPronouns(comment, gender);
        }
        return comment;
      }
      return ''; // Return empty if no comment selected
    }).filter(comment => comment).join(', ');

    const additionalComments = selectedComments[5].value.trim();
    if (additionalComments) {
      fullSummary += `. ${applyGenderPronouns(additionalComments, gender)}`;
    }

    fullSummary = cleanUpSummary(fullSummary);

    if (summaryGenerated) {
      summaryCell.textContent = fullSummary;
    } else {
      summaryCell.textContent = ''; // Keep it empty initially
    }

    summaryCell.dataset.fullSummary = fullSummary; // Store full summary in data attribute

    saveTableData(); // Save the current table data to localStorage
  };

  // Function to save the table data to localStorage
  const saveTableData = () => {
    const table = document.getElementById('commentTable');
    const tableData = [];

    for (let i = 1, row; row = table.rows[i]; i++) {
      const rowData = {
        name: row.cells[0].textContent,
        gender: row.cells[1].children[0].value,
        comments: Array.from(row.cells).slice(2, 7).map(cell => cell.children[0].value),
        additionalComments: row.cells[7].children[0].value,
        summary: row.cells[8].textContent
      };
      tableData.push(rowData);
    }

    localStorage.setItem('tableData', JSON.stringify(tableData));
  };

  // Function to load the table data from localStorage
  const loadTableData = () => {
    const tableData = JSON.parse(localStorage.getItem('tableData'));

    if (tableData && tableData.length > 0) {
      const tableBody = document.getElementById('commentTable').querySelector('tbody');
      tableBody.innerHTML = ''; // Clear any existing rows

      tableData.forEach(rowData => {
        const newRow = tableBody.insertRow();

        const nameCell = newRow.insertCell();
        nameCell.textContent = rowData.name;

        const genderCell = newRow.insertCell();
        const genderInput = document.createElement('input');
        genderInput.type = 'text';
        genderInput.maxLength = 1;
        genderInput.value = rowData.gender;
        genderCell.appendChild(genderInput);

        const selectedComments = [];

        ["introduction", "behavior", "classwork", "participation", "improvements"].forEach((category, index) => {
          const cell = newRow.insertCell();
          const select = document.createElement('select');
          select.dataset.category = category;

          if (sheetData[category]) {
            sheetData[category].forEach(comment => {
              const option = document.createElement('option');
              option.value = comment;
              option.textContent = comment;
              if (comment === rowData.comments[index]) {
                option.selected = true;
              }
              select.appendChild(option);
            });
          }

          cell.appendChild(select);
          selectedComments.push(select);
        });

        const additionalCommentsCell = newRow.insertCell();
        const input = document.createElement('input');
        input.type = 'text';
        input.value = rowData.additionalComments;
        additionalCommentsCell.appendChild(input);
        selectedComments.push(input);

        const summaryCell = newRow.insertCell();
        summaryCell.textContent = rowData.summary;
        summaryCell.dataset.fullSummary = rowData.summary;

        selectedComments.forEach(element => {
          element.addEventListener('change', function() {
            updateSummary(rowData.name, genderInput, selectedComments, summaryCell);
          });
          if (element.tagName === 'INPUT') {
            element.addEventListener('input', function() {
              updateSummary(rowData.name, genderInput, selectedComments, summaryCell);
            });
          }
        });

        genderInput.addEventListener('input', function() {
          if (selectedComments.some(select => select.value)) {
            updateSummary(rowData.name, genderInput, selectedComments, summaryCell);
          }
        });
      });
    }
  };

  // Fetch data for the default sheet initially
  if (!sheetData || Object.keys(sheetData).length === 0) {
    fetchData();
  } else {
    populateDropdowns();
  }

  updateClassDropdown(selectedSheet); // Update class dropdown for default sheet

  // Load the previously saved table data
  loadTableData();

  // Update selected sheet and fetch data when grade changes
  document.getElementById('gradeSelect').addEventListener('change', function () {
    selectedSheet = this.value;
    localStorage.setItem('selectedSheet', selectedSheet); // Save selected sheet to localStorage
    fetchData();
    updateClassDropdown(selectedSheet);
  });

  // Update selected subject and fetch data when subject changes
  document.getElementById('subjectSelect').addEventListener('change', function () {
    selectedSubject = this.value;
    localStorage.setItem('selectedSubject', selectedSubject); // Save selected subject to localStorage
    fetchData();
  });

  // Event listener for the Add Names button
  document.getElementById('addNameBtn').addEventListener('click', function() {
    const nameInput = document.getElementById('nameInput');
    const names = nameInput.value.trim().split("\n").map(name => name.trim()).filter(name => name !== "");

    if (names.length > 0) {
      const tableBody = document.getElementById('commentTable').querySelector('tbody');

      names.slice(0, 25).forEach(name => {
        const newRow = tableBody.insertRow();
        
        const nameCell = newRow.insertCell();
        nameCell.textContent = name;

        const genderCell = newRow.insertCell();
        const genderInput = document.createElement('input');
        genderInput.type = 'text';
        genderInput.maxLength = 1;
        genderCell.appendChild(genderInput);

        const selectedComments = [];

        ["introduction", "behavior", "classwork", "participation", "improvements"].forEach(category => {
          const cell = newRow.insertCell();
          const select = document.createElement('select');
          select.dataset.category = category;
          
          if (sheetData[category]) {
            sheetData[category].forEach(comment => {
              const option = document.createElement('option');
              option.value = comment;
              option.textContent = comment;
              select.appendChild(option);
            });
          }

          cell.appendChild(select);
          selectedComments.push(select);
        });

        const additionalCommentsCell = newRow.insertCell();
        const input = document.createElement('input');
        input.type = 'text';
        additionalCommentsCell.appendChild(input);
        selectedComments.push(input);

        const summaryCell = newRow.insertCell();

        selectedComments.forEach(element => {
          element.addEventListener('change', function() {
            updateSummary(name, genderInput, selectedComments, summaryCell);
          });
          if (element.tagName === 'INPUT') {
            element.addEventListener('input', function() {
              updateSummary(name, genderInput, selectedComments, summaryCell);
            });
          }
        });

        genderInput.addEventListener('input', function() {
          if (selectedComments.some(select => select.value)) {
            updateSummary(name, genderInput, selectedComments, summaryCell);
          }
        });

        summaryCell.textContent = '';
      });

      nameInput.value = "";
      saveTableData(); // Save the updated table data
    }
  });

  // Event listener for the Download Summary button
  document.getElementById('downloadBtn').addEventListener('click', function() {
    const table = document.getElementById('commentTable');
    const summaries = [];

    // Get selected Grade, Class, and Subject
    const selectedGrade = document.getElementById('gradeSelect').options[document.getElementById('gradeSelect').selectedIndex].text;
    const selectedClass = document.getElementById('classSelect').options[document.getElementById('classSelect').selectedIndex].text;
    const selectedSubject = document.getElementById('subjectSelect').options[document.getElementById('subjectSelect').selectedIndex].text;

    // Add Grade, Class, and Subject to the top of the Excel sheet
    summaries.push(['Grade:', selectedGrade]);
    summaries.push(['Class:', selectedClass]);
    summaries.push(['Subject:', selectedSubject]);
    summaries.push([]); // Add an empty row for spacing
    summaries.push(['Name', 'Summary']); // Headers for the summary data

    // Loop through the table rows to collect summaries
    for (let i = 1, row; row = table.rows[i]; i++) {
      const name = row.cells[0].textContent;
      const summary = row.cells[8].dataset.fullSummary || row.cells[8].textContent;
      if (summary) {
        summaries.push([name, summary]);
      }
    }

    // Create the worksheet with the collected data
    const worksheet = XLSX.utils.aoa_to_sheet(summaries);

    // Create a new workbook and append the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summaries');

    // Write the workbook to a file
    XLSX.writeFile(workbook, 'Teacher_Comments_Summary.xlsx');
  });

  // Event listener for the Create New button
  document.getElementById('createNewBtn').addEventListener('click', function() {
    // Clear localStorage
    localStorage.removeItem('selectedSheet');
    localStorage.removeItem('selectedSubject');
    localStorage.removeItem('sheetData');
    localStorage.removeItem('tableData');

    // Reload the page to reset everything
    location.reload();
  });
});
