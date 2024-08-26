document.addEventListener('DOMContentLoaded', function () {
  const sheetUrls = {
    'EnglishSheet': 'https://script.google.com/macros/s/AKfycbyXvhZFuFL1gllrdOjgnNfnqFfIkGiDAotHob2pZeFvU1dfWW0zgNzRLmCMDUekXVUuFA/exec',
    'MathSheet': 'https://script.google.com/macros/s/AKfycby3yb7vXEqdqNLU9z3xUeybkYqL2cygrsDUvtmKi5t0Defxf-BR-JeBJn4BPMVJDPYo/exec',
    'ScienceSheet': 'https://script.google.com/macros/s/AKfycbyfY-ZB1F-9dcChFe4vWKC23x78O9LpBSEbKtuCCLTEkpFYk8bWDW9KdbZ5jKU9yF-Riw/exec',
    'ICTSheet': 'https://script.google.com/macros/s/AKfycbzaMRyBSWO8e1pRcSryCLMNyqCI9fQhfpqu1m14XTCBZRnx-wDPcdyljHPdfNZ5iNGg/exec',
  };

  const classDataUrl = 'https://script.google.com/macros/s/AKfycbyQf9QiwTGFSsKIhNhOMXB15FJyBTO_r_ppkRl6vGNqerIt_rId-_ji6ngnDNfI_FqPTg/exec';

  let selectedSheet = localStorage.getItem('selectedSheet') || 'Sheet1'; 
  let selectedSubject = localStorage.getItem('selectedSubject') || 'ICTSheet';
  let sheetData = JSON.parse(localStorage.getItem('sheetData')) || {};

  const classOptions = {
    'Sheet1': ['','1A1', '1A2', '1A3', '1A4', '1A5'],  
    'Sheet2': ['','2A1', '2A2', '2A3', '2A4', '2A5', '2A6'],  
    'Sheet3': ['','3A1', '3A2', '3A3', '3A4', '3A5', '3A6'],  
    'Sheet4': ['','4A2', '4A3', '4A4', '4A5', '4A6', '4A7'],  
    'Sheet5': ['','5A1', '5A2', '5A3', '5A4', '5A5', '5A6']   
  };

  const fetchClassData = (className) => {
    const url = `${classDataUrl}?sheet=${className}`;
    return fetch(url)
      .then(response => response.json())
      .catch(error => console.error('Error fetching class data:', error));
  };

  const fetchData = () => {
    const url = sheetUrls[selectedSubject] + '?sheet=' + selectedSheet;
    fetch(url)
      .then(response => response.json())
      .then(data => {
        sheetData = data;
        localStorage.setItem('sheetData', JSON.stringify(sheetData)); 
        populateDropdowns();
      })
      .catch(error => console.error('Error fetching data:', error));
  };

  const updateClassDropdown = (grade) => {
    const classSelect = document.getElementById('classSelect');
    classSelect.innerHTML = '';

    const classes = classOptions[grade] || [];
    classes.forEach(className => {
      const option = document.createElement('option');
      option.value = className;
      option.textContent = className;
      classSelect.appendChild(option);
    });

    classSelect.disabled = false;
    document.getElementById('goBtn').disabled = false; 
    localStorage.setItem('selectedClass', classSelect.value);
  };

  const populateDropdowns = () => {
    ["introduction", "behavior", "classwork", "participation", "improvements"].forEach(category => {
      const data = sheetData[category];
      if (data && Array.isArray(data)) {
        const dropdowns = document.querySelectorAll(`select[data-category="${category}"]`);
        dropdowns.forEach(dropdown => {
          dropdown.innerHTML = ''; 
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

  const applyGenderPronouns = (text, gender) => {
    let modifiedText = text.trim();

    if (gender === 'f') {
        // Check for male pronouns and correct them to female pronouns
        modifiedText = modifiedText.replace(/\bhe\b/gi, 'she')
                                   .replace(/\bhis\b/gi, 'her')
                                   .replace(/\bhim\b/gi, 'her')
                                   .replace(/\bhe's\b/gi, 'she\'s')
                                   .replace(/\bhimself\b/gi, 'herself');
        // Ensure the paragraph starts with "She" if it doesn't already
        if (!/^(she|her)/i.test(modifiedText)) {
            modifiedText = `She ${modifiedText.charAt(0).toLowerCase()}${modifiedText.slice(1)}`;
        }
    } else if (gender === 'm') {
        // Check for female pronouns and correct them to male pronouns
        modifiedText = modifiedText.replace(/\bshe\b/gi, 'he')
                                   .replace(/\bher\b/gi, 'his')
                                   .replace(/\bhers\b/gi, 'his')
                                   .replace(/\bshe's\b/gi, 'he\'s')
                                   .replace(/\bherself\b/gi, 'himself');
        // Ensure the paragraph starts with "He" if it doesn't already
        if (!/^(he|his)/i.test(modifiedText)) {
            modifiedText = `He ${modifiedText.charAt(0).toLowerCase()}${modifiedText.slice(1)}`;
        }
    }

    return modifiedText;
  };

  const cleanUpSummary = (summary) => {
    let decodedSummary = summary.replace(/&#039;/g, "'");
    decodedSummary = decodedSummary.replace(/\. *,/g, '. ').replace(/\s+/g, ' ').replace(/,+/g, ',').replace(/,\s*$/, '').trim();
    decodedSummary = decodedSummary.replace(/(^\w|\.\s*\w)/g, letter => letter.toUpperCase());
    return decodedSummary;
  };

  const updateSummary = (name, genderInput, selectedComments, summaryCell) => {
    let gender = genderInput.value.trim().toLowerCase();
    let summaryGenerated = false;

    let fullSummary = `${name} ` + selectedComments.slice(0, 5).map((e, i) => {
      let comment = e.value; 
      if (comment) {
        summaryGenerated = true;
        if (i >= 1 && i <= 4) {  
          return applyGenderPronouns(comment, gender);
        }
        return comment;
      }
      return ''; 
    }).filter(comment => comment).join(', ');

    const additionalComments = selectedComments[5].value.trim();
    if (additionalComments) {
      fullSummary += `. ${applyGenderPronouns(additionalComments, gender)}`;
    }

    fullSummary = cleanUpSummary(fullSummary);

    if (summaryGenerated) {
      summaryCell.textContent = fullSummary;
    } else {
      summaryCell.textContent = ''; 
    }

    summaryCell.dataset.fullSummary = fullSummary;

    saveTableData();
  };

  const saveTableData = () => {
    const table = document.getElementById('commentTable');
    const tableData = [];

    for (let i = 1, row; row = table.rows[i]; i++) {
      const rowData = {
        name: row.cells[0].textContent,
        gender: row.cells[1].children[0].value,
        comments: Array.from(row.cells).slice(2, 7).map(cell => cell.children[0].value),
        additionalComments: row.cells[7].children[0].value,
        summary: row.cells[8].querySelector('div').textContent
      };
      tableData.push(rowData);
    }

    localStorage.setItem('tableData', JSON.stringify(tableData));
  };

  const loadTableData = () => {
    const tableData = JSON.parse(localStorage.getItem('tableData'));

    if (tableData && tableData.length > 0) {
      const tableBody = document.getElementById('commentTable').querySelector('tbody');
      tableBody.innerHTML = ''; 

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
        const summaryDiv = document.createElement('div');
        summaryDiv.contentEditable = true;
        summaryDiv.textContent = rowData.summary;
        summaryCell.appendChild(summaryDiv);

        selectedComments.forEach(element => {
          element.addEventListener('change', function() {
            updateSummary(rowData.name, genderInput, selectedComments, summaryDiv);
          });
          if (element.tagName === 'INPUT') {
            element.addEventListener('input', function() {
              updateSummary(rowData.name, genderInput, selectedComments, summaryDiv);
            });
          }
        });

        genderInput.addEventListener('input', function() {
          if (selectedComments.some(select => select.value)) {
            updateSummary(rowData.name, genderInput, selectedComments, summaryDiv);
          }
        });

        summaryDiv.addEventListener('input', saveTableData);
      });
    }
  };

  if (!sheetData || Object.keys(sheetData).length === 0) {
    fetchData();
  } else {
    populateDropdowns();
  }

  const classSelect = document.getElementById('classSelect');
  classSelect.disabled = true;
  document.getElementById('goBtn').disabled = true;

  updateClassDropdown(selectedSheet);

  loadTableData();

  document.getElementById('gradeSelect').addEventListener('change', function () {
    selectedSheet = this.value;
    localStorage.setItem('selectedSheet', selectedSheet);
    updateClassDropdown(selectedSheet);
  });

  document.getElementById('subjectSelect').addEventListener('change', function () {
    selectedSubject = this.value;
    localStorage.setItem('selectedSubject', selectedSubject);
  });

  document.getElementById('goBtn').addEventListener('click', function() {
    const selectedClass = document.getElementById('classSelect').value;

    fetchClassData(selectedClass).then(data => {
      const tableBody = document.getElementById('commentTable').querySelector('tbody');
      tableBody.innerHTML = ''; 

      if (data && Array.isArray(data)) {
        data.forEach(item => {
          const newRow = tableBody.insertRow();

          const nameCell = newRow.insertCell();
          nameCell.textContent = item.Name;

          const genderCell = newRow.insertCell();
          const genderInput = document.createElement('input');
          genderInput.type = 'text';
          genderInput.maxLength = 1;
          genderInput.value = item.Gender;
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
          additionalCommentsCell.appendChild(input);
          selectedComments.push(input);

          const summaryCell = newRow.insertCell();
          const summaryDiv = document.createElement('div');
          summaryDiv.contentEditable = true;
          summaryCell.appendChild(summaryDiv);

          selectedComments.forEach(element => {
            element.addEventListener('change', function() {
              updateSummary(item.Name, genderInput, selectedComments, summaryDiv);
            });
            if (element.tagName === 'INPUT') {
              element.addEventListener('input', function() {
                updateSummary(item.Name, genderInput, selectedComments, summaryDiv);
              });
            }
          });

          genderInput.addEventListener('input', function() {
            if (selectedComments.some(select => select.value)) {
              updateSummary(item.Name, genderInput, selectedComments, summaryDiv);
            }
          });

          summaryDiv.addEventListener('input', saveTableData);
        });
      }
    });

    fetchData();
  });

  document.getElementById('downloadBtn').addEventListener('click', function() {
    const table = document.getElementById('commentTable');
    const summaries = [];

    const selectedGrade = document.getElementById('gradeSelect').options[document.getElementById('gradeSelect').selectedIndex].text;
    const selectedClass = document.getElementById('classSelect').options[document.getElementById('classSelect').selectedIndex].text;
    const selectedSubject = document.getElementById('subjectSelect').options[document.getElementById('subjectSelect').selectedIndex].text;

    summaries.push(['Grade:', selectedGrade]);
    summaries.push(['Class:', selectedClass]);
    summaries.push(['Subject:', selectedSubject]);
    summaries.push([]); 

    summaries.push(['Name', 'Introduction', 'Behavior', 'Classwork', 'Participation', 'Improvements', 'Additional Comments', 'Summary']);

    for (let i = 1, row; row = table.rows[i]; i++) {
      const name = row.cells[0].textContent;
      const introduction = row.cells[2].querySelector('select')?.value || ''; 
      const behavior = row.cells[3].querySelector('select')?.value || ''; 
      const classwork = row.cells[4].querySelector('select')?.value || ''; 
      const participation = row.cells[5].querySelector('select')?.value || ''; 
      const improvements = row.cells[6].querySelector('select')?.value || ''; 
      const additionalComments = row.cells[7].querySelector('input')?.value || ''; 
      const summary = row.cells[8].querySelector('div').textContent || ''; 
      
      const gender = row.cells[1].querySelector('input')?.value.trim().toLowerCase();
      const correctedBehavior = applyGenderPronouns(behavior, gender);
      const correctedClasswork = applyGenderPronouns(classwork, gender);
      const correctedParticipation = applyGenderPronouns(participation, gender);
      const correctedImprovements = applyGenderPronouns(improvements, gender);

      summaries.push([name, introduction, correctedBehavior, correctedClasswork, correctedParticipation, correctedImprovements, additionalComments, summary]);
    }

    const worksheet = XLSX.utils.aoa_to_sheet(summaries);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summaries');
    XLSX.writeFile(workbook, 'Teacher_Comments_Summary.xlsx');
  });

  document.getElementById('createNewBtn').addEventListener('click', function() {
    localStorage.removeItem('selectedSheet');
    localStorage.removeItem('selectedSubject');
    localStorage.removeItem('sheetData');
    localStorage.removeItem('tableData');
    location.reload();
  });
});
