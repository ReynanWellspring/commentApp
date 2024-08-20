document.addEventListener('DOMContentLoaded', function () {
  const sheetUrl = 'https://script.google.com/macros/s/AKfycbxyXv5Fz0GO2Shu4ZNM-C4701DZE7iYRWqH0pV5GY90F6FQ9l186URfGfwF_BKXhAV9/exec'; // Replace with your Google Apps Script Web App URL

  let selectedSheet = 'Sheet1'; // Default sheet
  let sheetData = {}; // To store the fetched data

  // Function to fetch data from the Google Apps Script
  const fetchData = (sheet) => {
    fetch(sheetUrl + '?sheet=' + sheet)
      .then(response => response.json())
      .then(data => {
        sheetData = data;
      })
      .catch(error => console.error('Error fetching data:', error));
  };

  // Fetch data for the default sheet initially
  fetchData(selectedSheet);

  // Update selected sheet and fetch data when grade changes
  document.getElementById('gradeSelect').addEventListener('change', function () {
    selectedSheet = this.value;
    fetchData(selectedSheet); // Fetch new data when grade changes
  });

  // Event listener for the Add Names button
  document.getElementById('addNameBtn').addEventListener('click', function() {
    const nameInput = document.getElementById('nameInput');
    const names = nameInput.value.trim().split("\n").map(name => name.trim()).filter(name => name !== "");

    if (names.length > 0) {
      const tableBody = document.getElementById('commentTable').querySelector('tbody');

      // Loop through each name and add it to the table
      names.slice(0, 25).forEach(name => {
        const newRow = tableBody.insertRow();
        
        // Add Name cell
        const nameCell = newRow.insertCell();
        nameCell.textContent = name;

        // Add Gender cell with input field
        const genderCell = newRow.insertCell();
        const genderInput = document.createElement('input');
        genderInput.type = 'text';
        genderInput.maxLength = 1; // Limit to 1 character
        genderCell.appendChild(genderInput);

        const selectedComments = [];

        // List of categories to be used for dropdowns
        ["introduction", "behavior", "classwork", "participation", "improvements"].forEach(category => {
          const cell = newRow.insertCell();
          const select = document.createElement('select');
          
          // Populate each dropdown with comments fetched from the data
          sheetData[category].forEach(comment => {
            const option = document.createElement('option');
            option.value = comment;
            option.textContent = comment;
            select.appendChild(option);
          });

          cell.appendChild(select);
          selectedComments.push(select);
        });

        // Add an input field for additional comments
        const additionalCommentsCell = newRow.insertCell();
        const input = document.createElement('input');
        input.type = 'text';
        additionalCommentsCell.appendChild(input);
        selectedComments.push(input);

        // Add a new cell for the summary
        const summaryCell = newRow.insertCell();

        // Function to clean up the summary by fixing encoding, removing extra punctuation, and ensuring correct spacing
        const cleanUpSummary = (summary) => {
          // Decode HTML entities like &#039;
          let decodedSummary = summary.replace(/&#039;/g, "'");

          // Remove commas following periods, and ensure single space after periods
          decodedSummary = decodedSummary.replace(/\. *,/g, '. ').replace(/\s+/g, ' ');

          // Remove any leading, trailing, or multiple consecutive commas
          return decodedSummary.replace(/,+/g, ',').replace(/,\s*$/, '').trim();
        };

        // Function to replace gender pronouns based on the gender input
        const applyGenderPronouns = (text, gender) => {
          if (gender === 'f') {
            return text.replace(/\bHe\b/g, 'She').replace(/\bhis\b/g, 'her').replace(/\bhim\b/g, 'her');
          } else if (gender === 'm') {
            return text.replace(/\bShe\b/g, 'He').replace(/\bher\b/g, 'his').replace(/\bher\b/g, 'him');
          }
          return text; // No replacement if gender is not specified
        };

        // Function to update the summary whenever any input changes
        const updateSummary = () => {
          let gender = genderInput.value.trim().toLowerCase();
          let fullSummary = `${name} ` + selectedComments.slice(0, 5).map((e) => {
            let comment = e.value || e.textContent;
            return applyGenderPronouns(comment, gender);
          }).join(', ');

          // Append additional comments if provided
          const additionalComments = selectedComments[5].value.trim();
          if (additionalComments) {
            fullSummary += `. ${applyGenderPronouns(additionalComments, gender)}`;
          }

          fullSummary = cleanUpSummary(fullSummary); // Clean up the summary
          let displaySummary = fullSummary;

          if (fullSummary.length > 100) {
            displaySummary = fullSummary.substring(0, 97) + '...';
          }

          summaryCell.textContent = displaySummary;
          summaryCell.dataset.fullSummary = fullSummary; // Store full summary in data attribute
        };

        // Attach the event listener to each dropdown and the additional comments field to update the summary when a selection is made or input is provided
        selectedComments.forEach(element => {
          element.addEventListener('change', function() {
            updateSummary();
          });
          // For text input, also listen to input events
          if (element.tagName === 'INPUT') {
            element.addEventListener('input', function() {
              updateSummary();
            });
          }
        });

        // Attach the event listener to the gender input to update the summary when gender is entered
        genderInput.addEventListener('input', function() {
          updateSummary();
        });

        // Initialize the summary (optional: start with an empty summary)
        summaryCell.textContent = ''; // Clear the summary initially
      });

      // Clear the name input field after processing
      nameInput.value = "";
    }
  });

  // Function to download the summary data as Excel
  document.getElementById('downloadBtn').addEventListener('click', function() {
    const table = document.getElementById('commentTable');
    const summaries = [];

    // Collect summary data from the table
    for (let i = 1, row; row = table.rows[i]; i++) {
      const name = row.cells[0].textContent;
      const summary = row.cells[8].dataset.fullSummary || row.cells[8].textContent; // Use full summary for export
      if (summary) {
        summaries.push([name, summary]);
      }
    }

    // Convert summaries into a worksheet
    const worksheet = XLSX.utils.aoa_to_sheet([['Name', 'Summary'], ...summaries]);

    // Create a new workbook and add the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Summaries');

    // Export the workbook as an Excel file
    XLSX.writeFile(workbook, 'Teacher_Comments_Summary.xlsx');
  });
});
