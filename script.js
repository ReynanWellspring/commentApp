document.addEventListener('DOMContentLoaded', function () {
  const sheetUrl = 'https://script.google.com/macros/s/AKfycby-kdyJOCUk_I1kL_oT37hmBA8-CVtcUtWa2bFMh--88HQQs1I91qC0w0v0WnrjOymg/exec'; // Replace with your Google Apps Script Web App URL

  // Fetch data from the Google Apps Script
  fetch(sheetUrl)
    .then(response => response.json())
    .then(data => {
      document.getElementById('addNameBtn').addEventListener('click', function() {
        const nameInput = document.getElementById('nameInput');
        const names = nameInput.value.trim().split("\n").map(name => name.trim()).filter(name => name !== "");

        if (names.length > 0) {
          const tableBody = document.getElementById('commentTable').querySelector('tbody');

          // Loop through each name and add it to the table
          names.slice(0, 25).forEach(name => {
            const newRow = tableBody.insertRow();
            const nameCell = newRow.insertCell();
            nameCell.textContent = name;

            const selectedComments = [];

            // List of categories to be used for dropdowns
            ["introduction", "behavior", "classwork", "participation", "improvements"].forEach(category => {
              const cell = newRow.insertCell();
              const select = document.createElement('select');
              
              // Populate each dropdown with comments fetched from the data
              data[category].forEach(comment => {
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

            // Function to update the summary whenever any input changes
            const updateSummary = () => {
              let summary = `${name}: ` + selectedComments.map(e => e.value || e.textContent).join(', ');
              if (summary.length > 100) {
                summary = summary.substring(0, 97) + '...';
              }
              summaryCell.textContent = summary;
            };

            // Attach the event listener to each element to update the summary on change
            selectedComments.forEach(element => {
              element.addEventListener('change', updateSummary);
            });

            // Initialize the summary
            updateSummary();
          });

          // Clear the name input field after processing
          nameInput.value = "";
        }
      });
    })
    .catch(error => console.error('Error fetching data:', error));
});
