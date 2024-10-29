const maxLengthDisplay = 200; // Giới hạn số ký tự hiển thị là 200

//    Định nghĩa các URL của Google Apps Script để lấy dữ liệu cho từng môn học
//   'EnglishSheet': 'https://script.google.com/macros/s/AKfycbxVb6AfquBhjNOMXYRy2914DZu_08x-u1-4hbKnYe49HrUYDxkPNb7ltUxx3IAdgDdRiw/exec',
//   'MathSheet': 'https://script.google.com/macros/s/AKfycbyQoaYAruRHG8mJvpygzOAUSdWp9bnZAZ7958uFifW6GhlBZ4K_VIDMhJ2IMxsKAgxR/exec',
//   'ScienceSheet': 'https://script.google.com/macros/s/AKfycbyhyWjEqnbXDfbXCMX5U4saJ8qSrpoqUk4eixDC7dboc2dF7PZbg0X91-ewbo0vHZZPKw/exec',
//   'ICTSheet': 'https://script.google.com/macros/s/AKfycbza8qM9Brr9vl8qpeomrApgc8XbsF7qUu0inqUV3HQza30F45lpNNmKrDfbiFiGQSeT/exec',

document.addEventListener('DOMContentLoaded', function () {


  const googleSheetClassesId = '1GgdYi_m8trTQnkL0kp4P2xb_PBuqinmsvn1hZN-G1pM'; // ID của Google Sheet Classes
  const googleSheetICTId = '1utKoEwDu208dlHLs0-Pfa_m4Nm06dTkpoqhqeSn2YNE'; // ID của Google Sheet ICT

  // Lưu trữ thông tin trang tính và môn học đã chọn trong localStorage để dễ dàng khôi phục sau
  let gradeSelected = localStorage.getItem('selectedSheet') || '';
  let selectedSubject = localStorage.getItem('selectedSubject') || 'ICTSheet';
  let sheetData = JSON.parse(localStorage.getItem('sheetData')) || {};

  // Các lựa chọn lớp học tương ứng với các trang tính
  const classOptions = {
    'Sheet1': ['1A1', '1A2', '1A3', '1A4', '1A5'],
    'Sheet2': ['2A1', '2A2', '2A3', '2A4', '2A5', '2A6'],
    'Sheet3': ['3A1', '3A2', '3A3', '3A4', '3A5', '3A6'],
    'Sheet4': ['4A2', '4A3', '4A4', '4A5', '4A6', '4A7'],
    'Sheet5': ['5A1', '5A2', '5A3', '5A4', '5A5', '5A6']
  };

  // Cập nhật danh sách lớp khi thay đổi trang tính
  const updateClassDropdown = (grade) => {
    const classSelect = document.getElementById('classSelect');
    classSelect.innerHTML = '<option value="" disabled selected>Select a class</option>';

    // Tạo các tùy chọn lớp học dựa trên lớp được chọn
    const classes = classOptions[grade] || [];
    classes.forEach(className => {
      const option = document.createElement('option');
      option.value = className;
      option.textContent = className;
      classSelect.appendChild(option);
    });

    classSelect.disabled = false; // Bật tùy chọn lớp học khi lớp được chọn
  };


  const fetchDataFromGoogleSheetICT = async (gradeSelect) => {
  const baseUrl = `https://docs.google.com/spreadsheets/d/${googleSheetICTId}/gviz/tq?sheet=${gradeSelect}&tqx=out:json`;
  try {
    const response = await fetch(baseUrl);
    const text = await response.text();
    // Phân tích cú pháp JSON từ chuỗi trả về
    const json = JSON.parse(text.substr(47).slice(0, -2));
    const rows = json.table.rows;

    // Tạo đối tượng sheetData để lưu trữ dữ liệu theo các danh mục
    sheetData = {
      introduction: [],
      behavior: [],
      classwork: [],
      participation: [],
      improvements: []
    };

    // Lặp qua mỗi hàng và lưu các dữ liệu vào danh mục tương ứng
    rows.forEach((row, index) => {
      if (index > 0) { // Bỏ qua hàng đầu tiên là tiêu đề
        sheetData.introduction.push(row.c[0]?.v || '');
        sheetData.behavior.push(row.c[1]?.v || '');
        sheetData.classwork.push(row.c[2]?.v || '');
        sheetData.participation.push(row.c[3]?.v || '');
        sheetData.improvements.push(row.c[4]?.v || '');
      }
    });

    // Lưu dữ liệu mới vào localStorage
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
      // Phân tích cú pháp JSON từ chuỗi trả về
      const json = JSON.parse(text.substr(47).slice(0, -2));
      const newData = json.table.rows.map(row => row.c.map(cell => (cell ? cell.v : '')));
      // Lưu dữ liệu mới vào localStorage
      sheetData = newData;
      localStorage.setItem('sheetData', JSON.stringify(sheetData));
      return newData;
    } catch (error) {
      console.error('Error fetching data from Google Sheet:', error);
    }
  };


  // Cập nhật dữ liệu bảng dựa trên lớp học được chọn
  const updateTableData = async (selectedClass) => {

    const tableBody = document.getElementById('commentTable').querySelector('tbody');
    tableBody.innerHTML = ''; // Xóa nội dung bảng hiện có

    // Lấy dữ liệu từ Google Sheets dựa trên lớp học được chọn
    const data = await fetchDataFromGoogleSheetClasses(selectedClass);
   
    // Lấy dữ liệu từ Google Sheets ICT 
    const dataGrade = await fetchDataFromGoogleSheetICT(gradeSelected);

    if (data && data.length > 0) {
      // Lặp qua mỗi hàng dữ liệu và thêm vào bảng
      data.forEach((rowData, rowIndex) => {
        const newRow = tableBody.insertRow();
        // Thêm dữ liệu vào các cột tương ứng
        rowData.slice(0, 10).forEach((cellData, index) => {
          const cell = newRow.insertCell(index);
          cell.textContent = cellData;
        });

        // Gọi hàm để thêm các dropdown nhận xét cho từng hàng học sinh
        addCommentColumns(newRow, rowData[4], rowData[3], rowIndex);
      });
    }

    // Thiết lập các sự kiện cho cột Summary và độ rộng của dropdown
    setupSummaryEvents();
    enforceDropdownWidth();
    restoreSavedData();
  };

  // Hàm để thêm các dropdown nhận xét cho mỗi hàng của bảng
  const addCommentColumns = (newRow, gender, firstName, rowIndex) => {
    const selectedComments = [];

    // Lặp qua các danh mục nhận xét và thêm dropdown
    ["introduction", "behavior", "classwork", "participation", "improvements"].forEach((category, index) => {
      const cell = newRow.insertCell();
      const container = document.createElement('div');
      container.classList.add('comment-container');

      const select = document.createElement('select');
      select.dataset.category = category;
      select.dataset.rowIndex = rowIndex;
      select.dataset.columnIndex = index;
      populateOptions(select, category); // Đổ dữ liệu tùy chọn vào dropdown

      container.appendChild(select);
      cell.appendChild(container);

      selectedComments.push(select);
    });

    // Thêm cột tóm tắt (Summary)
    const summaryCell = newRow.insertCell();
    summaryCell.classList.add('summary-column');
    const summaryDiv = document.createElement('div');
    summaryDiv.contentEditable = true;
    summaryCell.appendChild(summaryDiv);

    // Thiết lập sự kiện thay đổi cho các dropdown để cập nhật tóm tắt
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

 // Đổ dữ liệu tùy chọn vào dropdown dựa trên danh mục
const populateOptions = (select, category) => {
  // Xóa nội dung cũ trong dropdown
  select.innerHTML = '';

  // Lấy dữ liệu từ sheetData dựa trên danh mục
  const data = sheetData[category];

  if (data && Array.isArray(data)) {
    // Lặp qua mỗi mục trong danh mục và thêm vào dropdown
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

  // Cập nhật nội dung tóm tắt khi người dùng thay đổi dropdown
  const updateSummary = (firstName, gender, selectedComments, summaryCell) => {
    const pronouns = {
      'M': { subject: 'He', object: 'him', possessive: 'his' },
      'F': { subject: 'She', object: 'her', possessive: 'her' }
    };

    const currentPronoun = pronouns[gender] || pronouns['M'];

    // Tạo nội dung tóm tắt từ các giá trị nhận xét đã chọn
    let summaryText = `${firstName} `;

    const introduction = selectedComments[0].value;
    const behavior = `${currentPronoun.subject} ${selectedComments[1].value.toLowerCase()}`;
    const classwork = `${currentPronoun.possessive.charAt(0).toUpperCase() + currentPronoun.possessive.slice(1)} classwork ${selectedComments[2].value.toLowerCase()}`;
    const participation = `${currentPronoun.subject} ${selectedComments[3].value.toLowerCase()}`;
    const improvements = selectedComments[4].value;

    summaryText += `${introduction}. ${behavior}. ${classwork}. ${participation}. ${improvements}.`;

    // Xử lý định dạng và giới hạn số ký tự
    summaryText = summaryText.replace(/\s+/g, ' ').replace(/(\.\s*)+/g, '. ').trim();

    // Giới hạn ký tự hiển thị
    if (summaryText.length > maxLengthDisplay) {
      summaryCell.dataset.fullText = summaryText;
      summaryText = summaryText.slice(0, maxLengthDisplay) + "...";
    }
    summaryCell.textContent = summaryText;
  };

  // Thiết lập sự kiện mở rộng/tắt cho cột Summary
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
          this.style.width = ''; // Reset chiều rộng về mặc định khi thu gọn
          this.textContent = this.dataset.fullText.slice(0, maxLengthDisplay) + "...";
        } else {
          this.classList.add('expanded');
          this.style.width = '100%'; // Mở rộng chiều rộng khi mở
          this.textContent = this.dataset.fullText;
        }
      });
    });
  };

  // Đảm bảo độ rộng dropdown phù hợp khi người dùng nhấp chuột
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

  // Lưu dữ liệu đã chọn vào localStorage
  const saveData = (rowIndex, columnIndex, value) => {
    let savedData = JSON.parse(localStorage.getItem('savedComments')) || {};
    if (!savedData[rowIndex]) {
      savedData[rowIndex] = {};
    }
    savedData[rowIndex][columnIndex] = value;
    localStorage.setItem('savedComments', JSON.stringify(savedData));
  };

  // Khôi phục dữ liệu đã lưu từ localStorage
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

  // Đặt lại bảng và xóa dữ liệu đã lưu
  const resetAll = () => {
    document.getElementById('commentTable').querySelector('tbody').innerHTML = '';
    localStorage.removeItem('savedComments');
  };

  document.getElementById('gradeSelect').addEventListener('change', function () {
    gradeSelected = this.value;
    localStorage.setItem('selectedSheet', gradeSelected);
    // Xóa dữ liệu sheet cũ trong localStorage khi chọn Grade mới
    localStorage.removeItem('sheetData');
    sheetData = {}; // Đặt lại sheetData để không lưu dữ liệu cũ
    updateClassDropdown(gradeSelected);
  });

  // document.getElementById('subjectSelect').addEventListener('change', function () {
  //   selectedSubject = this.value;
  //   localStorage.setItem('selectedSubject', selectedSubject);
  //   // Làm mới dữ liệu khi thay đổi môn học
  //   localStorage.removeItem('sheetData');
  //   sheetData = {}; // Đặt lại sheetData để không lưu dữ liệu cũ
  // });

  // Bật nút "Go" khi lớp được chọn
  document.getElementById('classSelect').addEventListener('change', function () {
    document.getElementById('goBtn').disabled = false;
  });

  // Sự kiện khi nhấn nút "Go" để cập nhật dữ liệu bảng
  document.getElementById('goBtn').addEventListener('click', async function () {
    const selectedClass = document.getElementById('classSelect').value;
    await updateTableData(selectedClass);
    populateDropdowns();
  });

  // Sự kiện khi nhấn nút "Save" để thông báo lưu dữ liệu
  document.getElementById('saveBtn').addEventListener('click', function () {
    // Dữ liệu được lưu tự động khi thay đổi dropdown
    alert("Data saved successfully!");
  });

  // Sự kiện khi nhấn nút "Create New" để đặt lại bảng
  document.getElementById('createNewBtn').addEventListener('click', function () {
    resetAll();
  });

  // Thiết lập sự kiện tóm tắt cho cột Summary
  setupSummaryEvents();
});
