// Utility functions to parse and format time (assumes "HH:MM" format)
function parseTime(timeStr) {
    const parts = timeStr.split(':');
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  }
  
  function formatTime(minutes) {
    const hrs = Math.floor(minutes / 60);
    const mins = minutes % 60;
    return (hrs < 10 ? "0" + hrs : hrs) + ":" + (mins < 10 ? "0" + mins : mins);
  }
  
  function calculateTimeDifference(start, end) {
    const startMinutes = parseTime(start);
    const endMinutes = parseTime(end);
    let diff = endMinutes - startMinutes;
    // If negative, assume the time spans past midnight
    if (diff < 0) {
      diff += 24 * 60;
    }
    return formatTime(diff);
  }
  
  // When the user clicks "Process Excel File"
  document.getElementById('processFile').addEventListener('click', function() {
    const fileInput = document.getElementById('excelFile');
    if (!fileInput.files || !fileInput.files[0]) {
      alert("Please select an Excel file first!");
      return;
    }
    
    const file = fileInput.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      // Assume the first row contains headers (e.g., day names)
      const headers = sheetData[0];
      const daysOfWeek = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
      
      // This array will store rows in the format: [Day, Start Time, End Time, Duration]
      let processedResults = [];
      
      // Loop through each header. If it is a day, process that column.
      headers.forEach((header, colIndex) => {
        if (daysOfWeek.includes(header)) {
          let times = [];
          // Collect all non-empty time entries from this column
          for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
            let cell = sheetData[rowIndex][colIndex];
            if (cell !== undefined && cell !== "") {
              // Convert cell to string (in case it's a number or formatted as time)
              times.push(cell.toString());
            }
          }
          // Pair the times: first entry is the start, second is the end, etc.
          for (let i = 0; i < times.length - 1; i += 2) {
            const startTime = times[i];
            const endTime = times[i + 1];
            const duration = calculateTimeDifference(startTime, endTime);
            processedResults.push([header, startTime, endTime, duration]);
          }
        }
      });
      
      // Display the processed results in the table
      displayTable(processedResults);
      
      // Store results globally for download
      window.processedResults = processedResults;
    };
    
    reader.readAsArrayBuffer(file);
  });
  
  // Display processed data in the table
  function displayTable(data) {
    const tableHead = document.querySelector('#dataTable thead');
    const tableBody = document.querySelector('#dataTable tbody');
    
    // Clear any existing content
    tableHead.innerHTML = "";
    tableBody.innerHTML = "";
    
    if (data.length === 0) {
      tableBody.innerHTML = "<tr><td colspan='4'>No data processed.</td></tr>";
      return;
    }
    
    // Create header row for the table
    const headerRow = document.createElement('tr');
    const headers = ["Day", "Start Time", "End Time", "Duration"];
    headers.forEach(text => {
      const th = document.createElement('th');
      th.textContent = text;
      headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);
    
    // Create rows for each processed record
    data.forEach(row => {
      const tr = document.createElement('tr');
      row.forEach(cell => {
        const td = document.createElement('td');
        td.textContent = cell;
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });
  }
  
  // Download the processed data as a new Excel file
  document.getElementById('downloadProcessed').addEventListener('click', function() {
    if (!window.processedResults || window.processedResults.length === 0) {
      alert("No processed data available. Please process an Excel file first.");
      return;
    }
    
    // Prepare data with headers for the new Excel file
    const dataForExcel = [
      ["Day", "Start Time", "End Time", "Duration"],
      ...window.processedResults
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(dataForExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ProcessedData");
    XLSX.writeFile(wb, "ProcessedTimeLogs.xlsx");
  });
  