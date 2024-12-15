let SHEET_ID = null;
const SHEETS = ["Пн", "Вт", "Ср", "Чт", "Пт"];

function toggleLoader(show) {
  const loader = document.getElementById("loader");
  loader.style.display = show ? "block" : "none";
}

function displayError(message) {
  const errorDisplay = document.getElementById("errorDisplay");
  errorDisplay.textContent = message;
}

function clearDisplays() {
  document.getElementById("jsonDisplay").innerHTML = "";
  document.getElementById("errorDisplay").textContent = "";
}

function displayData(data) {
  const display = document.getElementById("jsonDisplay");
  display.innerHTML = "";
  if (Object.keys(data).length === 0) {
    display.textContent = "Нет данных для отображения.";
    return;
  }
  for (const [employee, daysData] of Object.entries(data)) {
    const employeeCard = document.createElement("div");
    employeeCard.className = "employee-card";
    const employeeHeader = document.createElement("div");
    employeeHeader.className = "employee-header";
    employeeHeader.textContent = employee;
    employeeCard.appendChild(employeeHeader);
    for (const [day, valuesArray] of Object.entries(daysData)) {
      const dayCard = document.createElement("div");
      dayCard.className = "day-card";
      const dayHeader = document.createElement("div");
      dayHeader.className = "day-header";
      dayHeader.textContent = day;
      dayCard.appendChild(dayHeader);
      if (valuesArray.length > 0) {
        const valuesList = document.createElement("ul");
        valuesList.className = "values-list";
        valuesArray.forEach((value) => {
          const listItem = document.createElement("li");
          listItem.textContent = value;
          valuesList.appendChild(listItem);
        });
        dayCard.appendChild(valuesList);
      } else {
        const noData = document.createElement("div");
        noData.textContent = "Нет данных.";
        dayCard.appendChild(noData);
      }
      employeeCard.appendChild(dayCard);
    }
    display.appendChild(employeeCard);
  }
}

async function downloadAndStoreGoogleSheets(
  sheetId,
  sheetNames,
  range = "B4:M100"
) {
  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  try {
    toggleLoader(true);
    clearDisplays();
    const response = await fetch(exportUrl, {
      method: "GET",
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });
    if (!response.ok) {
      throw new Error(`Ошибка загрузки таблицы: ${response.statusText}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const masterData = {};
    for (const sheetName of sheetNames) {
      if (!workbook.SheetNames.includes(sheetName)) {
        continue;
      }
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        defval: null,
        range: range,
        header: 1,
        blankrows: false,
      });
      if (jsonData.length === 0) {
        continue;
      }
      for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        const employeeName = row[0];
        if (
          employeeName === undefined ||
          employeeName === null ||
          employeeName.toString().trim() === ""
        ) {
          continue;
        }
        if (!masterData.hasOwnProperty(employeeName)) {
          masterData[employeeName] = {};
        }
        const valuesArray = row
          .slice(1)
          .filter(
            (value) =>
              value !== null &&
              value !== undefined &&
              value.toString().trim() !== ""
          );
        masterData[employeeName][sheetName] = valuesArray;
      }
    }
    if (Object.keys(masterData).length === 0) {
      throw new Error("Не удалось создать объект из данных.");
    }
    localStorage.setItem("googleSheetDataMap", JSON.stringify(masterData));
    populateEmployeeSelect(masterData);
    setDefaultDaySelect();
    displaySelectedData();
    document.getElementById("selectContainer").style.display = "block";
  } catch (error) {
    displayError(error.message);
  } finally {
    toggleLoader(false);
  }
}

function loadMapFromLocalStorage() {
  const mapData = localStorage.getItem("googleSheetDataMap");
  if (mapData) {
    try {
      const dataObject = JSON.parse(mapData);
      populateEmployeeSelect(dataObject);
      setDefaultDaySelect();
      displaySelectedData();
      document.getElementById("selectContainer").style.display = "block";
    } catch (error) {
      displayError("Ошибка при загрузке данных из localStorage.");
    }
  } else {
    document.getElementById("selectContainer").style.display = "none";
  }
}

function extractSheetId(url) {
  const regex = /\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

function populateEmployeeSelect(data) {
  const employeeSelect = document.getElementById("employeeSelect");
  employeeSelect.innerHTML = '<option value="">Выберите сотрудника</option>';
  const employees = Object.keys(data).sort();
  employees.forEach((employee) => {
    const option = document.createElement("option");
    option.value = employee;
    option.textContent = employee;
    employeeSelect.appendChild(option);
  });
  const savedEmployee = localStorage.getItem("selectedEmployee");
  if (savedEmployee && employees.includes(savedEmployee)) {
    employeeSelect.value = savedEmployee;
  }
}

function setDefaultDaySelect() {
  const daySelect = document.getElementById("daySelect");
  const currentDay = new Date().getDay();
  let dayValue = "Пн";
  switch (currentDay) {
    case 1:
      dayValue = "Пн";
      break;
    case 2:
      dayValue = "Вт";
      break;
    case 3:
      dayValue = "Ср";
      break;
    case 4:
      dayValue = "Чт";
      break;
    case 5:
      dayValue = "Пт";
      break;
    default:
      dayValue = "Пн";
  }
  daySelect.value = dayValue;
}

function displaySelectedData() {
  const employeeSelect = document.getElementById("employeeSelect");
  const daySelect = document.getElementById("daySelect");
  const display = document.getElementById("jsonDisplay");
  const dataMap = JSON.parse(
    localStorage.getItem("googleSheetDataMap") || "{}"
  );
  const selectedEmployee = employeeSelect.value;
  const selectedDay = daySelect.value;
  const compareButton = document.getElementById("compareButton");
  const originalLink = localStorage.getItem("originalSheetLink");

  if (selectedEmployee && selectedDay) {
    const employeeData = dataMap[selectedEmployee];
    if (employeeData && employeeData[selectedDay]) {
      const valuesArray = employeeData[selectedDay];
      display.innerHTML = "";
      const employeeCard = document.createElement("div");
      employeeCard.className = "employee-card";
      const dayCard = document.createElement("div");
      dayCard.className = "day-card";
      if (valuesArray.length > 0) {
        const valuesList = document.createElement("ol");
        valuesList.className = "values-list";
        valuesArray.forEach((value) => {
          const listItem = document.createElement("li");
          listItem.textContent = value;
          valuesList.appendChild(listItem);
        });
        dayCard.appendChild(valuesList);
      } else {
        const noData = document.createElement("div");
        noData.textContent = "Нет данных.";
        dayCard.appendChild(noData);
      }
      employeeCard.appendChild(dayCard);
      display.appendChild(employeeCard);

      if (originalLink) {
        compareButton.href = originalLink;
        compareButton.style.display = "inline-block";
      } else {
        compareButton.style.display = "none";
      }
    } else {
      display.innerHTML = "Нет данных для выбранных опций.";
      compareButton.style.display = "none";
    }
  } else {
    display.innerHTML = "Пожалуйста, выберите сотрудника и день недели.";
    compareButton.style.display = "none";
  }
}

document.getElementById("uploadBtn").addEventListener("click", () => {
  const sheetLink = document.getElementById("sheetLinkInput").value.trim();
  if (!sheetLink) {
    displayError("Пожалуйста, введите ссылку");
    return;
  }
  const sheetId = extractSheetId(sheetLink);
  if (!sheetId) {
    displayError("Неверная ссылка на Google Таблицу.");
    return;
  }
  SHEET_ID = sheetId;
  localStorage.setItem("originalSheetLink", sheetLink);
  downloadAndStoreGoogleSheets(SHEET_ID, SHEETS, "B5:M50");
});

function setupSelectEventListeners() {
  const employeeSelect = document.getElementById("employeeSelect");
  const daySelect = document.getElementById("daySelect");
  employeeSelect.addEventListener("change", () => {
    const selectedEmployee = employeeSelect.value;
    if (selectedEmployee) {
      localStorage.setItem("selectedEmployee", selectedEmployee);
    } else {
      localStorage.removeItem("selectedEmployee");
    }
    displaySelectedData();
  });
  daySelect.addEventListener("change", () => {
    displaySelectedData();
  });
}

function initializeSelects() {
  setupSelectEventListeners();
}

window.addEventListener("DOMContentLoaded", () => {
  loadMapFromLocalStorage();
  initializeSelects();
  const currentDateElement = document.getElementById("currentDate");
  const now = new Date();
  const options = {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
  };
  const formattedDate = now.toLocaleDateString("ru-RU", options);
  currentDateElement.textContent = `${
    formattedDate.charAt(0).toUpperCase() + formattedDate.slice(1)
  }`;
});
