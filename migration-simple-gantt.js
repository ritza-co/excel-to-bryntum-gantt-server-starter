import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import xlsx from "xlsx";

function excelToJSDate(excelDate) {
  const dateObj = xlsx.SSF.parse_date_code(excelDate);
  const jsDate = new Date(
    dateObj.y,
    dateObj.m - 1,
    dateObj.d,
    dateObj.H,
    dateObj.M,
    dateObj.S
  );

  return jsDate.toString();
}

function createBryntumTasksRows(data) {
  let taskId = 0;
  let taskStore = [];
  let currentParentId;

  // name of first task is "Phase 1 Title"
  // only start adding tasks after it's found
  const firstTaskName = "Phase 1 Title";
  const exampleTaskName = "Insert new rows ABOVE this one";
  let isFirstTaskFound = false;
  for (let i = 0; i < data.length; i++) {
    if (
      data[i].hasOwnProperty("PROJECT TITLE") &&
      data[i]["PROJECT TITLE"].startsWith("Phase")
    ) {
      if (data[i]["PROJECT TITLE"] === firstTaskName) {
        isFirstTaskFound = true;
      }
      currentParentId = taskId;
      // parent tasks
      taskStore.push({
        id: taskId++,
        name: data[i]["PROJECT TITLE"],
        expanded: true,
      });
    } else if (data[i]["PROJECT TITLE"]) {
      if (!isFirstTaskFound) {
        continue;
      }
      // last task has been added
      if (data[i]["PROJECT TITLE"] === exampleTaskName) {
        break;
      }
      // child tasks
      taskStore.push({
        id: taskId++,
        name: data[i]["PROJECT TITLE"],
        parentId: currentParentId,
        resourceAssignment: data[i]["__EMPTY"],
        percentDone: data[i]["__EMPTY_1"] * 100,
        startDate: excelToJSDate(data[i]["__EMPTY_2"]),
        endDate: excelToJSDate(data[i]["__EMPTY_3"]),
        manuallyScheduled: true,
      });
    }
  }

  return taskStore;
}

// Read the Excel file
const workbook = xlsx.readFile("./simple-gantt-chart.xlsx");
const sheetName = workbook.SheetNames[0]; // select the sheet you want
const worksheet = workbook.Sheets[sheetName];

const jsonData = xlsx.utils.sheet_to_json(worksheet);
const tasksRows = createBryntumTasksRows(jsonData);
// create resources
const resourceNames = new Set();
const resourcesRows = tasksRows
  .filter(
    (item) =>
      item.resourceAssignment && !resourceNames.has(item.resourceAssignment)
  )
  .map((item, i) => {
    const name = item.resourceAssignment;
    return {
      id: i,
      name,
    };
  });

// create assignments
const resourcesWithAssignments = tasksRows.filter(
  (item) => item.resourceAssignment
);

const assignmentsRows = resourcesWithAssignments.map((item, i) => {
  const resource = resourcesRows.find(
    (resource) => resource.name === item.resourceAssignment
  );
  return {
    id: i,
    event: item.id,
    resource: resource.id,
  };
});

// Convert JSON data to the expected load response structure
const ganttLoadResponse = {
  success: true,
  tasks: {
    rows: tasksRows,
  },
  resources: {
    rows: resourcesRows,
  },
  assignments: {
    rows: assignmentsRows,
  },
};

const dataJson = JSON.stringify(ganttLoadResponse, null, 2); // Convert the data to JSON, indented with 2 spaces

// Define the path to the data folder
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const dataPath = path.join(__dirname, "data");

// Ensure the data directory exists
if (!fs.existsSync(dataPath)) {
  fs.mkdirSync(dataPath);
}

// Define the path to the JSON file in the data folder
const filePath = path.join(dataPath, "simple-gantt-chart.json");

// Write the JSON string to a file in the data directory
fs.writeFile(filePath, dataJson, (err) => {
  if (err) throw err;
  console.log("JSON data written to file");
});
