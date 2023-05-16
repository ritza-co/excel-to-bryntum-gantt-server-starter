import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import xlsx from "xlsx";
import csvToJson from "csvtojson";

function convertToDate(dateStr) {
  let [month, day, year] = dateStr.split("/");
  year = year.length === 2 ? "20" + year : year;
  month = month - 1;
  const date = new Date(year, month, day);
  return date.toString();
}

function createBryntumTasksRows(data) {
  let taskId = 0;
  let taskStore = [];
  let currentParentId;

  // name of first task is "Phase 1 Title"
  // only start adding tasks after it's found
  const firstTaskName = "Project development";
  const exampleTaskName = "Task 4";
  let isFirstTaskFound = false;
  for (let i = 0; i < data.length; i++) {
    // check for first task
    if (data[i]["field2"] === firstTaskName) {
      isFirstTaskFound = true;
    }
    // check for example task
    if (data[i]["field2"] === exampleTaskName) {
      break;
    }

    if (isFirstTaskFound) {
      // parent task
      if (data[i]["field6"] === "") {
        currentParentId = taskId;
        taskStore.push({
          id: taskId++,
          name: data[i]["field2"],
          expanded: true,
        });
      } else {
        // child tasks
        taskStore.push({
          id: taskId++,
          name: data[i]["field2"],
          parentId: currentParentId,
          category: data[i]["field3"],
          resourceAssignment: data[i]["field4"] ? data[i]["field4"] : undefined,
          percentDone: data[i]["field5"].slice(0, -1) * 1,
          startDate: convertToDate(data[i]["field6"]),
          duration: data[i]["field3"] === "Milestone" ? 0 : data[i]["field7"],
          manuallyScheduled: true,
        });
      }
    }
  }

  return taskStore;
}

const workbook = xlsx.readFile("./agile-gantt-chart.xlsx");
const sheetName = workbook.SheetNames[1]; // select the sheet you want
const worksheet = workbook.Sheets[sheetName];

const csv = xlsx.utils.sheet_to_csv(worksheet);
const jsonData = await csvToJson().fromString(csv);
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

let dataJson = JSON.stringify(ganttLoadResponse, null, 2); // Convert the data to JSON, indented with 2 spaces

// Define the path to the data folder
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const dataPath = path.join(__dirname, "data");

// Ensure the data directory exists
if (!fs.existsSync(dataPath)) {
  fs.mkdirSync(dataPath);
}

// Define the path to the JSON file in the data folder
const filePath = path.join(dataPath, "agile-gantt-chart.json");

// Write the JSON string to a file in the data directory
fs.writeFile(filePath, dataJson, (err) => {
  if (err) throw err;
  console.log("JSON data written to file");
});
