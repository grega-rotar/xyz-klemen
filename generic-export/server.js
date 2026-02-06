const express = require("express");
const fs = require("fs")
const { XMLParser } = require("fast-xml-parser"); // install with npm i fast-xml-parser
const ExcelJS = require('exceljs');
const multer = require("multer");
const AdmZip = require("adm-zip");
const path = require("path");
const { json } = require("stream/consumers");
const app = express();
const PORT = 3001;

// Serve static files (like index.html)
app.use(express.static("public"));

// Configure multer storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, "uploads/"),
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);
        // cb(null, file.fieldname + "-" + uniqueSuffix + path.extname(file.originalname));
        cb(null, file.fieldname + path.extname(file.originalname));
    },
});

const upload = multer({ storage });

function unzipTWX(filePath) {
    try {
        const zip = new AdmZip(filePath);
        const extractPath = path.join("uploads", path.basename(filePath, path.extname(filePath)));

        // Create folder if it doesn't exist
        if (!fs.existsSync(extractPath)) fs.mkdirSync(extractPath, { recursive: true });

        // Extract main TWX
        zip.extractAllTo(extractPath, true);
        console.log(`✅ TWX unzipped to: ${extractPath}`);

        // Check for 'toolkits' folder
        const toolkitsPath = path.join(extractPath, "toolkits");
        if (fs.existsSync(toolkitsPath)) {
            const files = fs.readdirSync(toolkitsPath);

            files.forEach(file => {
                const fullPath = path.join(toolkitsPath, file);

                // Only process ZIP files
                if (fs.statSync(fullPath).isFile() && path.extname(fullPath).toLowerCase() === ".zip") {
                    const innerZip = new AdmZip(fullPath);
                    const innerExtractPath = path.join(toolkitsPath, path.basename(file, ".zip"));

                    // Create folder for inner zip
                    if (!fs.existsSync(innerExtractPath)) fs.mkdirSync(innerExtractPath);

                    innerZip.extractAllTo(innerExtractPath, true);
                    console.log(`✅ Extracted inner zip: ${file} → ${innerExtractPath}`);
                }
            });
        }

        return extractPath;
    } catch (err) {
        console.error("❌ Error unzipping TWX:", err);
        return null;
    }
}

function cleanTask(obj) {
    return Object.fromEntries(
        Object.entries(obj).filter(([key, value]) => {
            if (value === null || value === undefined) return false;
            if (typeof value === "string" && value.trim() === "") return false;
            if (Array.isArray(value) && value.length === 0) return false;
            if (value === false) return false;

            // Special rules
            if (key === "poType" && value === "Service") return false;
            if (key === "eventActionType" && value === 0) return false;
            if (key === "loopType" && value === "none") return false;
            if (key === "type" && value === "activity") return false;
            if (key === "category" && value === "Activity") return false;
            if (key === "colorIcon") return false;
            if (key === "attachedEvents") return false;
            if (key === "postAssignment") return false;
            if (key === "loopType") return false;
            if (key === "MIOrdering") return false;
            if (key === "conditional") return false;


            return true;
        })
    );
}

function formatTaskLabels(item) {
    if (item.label) {
        return {
            ...item,
            label: item.label.replace(/\n/g, " ")
        };
    }
    return item;
}

/**
 * Recursively find all XML files in a directory
 */
function findXmlFiles(directory) {
    let xmlFiles = [];
    const files = fs.readdirSync(directory, { withFileTypes: true });

    for (const file of files) {
        const fullPath = path.join(directory, file.name);
        if (file.isDirectory()) {
            xmlFiles = xmlFiles.concat(findXmlFiles(fullPath));
        } else if (file.isFile() && file.name.toLowerCase().endsWith(".xml")) {
            xmlFiles.push(fullPath);
        }
    }

    return xmlFiles;
}

/**
 * Clean template-like string e.g. "<#= 'Vnos vloge na Portalu' #>" => "Vnos vloge na Portalu"
 */
function cleanTemplateValue(str) {
    if (!str || typeof str !== "string") return str;
    const match = str.match(/<#=\s*'(.*)'\s*#>/);
    return match ? match[1] : str;
}

/**
 * Get both name and subject by ID from an XML file (synchronous)
 */
function getNameAndSubjectById(xmlFile, searchId) {
    const xmlData = fs.readFileSync(xmlFile, "utf-8");
    const parser = new XMLParser({ ignoreAttributes: false });
    const result = parser.parse(xmlData);

    function search(obj) {
        if (!obj || typeof obj !== "object") return null;

        if (obj["@_id"] === searchId) {
            // Extract name
            const nameValue = obj.name || null;

            // Extract subject nested in component -> implementation
            let subjectValue = null;
            if (obj.component && obj.component.implementation && obj.component.implementation.subject) {
                subjectValue = cleanTemplateValue(obj.component.implementation.subject);
            }

            return { name: nameValue, subject: subjectValue };
        }

        // Recursively search all children
        for (const key in obj) {
            const child = obj[key];
            if (typeof child === "object") {
                if (Array.isArray(child)) {
                    for (const item of child) {
                        const found = search(item);
                        if (found) return found;
                    }
                } else {
                    const found = search(child);
                    if (found) return found;
                }
            }
        }

        return null;
    }

    return search(result);
}

/**
 * Find name and subject by ID in any XML file in a directory (synchronous)
 */
function findNameAndSubjectInXml(directory, searchId) {
    const xmlFiles = findXmlFiles(directory);
    for (const xmlFile of xmlFiles) {
        const result = getNameAndSubjectById(xmlFile, searchId);
        if (result) return result;
    }
    return { name: null, subject: null };
}

async function jsonToXlsx(dataSheets, sheetNames, headersSheets, name) {
    const workbook = new ExcelJS.Workbook();

    if (dataSheets.length !== sheetNames.length || dataSheets.length !== headersSheets.length) {
        throw new Error("dataSheets, sheetNames, and headersSheets must have the same length");
    }

    for (let i = 0; i < dataSheets.length; i++) {
        const data = dataSheets[i];
        const sheetName = sheetNames[i];
        const headers = headersSheets[i];

        const worksheet = workbook.addWorksheet(sheetName);

        // Set columns
        worksheet.columns = headers;

        // Add rows
        data.forEach(item => worksheet.addRow(item));

        // Auto-fit column widths
        worksheet.columns.forEach(column => {
            let maxLength = 10; // minimum width
            column.eachCell({ includeEmpty: true }, cell => {
                const cellValue = cell.value ? cell.value.toString() : '';
                if (cellValue.length > maxLength) maxLength = cellValue.length;
            });
            column.width = maxLength + 2; // add some padding
        });

        // Add table (optional Excel table with filters)
        worksheet.addTable({
            name: sheetName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 31), // Excel table name restrictions
            ref: 'A1',
            headerRow: true,
            totalsRow: false,
            columns: headers.map(h => ({ name: h.header, filterButton: true })),
            rows: data.map(obj => headers.map(h => obj[h.key] || ''))
        });
    }

    await workbook.xlsx.writeFile('report/' + name + '.xlsx');
    console.log('Excel file created with multiple sheets and auto column widths!');
}

async function xlsxToJson(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath); // await the async read

    const worksheet = workbook.worksheets[0]; // first sheet
    const jsonData = [];

    // Get headers from the first row
    const headers = worksheet.getRow(1).values.slice(1); // slice(1) skips empty first index

    // Iterate rows
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // skip header

        const rowData = {};
        row.values.slice(1).forEach((value, index) => {
            rowData[headers[index]] = value;
        });

        jsonData.push(rowData);
    });

    return jsonData;
}

/**
 * Recursively finds and formats tasks from a JSON object.
 * It looks for objects where type is 'activity' and activityType is 'task'.
 *
 * @param {object | object[]} data The input JSON data (can be an object or an array).
 * @returns {object} A flattened object with task names as keys.
 */
function findAndFormatTasks(data) {
    const result = [];

    function traverse(obj) {
        // 1. Check if the current object is a task we are looking for
        if (obj && typeof obj === 'object' && !Array.isArray(obj)) {
            if (obj.type === 'activity' && obj.activityType === 'task' && obj.name) {
                // Create a copy of the object to avoid modifying the original
                const taskData = { ...obj };
                result.push(taskData);
            }
        }

        // 2. Recurse through the object's properties
        if (obj && typeof obj === 'object') {
            for (const key in obj) {
                // hasOwnProperty check is not needed with for...in on JSON-parsed objects
                // but it's good practice for objects that might have inherited properties.
                if (Object.prototype.hasOwnProperty.call(obj, key)) {
                    traverse(obj[key]);
                }
            }
        }
    }

    traverse(data);
    return result;
}





app.post(
    "/upload",
    upload.fields([
        { name: "twxFile", maxCount: 1 },
        { name: "objektiFile", maxCount: 1 },
        { name: "taskDataFile", maxCount: 1 },
    ]),
    async (req, res) => {
        // Capture all non-file fields (e.g., appName, environment, etc.)
        const formData = { ...req.body };
        console.log("Form Data:", formData);

        // Save formData temporarily
        fs.writeFileSync("uploads/formData.json", JSON.stringify(formData), "utf-8");

        if (req.files.twxFile) {
            const twxPath = req.files.twxFile[0].path;
            const extractedFolder = unzipTWX(twxPath);
            console.log("TWX extracted folder:", extractedFolder);
        }

        if (req.files.taskDataFile) {
            var json = await xlsxToJson("uploads/taskDataFile.xlsx");
            fs.writeFileSync("uploads/taskDataFile.json", JSON.stringify(json), "utf-8");
            console.log("Task Data parsed to json.")

        }

        res.send(`
            <h2>✅ Files uploaded successfully!</h2>
            <a href="/cleanData">Next step - CLEAN DATA</a>
        `);
    }
);

// deprecated 
app.get("/cleanData", (req, res) => {
    const data = fs.readFileSync(`uploads/objektiFile.json`, "utf-8");

    const objekti = JSON.parse(data);

    // Define allowed task types
    const allowedTaskTypes = ["ServiceTask", "UserTask"];

    // Filter objekti
    var filteredTasks = Array.isArray(objekti)
        ? objekti.filter(item => allowedTaskTypes.includes(item.bpmn2TaskType))
        : [];

    // clean empty keys
    filteredTasks = filteredTasks.map(task => cleanTask(task));
    // remove \n
    filteredTasks = filteredTasks.map(task => formatTaskLabels(task));

    fs.writeFileSync("report/objektiFileCleaned.json", JSON.stringify(filteredTasks), "utf-8");

    res.send(`
        <h2>✅ Objekti files cleaned successfuly!</h2>
        <a href="/findNameSubject">Next step - Find name & subject from TWX</a>
    `);
});

// deprecated 
// app.get("/findNameSubject", (req, res) => {
//     const data = fs.readFileSync(`report/objektiFileCleaned.json`, "utf-8");

//     var objekti = JSON.parse(data);
//     objekti.forEach(task => {
//         let id = task.id;
//         var { name, subject } = findNameAndSubjectInXml("uploads/twxFile", id);
//         console.log(name, " | ", subject);

//         task.name = name;
//         task.subject = subject;
//         task.isEqual = (name == subject);

//     });

//     fs.writeFileSync("report/objektiNameSubject.json", JSON.stringify(objekti), "utf-8");

//     res.send(`
//         <h2>✅ Names & Subjects found successfuly!</h2>
//         <a href="/addNazivTaskaAndAppName">Next step - Add Naziv Taska</a>
//     `);


// });

app.get("/findNameSubject", (req, res) => {
    const data = fs.readFileSync(`uploads/diagram.json`, "utf-8");
    const taskList = findAndFormatTasks(JSON.parse(data));

    const updatedTaskList = taskList.map(({ ID, ...rest }) => ({
        id: ID,
        prazen: "",
        da: true,
        userTask: "userTask",
        ...rest
    }));

    fs.writeFileSync("report/objektiNameSubject.json", JSON.stringify(updatedTaskList), "utf-8");

    res.send(`
        <h2>✅ Names & Subjects found successfuly!</h2>
        <a href="/addNazivTaskaAndAppName">Next step - Add Naziv Taska</a>
    `);


});


app.get("/addNazivTaskaAndAppName", (req, res) => {
    const data = fs.readFileSync("report/objektiNameSubject.json", "utf-8");
    let objekti = JSON.parse(data);

    // Load formData (contains appName)
    let formData = {};
    if (fs.existsSync("uploads/formData.json")) {
        formData = JSON.parse(fs.readFileSync("uploads/formData.json", "utf-8"));
    }

    const appName = formData.appName || "Unknown App";

    objekti.forEach((objekt) => {
        objekt.nazivTaska = objekt.name;
        objekt.app = appName;
    });

    // ---- Remove duplicates by 'id' ----
    objekti = Array.from(
        new Map(objekti.map(o => [o.id, o])).values()
    );

    fs.writeFileSync("report/objektiWithNazivTaska.json", JSON.stringify(objekti), "utf-8");

    res.send(`
        <h2>✅ Naziv taska and App name addedd successfully!</h2>
        <a href="/joinData">Next step - Join Data</a>
    `);
});
const periods = [
    // // 1) Do 9. 11.
    // { key: "countPeriod1", name: "Do 9.11.2025", start: null, end: "2025-11-09T23:59:59.999Z" },

    // // 2) 10.–16. november
    // { key: "countPeriod2", name: "Od 10.11. do 16.11.2025", start: "2025-11-10T00:00:00.000Z", end: "2025-11-16T23:59:59.999Z" },

    // // 3) 17.–23. november
    // { key: "countPeriod3", name: "Od 17.11. do 23.11.2025", start: "2025-11-17T00:00:00.000Z", end: "2025-11-23T23:59:59.999Z" },

    // // 4) 24.–30. november
    // { key: "countPeriod4", name: "Od 24.11. do 30.11.2025", start: "2025-11-24T00:00:00.000Z", end: "2025-11-30T23:59:59.999Z" },

    // // 5) 1. 12. – 7. 12.
    // { key: "countPeriod5", name: "Od 1.12. do 7.12.2025", start: "2025-12-01T00:00:00.000Z", end: "2025-12-07T23:59:59.999Z" },

    // // 6) 8. 12. – 14. 12.
    // { key: "countPeriod6", name: "Od 8.12. do 14.12.2025", start: "2025-12-08T00:00:00.000Z", end: "2025-12-14T23:59:59.999Z" },

    // // 7) 15. 12. – 21. 12.
    // { key: "countPeriod7", name: "Od 15.12. do 21.12.2025", start: "2025-12-15T00:00:00.000Z", end: "2025-12-21T23:59:59.999Z" },

    // // 8) 22. 12. – 28. 12.
    // { key: "countPeriod8", name: "Od 22.12. do 28.12.2025", start: "2025-12-22T00:00:00.000Z", end: "2025-12-28T23:59:59.999Z" },

    // // 9) 29. 12. – 4. 1. 2026
    // { key: "countPeriod9", name: "Od 29.12.2025 do 4.1.2026", start: "2025-12-29T00:00:00.000Z", end: "2026-01-04T23:59:59.999Z" },

    // // 10) 5. 1. – 11. 1. (Closed previously open period)
    // { key: "countPeriod10", name: "Od 5.1. do 11.1.2026", start: "2026-01-05T00:00:00.000Z", end: "2026-01-11T23:59:59.999Z" },

        // 1) Do 9. 11.
    { key: "countPeriod1", name: "Do 9.11.2025", start: null, end: "2025-11-09T23:59:59.999Z" },

    // 2) 10.–16. november
    { key: "countPeriod2", name: "Od 10.11. do 11.1.2026", start: "2025-11-10T00:00:00.000Z", end: "2026-01-11T23:59:59.999Z" },

    // 11) 12. 1. – 18. 1. (Current week, ending today Jan 18th)
    { key: "countPeriod3", name: "Od 12.1. do 18.1.2026", start: "2026-01-12T00:00:00.000Z", end: "2026-01-18T23:59:59.999Z" },

    // 12) 19. 1. – 25. 1. (Current period ending today)
    { key: "countPeriod4", name: "Od 19.1. do 25.1.2026", start: "2026-01-19T00:00:00.000Z", end: "2026-01-25T23:59:59.999Z" },

    // 13) 26. 1. – 30. 1. (Current period ending today)
    { key: "countPeriod5", name: "Od 26.1. do 30.1.2026", start: "2026-01-26T00:00:00.000Z", end: "2026-01-30T23:59:59.999Z" },
];


let dateHeaderForSort = "DATUM KREIRANJA INSTANCE";
dateHeaderForSort = "TASK CREATED";

app.get("/joinData", async (req, res) => {
    // Convert XLSX to JSON
    const json = await xlsxToJson("uploads/taskDataFile.xlsx");
    fs.writeFileSync("uploads/taskDataFile.json", JSON.stringify(json), "utf-8");

    const taskData = JSON.parse(fs.readFileSync("uploads/taskDataFile.json", "utf-8"));
    const objekti = JSON.parse(fs.readFileSync("report/objektiWithNazivTaska.json", "utf-8"));

    let formData = {};
    if (fs.existsSync("uploads/formData.json")) {
        formData = JSON.parse(fs.readFileSync("uploads/formData.json", "utf-8"));
    }

    const appName = formData.appName || "Unknown App";

    // Filter tasks for the current app only
    const taskDataCurrAppOnly = taskData.filter(item => item.APLIKACIJA === appName);

    objekti.forEach(obj => {
        const relatedTasks = taskData.filter(task =>
            task["APLIKACIJA"] === obj.app &&
            task["NAZIV TASKA"].startsWith(obj.nazivTaska)
        );

        // Count tasks per period dynamically
        periods.forEach(period => {
            const count = relatedTasks.filter(task => {
                const taskDate = new Date(task[dateHeaderForSort]);
                const startOk = period.start ? taskDate >= new Date(period.start) : true;
                const endOk = period.end ? taskDate <= new Date(period.end) : true;
                return startOk && endOk;
            }).length;
            obj[period.key] = count;
        });

        // Total count across all tasks
        obj.totalCount = relatedTasks.length;
        obj.show = false;
        obj.show = obj.totalCount > 0 || obj.bpmn2TaskType == "UserTask";

    });

    // Match tasks to objekti
    taskDataCurrAppOnly.forEach(task => {
        const matchedObj = objekti.find(obj => task["NAZIV TASKA"].startsWith(obj.nazivTaska));
        task.nazivTaskaIzObjekti = matchedObj ? matchedObj.nazivTaska : null;
    });

    fs.writeFileSync("report/objektiWithTestCount.json", JSON.stringify(objekti, null, 2), "utf-8");
    fs.writeFileSync("report/taskDataWithMatch.json", JSON.stringify(taskDataCurrAppOnly, null, 2), "utf-8");

    res.send(`
        <h2>✅ Naziv taska and App name addedd successfully!</h2>
        <a href="/generateXLSX">Next step - Generate XLSX</a>
    `);
});


app.get("/generateXLSX", (req, res) => {
    const objekti = JSON.parse(fs.readFileSync("report/objektiWithTestCount.json", "utf-8"));

    const taskData = JSON.parse(fs.readFileSync("report/taskDataWithMatch.json", "utf-8"));

    let formData = {};
    if (fs.existsSync("uploads/formData.json")) {
        formData = JSON.parse(fs.readFileSync("uploads/formData.json", "utf-8"));
    }

    const appName = formData.appName || "Unknown App";

    // Base headers
    const headersObjekti = [
        { header: 'app', key: 'app' },
        { header: 'lane', key: 'lane' },
        { header: 'bpmn2TaskType', key: 'activityType' },
        { header: 'serviceType', key: 'userTask' },
        { header: 'poId', key: 'id' },
        { header: 'snapshotId', key: 'prazen' },
        { header: 'id', key: 'id' },
        { header: 'label', key: 'name' },
        { header: 'x', key: 'x' },
        { header: 'y', key: 'y' },
        { header: 'parent', key: 'prazen' },
        { header: 'name', key: 'name' },
        { header: 'subject', key: 'name' },
        { header: 'isEqual', key: 'prazen' },
        { header: 'NAZIV TASKA', key: 'name' },
        { header: "show", key: "da" }
    ];

    // Add period headers dynamically
    periods.forEach(period => headersObjekti.push({ header: period.name, key: period.key }));


    // Totals
    headersObjekti.push({ header: 'Skupno število (vsa obdobja)', key: 'totalCount' });

    const headersTaskData = [
        { header: "APLIKACIJA", key: "APLIKACIJA" },
        { header: "NAZIV TASKA", key: "NAZIV TASKA" },
        { header: "NAZIV TASKA IZ OBJEKTI", key: "nazivTaskaIzObjekti" }
    ];

    // Generate XLSX
    jsonToXlsx([objekti, taskData], [appName, "test"], [headersObjekti, headersTaskData], formData.appName + " " + dateHeaderForSort);

    res.send("Good");
});



app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
