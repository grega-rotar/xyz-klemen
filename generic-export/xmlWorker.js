const { parentPort, workerData } = require("worker_threads");
const fs = require("fs");
const { XMLParser } = require("fast-xml-parser");

// Clean template values like "<#= 'text' #>" → "text"
function cleanTemplateValue(str) {
    if (!str || typeof str !== "string") return str;
    const match = str.match(/<#=\s*'(.*)'\s*#>/);
    return match ? match[1] : str;
}

// Search for id in XML
function getNameAndSubjectById(xmlFile, searchId) {
    const xmlData = fs.readFileSync(xmlFile, "utf-8");
    const parser = new XMLParser({ ignoreAttributes: false });
    const result = parser.parse(xmlData);

    function search(obj) {
        if (!obj || typeof obj !== "object") return null;

        if (obj["@_id"] === searchId) {
            const nameValue = obj.name || null;
            let subjectValue = null;
            if (obj.component?.implementation?.subject) {
                subjectValue = cleanTemplateValue(obj.component.implementation.subject);
            }
            return { name: nameValue, subject: subjectValue };
        }

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

// Worker entry
const { xmlFile, id } = workerData;
const result = getNameAndSubjectById(xmlFile, id);
parentPort.postMessage(result);
