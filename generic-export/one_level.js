const fs = require('fs');

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
// --- Example Usage ---

// 1. Read the JSON data from the example.json file
const rawData = fs.readFileSync('example.json');
const jsonData = JSON.parse(rawData);

// 2. Call the function with the example data
const formattedTasks = findAndFormatTasks(jsonData);

// 3. Output the one-level result
fs.writeFileSync('one_level.json', JSON.stringify(formattedTasks, null, 2));
console.log('Successfully created one_level.json!');