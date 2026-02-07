// ========================================
// XLSX Generator for IBM BAW
// Pure ES5 - No external libraries
// ========================================

// CRC32 calculation for ZIP format
function crc32(str) {
    var crcTable = [];
    for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) {
            c = ((c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1));
        }
        crcTable[n] = c;
    }

    var crc = 0 ^ (-1);
    for (var i = 0; i < str.length; i++) {
        crc = (crc >>> 8) ^ crcTable[(crc ^ str.charCodeAt(i)) & 0xFF];
    }
    return (crc ^ (-1)) >>> 0;
}

// Convert string to array of bytes
function stringToBytes(str) {
    var bytes = [];
    for (var i = 0; i < str.length; i++) {
        var c = str.charCodeAt(i);
        bytes.push(c & 0xFF);
    }
    return bytes;
}

// Convert number to little-endian bytes
function numberToBytes(num, byteCount) {
    var bytes = [];
    for (var i = 0; i < byteCount; i++) {
        bytes.push(num & 0xFF);
        num = num >>> 8;
    }
    return bytes;
}

// Helper function to escape XML special characters
function escapeXML(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Helper function to get column letter (A, B, C, ... Z, AA, AB, ...)
function getColumnLetter(index) {
    var letter = '';
    while (index >= 0) {
        letter = String.fromCharCode(65 + (index % 26)) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

// Create ZIP file from files array
function createZIP(files) {
    var zipData = [];
    var centralDirectory = [];
    var offset = 0;

    // Write local file headers and file data
    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        var content = file.content;
        var contentBytes = stringToBytes(content);
        var crc = crc32(content);
        var fileName = file.name;
        var fileNameBytes = stringToBytes(fileName);

        // Local file header
        var localHeader = [];
        localHeader = localHeader.concat([0x50, 0x4b, 0x03, 0x04]); // Signature
        localHeader = localHeader.concat([0x0A, 0x00]); // Version needed
        localHeader = localHeader.concat([0x00, 0x00]); // Flags
        localHeader = localHeader.concat([0x00, 0x00]); // Compression method (0 = stored)
        localHeader = localHeader.concat([0x00, 0x00]); // Mod time
        localHeader = localHeader.concat([0x00, 0x00]); // Mod date
        localHeader = localHeader.concat(numberToBytes(crc, 4)); // CRC32
        localHeader = localHeader.concat(numberToBytes(contentBytes.length, 4)); // Compressed size
        localHeader = localHeader.concat(numberToBytes(contentBytes.length, 4)); // Uncompressed size
        localHeader = localHeader.concat(numberToBytes(fileNameBytes.length, 2)); // Filename length
        localHeader = localHeader.concat([0x00, 0x00]); // Extra field length
        localHeader = localHeader.concat(fileNameBytes); // Filename

        zipData = zipData.concat(localHeader);
        zipData = zipData.concat(contentBytes);

        // Central directory entry
        var centralEntry = [];
        centralEntry = centralEntry.concat([0x50, 0x4b, 0x01, 0x02]); // Signature
        centralEntry = centralEntry.concat([0x0A, 0x00]); // Version made by
        centralEntry = centralEntry.concat([0x0A, 0x00]); // Version needed
        centralEntry = centralEntry.concat([0x00, 0x00]); // Flags
        centralEntry = centralEntry.concat([0x00, 0x00]); // Compression method
        centralEntry = centralEntry.concat([0x00, 0x00]); // Mod time
        centralEntry = centralEntry.concat([0x00, 0x00]); // Mod date
        centralEntry = centralEntry.concat(numberToBytes(crc, 4)); // CRC32
        centralEntry = centralEntry.concat(numberToBytes(contentBytes.length, 4)); // Compressed size
        centralEntry = centralEntry.concat(numberToBytes(contentBytes.length, 4)); // Uncompressed size
        centralEntry = centralEntry.concat(numberToBytes(fileNameBytes.length, 2)); // Filename length
        centralEntry = centralEntry.concat([0x00, 0x00]); // Extra field length
        centralEntry = centralEntry.concat([0x00, 0x00]); // File comment length
        centralEntry = centralEntry.concat([0x00, 0x00]); // Disk number
        centralEntry = centralEntry.concat([0x00, 0x00]); // Internal attributes
        centralEntry = centralEntry.concat([0x00, 0x00, 0x00, 0x00]); // External attributes
        centralEntry = centralEntry.concat(numberToBytes(offset, 4)); // Relative offset
        centralEntry = centralEntry.concat(fileNameBytes); // Filename

        centralDirectory = centralDirectory.concat(centralEntry);
        offset += localHeader.length + contentBytes.length;
    }

    var centralDirOffset = zipData.length;
    zipData = zipData.concat(centralDirectory);

    // End of central directory record
    var endRecord = [];
    endRecord = endRecord.concat([0x50, 0x4b, 0x05, 0x06]); // Signature
    endRecord = endRecord.concat([0x00, 0x00]); // Disk number
    endRecord = endRecord.concat([0x00, 0x00]); // Central dir start disk
    endRecord = endRecord.concat(numberToBytes(files.length, 2)); // Entries on this disk
    endRecord = endRecord.concat(numberToBytes(files.length, 2)); // Total entries
    endRecord = endRecord.concat(numberToBytes(centralDirectory.length, 4)); // Central dir size
    endRecord = endRecord.concat(numberToBytes(centralDirOffset, 4)); // Central dir offset
    endRecord = endRecord.concat([0x00, 0x00]); // Comment length

    zipData = zipData.concat(endRecord);

    // Convert to Uint8Array
    var uint8Array = new Uint8Array(zipData.length);
    for (var i = 0; i < zipData.length; i++) {
        uint8Array[i] = zipData[i];
    }

    return uint8Array;
}

// Create XLSX file structure from JSON data
function createXLSX(jsonData) {
    if (!jsonData || jsonData.length === 0) {
        jsonData = [{}];
    }

    // Get headers from first object
    var headers = [];
    for (var key in jsonData[0]) {
        if (jsonData[0].hasOwnProperty(key)) {
            headers.push(key);
        }
    }

    // Generate XML content for sheet1.xml
    var sheetXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    sheetXML += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    sheetXML += '<sheetData>';

    // Header row
    sheetXML += '<row r="1">';
    for (var h = 0; h < headers.length; h++) {
        sheetXML += '<c r="' + getColumnLetter(h) + '1" t="inlineStr">';
        sheetXML += '<is><t>' + escapeXML(headers[h]) + '</t></is>';
        sheetXML += '</c>';
    }
    sheetXML += '</row>';

    // Data rows
    for (var i = 0; i < jsonData.length; i++) {
        sheetXML += '<row r="' + (i + 2) + '">';
        for (var j = 0; j < headers.length; j++) {
            var value = jsonData[i][headers[j]];
            var cellRef = getColumnLetter(j) + (i + 2);

            if (typeof value === 'number') {
                sheetXML += '<c r="' + cellRef + '"><v>' + value + '</v></c>';
            } else {
                var strValue = value !== null && value !== undefined ? String(value) : '';
                sheetXML += '<c r="' + cellRef + '" t="inlineStr">';
                sheetXML += '<is><t>' + escapeXML(strValue) + '</t></is>';
                sheetXML += '</c>';
            }
        }
        sheetXML += '</row>';
    }

    sheetXML += '</sheetData></worksheet>';

    // Other required XML files
    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    contentTypes += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    contentTypes += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    contentTypes += '<Default Extension="xml" ContentType="application/xml"/>';
    contentTypes += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    contentTypes += '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    contentTypes += '</Types>';

    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
    rels += '</Relationships>';

    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    workbook += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    workbook += '<sheets>';
    workbook += '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>';
    workbook += '</sheets>';
    workbook += '</workbook>';

    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    workbookRels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    workbookRels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>';
    workbookRels += '</Relationships>';

    // Create ZIP file
    var files = [
        { name: '[Content_Types].xml', content: contentTypes },
        { name: '_rels/.rels', content: rels },
        { name: 'xl/workbook.xml', content: workbook },
        { name: 'xl/_rels/workbook.xml.rels', content: workbookRels },
        { name: 'xl/worksheets/sheet1.xml', content: sheetXML }
    ];

    return createZIP(files);
}

// Convert XLSX binary data to Base64
function xlsxToBase64(xlsxData) {
    var binary = '';
    for (var i = 0; i < xlsxData.length; i++) {
        binary += String.fromCharCode(xlsxData[i]);
    }
    return btoa(binary);
}

// ========================================
// MAIN FUNCTION FOR IBM BAW
// ========================================
// Generates XLSX file from JSON data
// Parameters:
//   - inputJson: Array of objects (your data)
//   - fileName: Desired filename (string)
// Returns:
//   - Object with fileName and base64 properties
// ========================================
function generateXLSX(inputJson, fileName) {
    // Ensure fileName has .xlsx extension
    var finalFileName = fileName;
    if (finalFileName.indexOf('.xlsx') === -1) {
        finalFileName = finalFileName + '.xlsx';
    }

    // Generate XLSX file
    var xlsxData = createXLSX(inputJson);

    // Convert to base64
    var base64Content = xlsxToBase64(xlsxData);

    // Return object with fileName and base64
    return {
        fileName: finalFileName,
        base64: base64Content
    };
}
