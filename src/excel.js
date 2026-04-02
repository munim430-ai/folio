'use strict';

const XLSX = require('xlsx');

const COLUMNS = [
  'student_id',       // A
  'full_name',        // B
  'gender',           // C
  'birth_year',       // D
  'birth_month',      // E
  'birth_day',        // F
  'passport_number',  // G
  'place_of_birth',   // H
  'address',          // I
  'phone',            // J
  'school_name',      // K
  'agency_name',      // L
  'sponsor_name',     // M
  'sponsor_address',  // N
  'sponsor_occupation', // O
  'sponsor_relation', // P
  'sponsor_contact_HP', // Q
  'photo_path',         // R (optional — relative or absolute path to student photo)
];

const REQUIRED_FIELDS = ['full_name', 'passport_number'];

/**
 * Parse an Excel buffer and return an array of student objects.
 * @param {Buffer} buffer - Raw Excel file buffer
 * @returns {Array<Object>} Array of student objects with validation info
 */
function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellText: true, raw: false });

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('Excel file has no sheets');
  }

  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false,
  });

  // Skip header row (index 0), process from row index 1
  const students = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    // Skip completely empty rows
    if (!row || row.every((cell) => !cell || String(cell).trim() === '')) {
      continue;
    }

    const student = {};
    COLUMNS.forEach((col, colIdx) => {
      const val = row[colIdx];
      student[col] = val != null ? String(val).trim() : '';
    });

    // Validate required fields
    const missingFields = REQUIRED_FIELDS.filter((f) => !student[f]);
    student._valid = missingFields.length === 0;
    student._missingFields = missingFields;
    student._rowIndex = i + 1; // 1-based Excel row number

    students.push(student);
  }

  return students;
}

module.exports = { parseExcel };
