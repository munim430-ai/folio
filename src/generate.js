'use strict';

const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');

// ─── XML escape ────────────────────────────────────────────────────────────────
function xmlEscape(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ─── Date formatters ───────────────────────────────────────────────────────────
function formatDateP1P2(year, month, day) {
  const y = String(year || '').padStart(4, '0');
  const m = String(month || '').padStart(2, '0');
  const d = String(day || '').padStart(2, '0');
  return `${y} / ${m} / ${d}`;
}

function formatDateP3(year, month, day) {
  const m = String(month || '').padStart(2, '0');
  const d = String(day || '').padStart(2, '0');
  const y = String(year || '').padStart(4, '0');
  return `: ${m} / ${d} / ${y}`;
}

// ─── Simple marker replacement in XML text ─────────────────────────────────────
function replaceMarkers(xml, markers) {
  let result = xml;
  for (const [marker, value] of Object.entries(markers)) {
    // Markers appear in <w:t> elements, escape value for XML
    const escaped = xmlEscape(value);
    // Use global replace (markers should be unique but be safe)
    result = result.split(marker).join(escaped);
  }
  return result;
}

// ─── Photo injection ───────────────────────────────────────────────────────────
const PHOTO_REL_ID = 'rId99';

function buildPhotoDrawingXml(relId) {
  // EMU: 1cm = 360000 EMU. 3cm wide = 1080000, 4cm tall = 1440000
  return `<w:p>` +
    `<w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr/><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<wp:extent cx="1080000" cy="1440000"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="99" name="Photo"/>` +
    `<wp:cNvGraphicFramePr>` +
    `<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>` +
    `</wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr>` +
    `<pic:cNvPr id="99" name="Photo"/>` +
    `<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>` +
    `</pic:nvPicPr>` +
    `<pic:blipFill>` +
    `<a:blip r:embed="${relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
    `<a:stretch><a:fillRect/></a:stretch>` +
    `</pic:blipFill>` +
    `<pic:spPr bwMode="auto">` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1080000" cy="1440000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</pic:spPr>` +
    `</pic:pic>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    `</wp:inline>` +
    `</w:drawing></w:r>` +
    `</w:p>`;
}

/**
 * Replace the content of the paragraph identified by paraId with new XML.
 * Preserves everything outside the <w:p...>...</w:p> block.
 */
function replaceParagraphContent(xml, paraId, newPContent) {
  const idx = xml.indexOf(paraId);
  if (idx < 0) return xml;

  // Find the <w:p opening tag start
  const pStart = xml.lastIndexOf('<w:p ', idx);
  if (pStart < 0) return xml;

  // Find the closing </w:p>
  const pEnd = xml.indexOf('</w:p>', idx);
  if (pEnd < 0) return xml;

  const before = xml.substring(0, pStart);
  const after = xml.substring(pEnd + 6);

  return before + newPContent + after;
}

/**
 * Find photo file for a student name, trying multiple extensions.
 */
function findPhotoFile(photosDir, fullName) {
  const extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG'];
  for (const ext of extensions) {
    const filePath = path.join(photosDir, fullName + ext);
    if (fs.existsSync(filePath)) return filePath;
  }
  return null;
}

/**
 * Inject photo into the zip:
 * - Add image file as word/media/photo_student.{ext}
 * - Add relationship entry to word/_rels/document.xml.rels
 * Returns the relationship ID used, or null on failure.
 */
async function injectPhoto(zip, photoPath) {
  try {
    const imageBuffer = fs.readFileSync(photoPath);
    const ext = path.extname(photoPath).toLowerCase().replace('.', '');
    const mediaName = `photo_student.${ext}`;

    // Add image to media
    zip.file(`word/media/${mediaName}`, imageBuffer);

    // Update relationships
    const relsPath = 'word/_rels/document.xml.rels';
    let relsXml = await zip.file(relsPath).async('string');

    const relEntry = `<Relationship Id="${PHOTO_REL_ID}" ` +
      `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
      `Target="media/${mediaName}"/>`;

    // Insert before </Relationships>
    relsXml = relsXml.replace('</Relationships>', relEntry + '</Relationships>');
    zip.file(relsPath, relsXml);

    return PHOTO_REL_ID;
  } catch (err) {
    return null;
  }
}

// ─── Main generation function ──────────────────────────────────────────────────

/**
 * Generate DOCX forms for all students.
 *
 * @param {Object} opts
 * @param {Array} opts.students - Array of student objects from parseExcel
 * @param {string} opts.templatePath - Path to the tagged mokwon.docx template
 * @param {string} opts.photosDir - Directory containing student photos
 * @param {string} opts.outputDir - Directory to write output DOCX files
 * @returns {Promise<Object>} { files, generated, warnings }
 */
async function generateForms({ students, templatePath, photosDir, outputDir }) {
  const templateBuffer = fs.readFileSync(templatePath);
  const warnings = [];
  const files = [];
  let generated = 0;

  for (let i = 0; i < students.length; i++) {
    const student = students[i];

    if (!student._valid) {
      warnings.push({
        type: 'error',
        row: student._rowIndex,
        message: `Skipped row ${student._rowIndex}: missing required fields: ${student._missingFields.join(', ')}`,
      });
      continue;
    }

    try {
      // Load fresh zip for each student
      const zip = await JSZip.loadAsync(templateBuffer);

      // Build markers map — signature dates use TODAY, birth fields use student data
      const now = new Date();
      const todayYear = now.getFullYear();
      const todayMonth = String(now.getMonth() + 1).padStart(2, '0');
      const todayDay = String(now.getDate()).padStart(2, '0');
      const dateP1P2 = `${todayYear} / ${todayMonth} / ${todayDay}`;
      const dateP3 = `: ${todayMonth} / ${todayDay} / ${todayYear}`;

      const isMale = String(student.gender).toUpperCase() === 'M';
      const isFemale = String(student.gender).toUpperCase() === 'F';

      const markers = {
        FOLIOFULL: student.full_name,
        FOLIOGM: isMale ? '☑' : '□',
        FOLIOGF: isFemale ? '☑' : '□',
        FOLIOBIRY: student.birth_year,
        FOLIOBIRMTH: student.birth_month,
        FOLIOBIRDAY: student.birth_day,
        FOLIONAT: 'Bangladesh',
        FOLIOPASS: student.passport_number,
        FOLIOPOB: student.place_of_birth,
        FOLIOADDR: student.address,
        FOLIOPHONE: student.phone,
        FOLIOSCHOOL: student.school_name,
        FOLIOAGNCY: student.agency_name,
        FOLIOSIG1: student.full_name,
        FOLIODATE1: dateP1P2,
        FOLIOSIG2: student.full_name,
        FOLIODATE2: dateP1P2,
        FOLISPONM: student.sponsor_name,
        FOLISPONAD: student.sponsor_address,
        FOLISPONOCC: student.sponsor_occupation,
        FOLISPONREL: student.sponsor_relation,
        FOLISPONHP: student.sponsor_contact_HP,
        FOLIOSIG3: student.sponsor_name,
        FOLIODATE3: dateP3,
      };

      // Replace markers in document.xml
      let docXml = await zip.file('word/document.xml').async('string');
      docXml = replaceMarkers(docXml, markers);

      // Handle photo — use explicit photo_path from Excel if provided, else name-based lookup
      let photoPath = null;
      if (student.photo_path && String(student.photo_path).trim()) {
        const explicit = path.isAbsolute(String(student.photo_path).trim())
          ? String(student.photo_path).trim()
          : path.join(photosDir, String(student.photo_path).trim());
        if (fs.existsSync(explicit)) photoPath = explicit;
      }
      if (!photoPath) photoPath = findPhotoFile(photosDir, student.full_name);
      if (photoPath) {
        const relId = await injectPhoto(zip, photoPath);
        if (relId) {
          const photoDrawing = buildPhotoDrawingXml(relId);
          docXml = replaceParagraphContent(docXml, '61ED9A81', photoDrawing);
        } else {
          warnings.push({
            type: 'warning',
            row: student._rowIndex,
            message: `Photo injection failed for ${student.full_name}, using placeholder`,
          });
        }
      } else {
        warnings.push({
          type: 'warning',
          row: student._rowIndex,
          message: `Photo not found for ${student.full_name}, using placeholder`,
        });
      }

      zip.file('word/document.xml', docXml);

      // Build output filename
      const safeId = String(student.student_id || '').replace(/[/\\:*?"<>|]/g, '_');
      const safeName = String(student.full_name || '').replace(/[/\\:*?"<>|]/g, '_');
      const filename = `${safeId}_${safeName}.docx`;
      const outputPath = path.join(outputDir, filename);

      // Write zip
      const outputBuffer = await zip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      });

      fs.writeFileSync(outputPath, outputBuffer);

      files.push({
        applicant: student.full_name,
        filename,
        path: outputPath,
        index: i,
      });

      generated++;
    } catch (err) {
      warnings.push({
        type: 'error',
        row: student._rowIndex,
        message: `Failed to generate for row ${student._rowIndex} (${student.full_name}): ${err.message}`,
      });
    }
  }

  return { files, generated, warnings };
}

module.exports = { generateForms };
