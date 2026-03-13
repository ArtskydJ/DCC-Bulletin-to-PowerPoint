'use strict';

const AdmZip = require('adm-zip');
const path = require('path');

// Read the slide canvas dimensions (in EMU) from ppt/presentation.xml.
// Returns { cx, cy } or null if not found.
function readSlideDimensions(zip) {
  const presEntry = zip.getEntry('ppt/presentation.xml');
  if (!presEntry) return null;
  const xml = presEntry.getData().toString('utf8');
  const cx = xml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/)?.[1];
  const cy = xml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/)?.[1];
  if (!cx || !cy) return null;
  return { cx: parseInt(cx), cy: parseInt(cy) };
}

// Extract all slides from a .pptx file.
// Returns { slides: [{ slideXml, relsXml, mediaFiles, slideNum }], sourceDimensions: { cx, cy } | null }
function extractSlides(pptxPath) {
  const zip = new AdmZip(pptxPath);
  const entries = zip.getEntries();
  const sourceDimensions = readSlideDimensions(zip);

  // Find all slide files
  const slideEntries = entries
    .filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
    .sort((a, b) => {
      const na = parseInt(a.entryName.match(/slide(\d+)/)[1]);
      const nb = parseInt(b.entryName.match(/slide(\d+)/)[1]);
      return na - nb;
    });

  const slides = [];

  for (const entry of slideEntries) {
    const slideXml = entry.getData().toString('utf8');
    const slideNum = parseInt(entry.entryName.match(/slide(\d+)/)[1]);

    // Get relationships for this slide
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const relsEntry = zip.getEntry(relsPath);
    const relsXml = relsEntry ? relsEntry.getData().toString('utf8') : null;

    // Find referenced media
    const mediaFiles = {};
    if (relsXml) {
      const relMatches = relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g);
      for (const m of relMatches) {
        const mediaName = m[1];
        const mediaEntry = zip.getEntry(`ppt/media/${mediaName}`);
        if (mediaEntry) {
          mediaFiles[mediaName] = mediaEntry.getData();
        }
      }
    }

    slides.push({ slideXml, relsXml, mediaFiles, slideNum });
  }

  return { slides, sourceDimensions };
}

// Extract just the text content from a pptx for preview purposes
function extractSlideTexts(pptxPath) {
  const zip = new AdmZip(pptxPath);
  const entries = zip.getEntries();

  const slideEntries = entries
    .filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
    .sort((a, b) => {
      const na = parseInt(a.entryName.match(/slide(\d+)/)[1]);
      const nb = parseInt(b.entryName.match(/slide(\d+)/)[1]);
      return na - nb;
    });

  const results = [];
  for (const entry of slideEntries) {
    const xml = entry.getData().toString('utf8');
    const textNodes = xml.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g) || [];
    const text = textNodes
      .map(t => t.replace(/<[^>]+>/g, '').trim())
      .filter(Boolean)
      .join(' ');
    results.push(text);
  }

  return results;
}

// Get song title (first text from first slide)
function getSongTitle(pptxPath) {
  const texts = extractSlideTexts(pptxPath);
  return texts.length > 0 ? texts[0].split('|')[0].trim() : path.basename(pptxPath, '.pptx');
}

module.exports = { extractSlides, extractSlideTexts, getSongTitle };
