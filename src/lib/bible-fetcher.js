'use strict';

// Downloaded from https://github.com/Amosamevor/Bible-json/blob/main/versions/en/NEW%20KING%20JAMES%20VERSION.json
const bible = require('./assets/NEW KING JAMES VERSION.json');

// Map common book name variants to canonical JSON keys
const BOOK_ALIASES = {
  'gen': 'Genesis',
  'exo': 'Exodus', 'exod': 'Exodus',
  'lev': 'Leviticus',
  'num': 'Numbers',
  'deut': 'Deuteronomy', 'deu': 'Deuteronomy',
  'josh': 'Joshua', 'jos': 'Joshua',
  'judg': 'Judges', 'jdg': 'Judges',
  'rth': 'Ruth',
  '1 sam': '1 Samuel', '1sam': '1 Samuel',
  '2 sam': '2 Samuel', '2sam': '2 Samuel',
  '1 kgs': '1 Kings', '1kgs': '1 Kings', '1 kings': '1 Kings',
  '2 kgs': '2 Kings', '2kgs': '2 Kings', '2 kings': '2 Kings',
  '1 chr': '1 Chronicles', '1chr': '1 Chronicles',
  '2 chr': '2 Chronicles', '2chr': '2 Chronicles',
  'ezr': 'Ezra',
  'neh': 'Nehemiah',
  'est': 'Esther', 'esth': 'Esther',
  // job
  'psalm': 'Psalms',
  'prov': 'Proverbs', 'pro': 'Proverbs',
  'eccl': 'Ecclesiastes', 'ecc': 'Ecclesiastes', 'qoh': 'Ecclesiastes',
  'song of songs': 'Song of Solomon', 'sos': 'Song of Solomon', 'song': 'Song of Solomon',
  'isa': 'Isaiah',
  'jer': 'Jeremiah',
  'lam': 'Lamentations',
  'ezek': 'Ezekiel', 'eze': 'Ezekiel',
  'dan': 'Daniel',
  'hos': 'Hosea',
  // joel
  'amos': 'Amos',
  'obad': 'Obadiah', 'oba': 'Obadiah',
  'jon': 'Jonah',
  'mic': 'Micah',
  'nah': 'Nahum',
  'hab': 'Habakkuk',
  'zeph': 'Zephaniah', 'zep': 'Zephaniah',
  'hag': 'Haggai',
  'zech': 'Zechariah', 'zec': 'Zechariah',
  'mal': 'Malachi',
  'matt': 'Matthew', 'mat': 'Matthew',
  'mrk': 'Mark',
  'luk': 'Luke',
  'jhn': 'John',
  'act': 'Acts',
  'rom': 'Romans',
  '1 cor': '1 Corinthians', '1cor': '1 Corinthians',
  '2 cor': '2 Corinthians', '2cor': '2 Corinthians',
  'gal': 'Galatians',
  'eph': 'Ephesians',
  'phil': 'Philippians', 'php': 'Philippians',
  'col': 'Colossians',
  '1 thess': '1 Thessalonians', '1thess': '1 Thessalonians', '1 thes': '1 Thessalonians',
  '2 thess': '2 Thessalonians', '2thess': '2 Thessalonians', '2 thes': '2 Thessalonians',
  '1 tim': '1 Timothy', '1tim': '1 Timothy',
  '2 tim': '2 Timothy', '2tim': '2 Timothy',
  'tit': 'Titus',
  'phlm': 'Philemon', 'phm': 'Philemon',
  'heb': 'Hebrews',
  'jas': 'James', 'jam': 'James',
  '1 pet': '1 Peter', '1pet': '1 Peter',
  '2 pet': '2 Peter', '2pet': '2 Peter',
  '1 jn': '1 John', '1jn': '1 John',
  '2 jn': '2 John', '2jn': '2 John',
  '3 jn': '3 John', '3jn': '3 John',
  'jud': 'Jude',
  'rev': 'Revelation', 'revelations': 'Revelation',
};

// Normalize a scripture reference to "Book Chapter:Verse" format
function normalizeRef(ref) {
  let s = ref.trim();
  // Replace dots used as chapter:verse separator (e.g. "Proverbs 16.6")
  s = s.replace(/(\d+)\.(\d)/g, '$1:$2');
  // "v." notation
  s = s.replace(/\bv\.?\s*(\d)/gi, ':$1');
  return s;
}

// Resolve a book name string to a canonical key in the JSON
function resolveBook(bookStr) {
  // Try exact match first
  if (bible[bookStr]) return bookStr;
  // Try alias lookup (lowercase)
  const lower = bookStr.toLowerCase().trim();
  if (BOOK_ALIASES[lower]) return BOOK_ALIASES[lower];
  // Try case-insensitive match against actual keys
  const keys = Object.keys(bible);
  const match = keys.find(k => k.toLowerCase() === lower);
  if (match) return match;
  return null;
}

// Parse "Book Chapter:StartVerse[-EndVerse]" or "Book Chapter" into parts.
// Returns { book, chapter, startVerse, endVerse } or throws.
function parseRef(ref) {
  const normalized = normalizeRef(ref);

  // Match: optional-number + spaces + book-name + space + chapter + optional(:verse[-verse])
  // e.g. "1 Corinthians 13:4-7", "Psalm 23", "John 3:16", "Proverbs 16:6"
  const m = normalized.match(/^((?:\d\s+)?\w[\w\s]*?)\s+(\d+)(?::(\d+)(?:[–\-](\d+))?)?$/);
  if (!m) throw new Error(`Cannot parse scripture reference: "${ref}"`);

  const bookStr = m[1].trim();
  const chapter = String(parseInt(m[2]));
  const startVerse = m[3] ? parseInt(m[3]) : null;
  const endVerse = m[4] ? parseInt(m[4]) : null;

  const book = resolveBook(bookStr);
  if (!book) throw new Error(`Unknown book: "${bookStr}" in reference "${ref}"`);

  return { book, chapter, startVerse, endVerse };
}

// Fetch verses from local JSON. Returns array of { verse, text }.
function fetchVerses(ref) {
  const { book, chapter, startVerse, endVerse } = parseRef(ref);
  const chapterData = bible[book] && bible[book][chapter];
  if (!chapterData) throw new Error(`Chapter not found: ${book} ${chapter}`);

  const allVerseNums = Object.keys(chapterData).map(Number).sort((a, b) => a - b);

  let verseNums;
  if (startVerse === null) {
    // Whole chapter
    verseNums = allVerseNums;
  } else if (endVerse === null) {
    // Single verse
    verseNums = allVerseNums.filter(n => n === startVerse);
  } else {
    // Range
    verseNums = allVerseNums.filter(n => n >= startVerse && n <= endVerse);
  }

  if (verseNums.length === 0) {
    throw new Error(`No verses found for "${ref}"`);
  }

  return verseNums.map(n => ({ verse: n, text: chapterData[String(n)] }));
}

// Split verses into slide-sized chunks (~10-12 lines each)
function splitIntoSlides(verses, ref, maxLinesPerSlide = 10) {
  const normalized = normalizeRef(ref);
  if (verses.length === 0) return [];

  const chunks = [];
  let current = [];

  for (const v of verses) {
    current.push(v);
    const lineCount = current.reduce((sum, v) => sum + Math.ceil(v.text.length / 70), 0);
    if (lineCount >= maxLinesPerSlide) {
      chunks.push(current);
      current = [];
    }
  }
  if (current.length > 0) chunks.push(current);

  return chunks.map((chunk) => {
    const firstVerse = chunk[0].verse;
    const lastVerse = chunk[chunk.length - 1].verse;
    const rangeLabel = chunks.length === 1
      ? normalized
      : `${normalized} (${firstVerse}${lastVerse !== firstVerse ? '–' + lastVerse : ''})`;
    const title = `${rangeLabel} NKJV`;
    return {
      title,
      lines: chunk.map(v => `${v.verse} ${v.text}`)
    };
  });
}

// Main export: fetch and split (sync under the hood, async interface for compatibility)
async function fetchScripture(ref) {
  const normalized = normalizeRef(ref);
  const verses = fetchVerses(normalized);
  return splitIntoSlides(verses, normalized);
}

module.exports = { fetchScripture, normalizeRef };
