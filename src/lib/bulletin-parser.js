'use strict';

const mammoth = require('mammoth');
const cheerio = require('cheerio');

// Known responsive reading section headings
const READING_HEADINGS = [
  'covenant promises',
  'call to worship',
  'reading of the law',
  'reading of the law & confession of sin',
  'law & confession',
  'reading of the gospel',
  'gospel reading',
  'nicene creed',
  "apostles' creed",
  "lord's prayer",
];

// Patterns that indicate a song line (has page reference)
const SONG_PAGE_RE = /\(pages?\s*[\d,\s–-]+\)/i;

// Strip ∆, (seated), and (page N) from song title; keep all other parentheticals
function extractSongTitle(raw) {
  let s = raw;
  s = s.replace(/^∆\s*/, '');
  s = s.replace(/^\(seated\)\s*/i, '');
  s = s.replace(/\(pages?\s*[\d,\s–-]+\)/gi, '');
  return s.trim();
}

// Check if a line looks like a song entry
function looksLikeSong(line) {
  return SONG_PAGE_RE.test(line);
}

// Parse <em> italic runs from a HTML segment.
// Returns [{text, italic}] if any italic content is present, otherwise null.
function htmlToRuns(segHtml) {
  if (!segHtml || !/<em\b/.test(segHtml)) return null;
  const runs = [];
  let idx = 0;
  const re = /<em\b[^>]*>([\s\S]*?)<\/em>/gi;
  let m;
  while ((m = re.exec(segHtml)) !== null) {
    if (m.index > idx) {
      const t = cheerio.load(`<span>${segHtml.slice(idx, m.index)}</span>`)('span').text();
      if (t) runs.push({ text: t, italic: false });
    }
    const t = cheerio.load(`<span>${m[1]}</span>`)('span').text();
    if (t) runs.push({ text: t, italic: true });
    idx = m.index + m[0].length;
  }
  if (idx < segHtml.length) {
    const t = cheerio.load(`<span>${segHtml.slice(idx)}</span>`)('span').text();
    if (t) runs.push({ text: t, italic: false });
  }
  return runs.length > 0 ? runs : null;
}

// Parse Leader/All lines from a block.
// lines: array of { text: string, bold: boolean, html?: string } objects (or plain strings for compat)
// Role inference: explicit (Leader)/(All) markers take priority;
// otherwise bold → All, normal → Leader.
// Returns [{role, text, runs}] where runs is [{text, italic}]|null.
function parseLeaderAllLines(lines) {
  const result = [];
  let i = 0;

  while (i < lines.length) {
    const item = lines[i];
    const lineText = (typeof item === 'string' ? item : item.text).trim();
    const isBold   = typeof item === 'string' ? undefined : item.bold;
    const itemHtml = typeof item === 'object'  ? item.html  : null;

    if (!lineText) { i++; continue; }

    let role, text, runs;

    if (/^\(Leader\)/i.test(lineText)) {
      role = 'Leader';
      text = lineText.replace(/^\(Leader\)\s*/i, '').trim();
      runs = htmlToRuns(itemHtml);
      if (runs) {
        // Strip "(Leader)" prefix text from first run
        if (runs.length > 0 && /^\(Leader\)\s*/i.test(runs[0].text)) {
          runs[0] = { ...runs[0], text: runs[0].text.replace(/^\(Leader\)\s*/i, '') };
          if (!runs[0].text.trim()) runs.shift();
        }
      }
      i++;
    } else if (/^\(All\)/i.test(lineText)) {
      role = 'All';
      text = lineText.replace(/^\(All\)\s*/i, '').trim();
      runs = htmlToRuns(itemHtml);
      if (runs) {
        if (runs.length > 0 && /^\(All\)\s*/i.test(runs[0].text)) {
          runs[0] = { ...runs[0], text: runs[0].text.replace(/^\(All\)\s*/i, '') };
          if (!runs[0].text.trim()) runs.shift();
        }
      }
      i++;
    } else {
      // Infer role from bold formatting
      role = isBold ? 'All' : 'Leader';
      text = lineText;
      runs = htmlToRuns(itemHtml);
      i++;
    }

    // Collect continuation lines: no explicit marker AND same bold status
    while (i < lines.length) {
      const nextItem = lines[i];
      const nextText = (typeof nextItem === 'string' ? nextItem : nextItem.text).trim();
      const nextBold = typeof nextItem === 'string' ? undefined : nextItem.bold;
      const nextHtml = typeof nextItem === 'object'  ? nextItem.html  : null;

      if (!nextText) { i++; continue; }

      // Stop at explicit role marker
      if (/^\(All\)/i.test(nextText) || /^\(Leader\)/i.test(nextText)) break;

      // Stop when bold status changes (indicates new role)
      if (nextBold !== undefined && isBold !== undefined && nextBold !== isBold) break;

      text += ' ' + nextText;
      if (runs !== null) {
        const nextRuns = htmlToRuns(nextHtml);
        if (nextRuns) {
          runs = [...runs, { text: ' ', italic: false }, ...nextRuns];
        } else {
          runs = [...runs, { text: ' ' + nextText, italic: false }];
        }
      }
      i++;
    }

    if (text) result.push({ role, text, runs: runs || null });
  }

  return result;
}

// Check if a heading matches a known reading type
function classifyHeading(heading) {
  const lower = heading.toLowerCase();
  if (lower.includes('covenant promises')) return 'Covenant Promises';
  if (lower.includes('call to worship')) return 'Call to Worship';
  if (lower.includes('reading of the law') || lower.includes('law & confession') || lower.includes('confession of sin')) return 'Reading of the Law & Confession of Sin';
  if (lower.includes('reading of the gospel') || lower === 'gospel') return 'Reading of the Gospel';
  if (/heidelberg catechism/i.test(lower)) {
    const qMatch = heading.match(/Q\s*[&]\s*A\s*(\d+)/i) || heading.match(/Q\s*(\d+)/i) || heading.match(/#\s*(\d+)/);
    const num = qMatch ? qMatch[1] : '';
    return `Heidelberg Catechism #${num}`;
  }
  if (/^nicene creed/.test(lower)) return 'Nicene Creed';
  if (/^apostles\W\s*creed/.test(lower)) return "Apostles' Creed";
  if (/^lord\W?s?\s+prayer/.test(lower)) return "Lord's Prayer";
  return null;
}

// Parse scripture reference from "Communion Meditation: Psalm 21.6" etc.
function parseScriptureRef(line) {
  // Strip trailing speaker/annotation from a scripture ref line
  function cleanRef(raw) {
    return raw
      .replace(/\s*[–—]\s*.*$/, '')              // strip em/en-dash + anything after
      .replace(/\s+-\s*[A-Z].*$/, '')             // strip " - Pastor/Elder Name"
      .replace(/-[A-Z][a-z].*$/, '')              // strip "-Pastor" with no space
      .replace(/,\s+[A-Z][a-z].*$/, '')           // strip ", Walking Uprightly..." (sermon title)
      .replace(/\s*\(.*$/, '')                    // strip " (below)" etc.
      .trim();
  }

  // Communion Meditation: ref
  let m = line.match(/Communion Meditation:\s*(.+)/i);
  if (m) return normalizeRef(cleanRef(m[1]));

  // Scripture Reading: ref
  m = line.match(/Scripture Reading:\s*(.+)/i);
  if (m) return normalizeRef(cleanRef(m[1]));

  return null;
}

function normalizeRef(ref) {
  let s = ref.trim();
  // Expand common abbreviations
  s = s.replace(/^Ps\.\s*/i, 'Psalm ');
  s = s.replace(/^Prov\.\s*/i, 'Proverbs ');
  s = s.replace(/^Gen\.\s*/i, 'Genesis ');
  s = s.replace(/^Ex\.\s*/i, 'Exodus ');
  s = s.replace(/^Deut\.\s*/i, 'Deuteronomy ');
  s = s.replace(/^Rom\.\s*/i, 'Romans ');
  s = s.replace(/^Eph\.\s*/i, 'Ephesians ');
  s = s.replace(/^Phil\.\s*/i, 'Philippians ');
  s = s.replace(/^Col\.\s*/i, 'Colossians ');
  // Replace dots used as chapter:verse separator: "16.6" → "16:6"
  s = s.replace(/(\d+)\.(\d)/g, '$1:$2');
  return s.trim();
}

async function parseBulletin(docxPath) {
  const result = await mammoth.convertToHtml({ path: docxPath });
  const html = result.value;
  const $ = cheerio.load(html);

  // Extract paragraphs with bold info from HTML.
  // Split each <p> on <br> tags so that mixed-bold paragraphs (e.g. an All line
  // followed by a Leader line within the same <p>) produce separate entries.
  const paragraphs = [];
  $('p').each((_, el) => {
    const innerHTML = $(el).html() || '';
    // Protect <br> tags that fall inside <strong>...</strong> from being treated as
    // segment boundaries — otherwise the bold/non-bold detection breaks because the
    // <strong> pair gets split across two segments and neither matches /<strong>...<\/strong>/.
    const BR_PLACEHOLDER = '\x01';
    const protectedHtml = innerHTML.replace(/<strong\b[^>]*>[\s\S]*?<\/strong>/gi,
      m => m.replace(/<br\s*\/?>/gi, BR_PLACEHOLDER));
    const segments = protectedHtml.split(/<br\s*\/?>/i);
    for (const rawSegHtml of segments) {
      const segHtml = rawSegHtml.replace(/\x01/g, '<br/>');
      const fullText = cheerio.load(`<span>${segHtml}</span>`)('span').text()
        .replace(/\s+/g, ' ').trim()
        .replace(/^[•·◆◇▪▸►‣⁃–—]\s*/, '');
      if (!fullText) continue;

      // Count characters inside <strong> tags for bold detection
      let boldLen = 0;
      const strongRe = /<strong\b[^>]*>([\s\S]*?)<\/strong>/gi;
      let m;
      while ((m = strongRe.exec(segHtml)) !== null) {
        boldLen += m[1].replace(/<[^>]+>/g, '').length;
      }
      const isBold = boldLen / fullText.length > 0.5;
      paragraphs.push({ text: fullText, bold: isBold, html: segHtml });
    }
  });

  const items = [];
  let i = 0;

  // State for reading collection
  let currentReadingHeading = null;
  let currentReadingLines = []; // array of { text, bold }

  function flushReading() {
    if (currentReadingHeading && currentReadingLines.length > 0) {
      const parsed = parseLeaderAllLines(currentReadingLines);
      if (parsed.length > 0) {
        items.push({
          type: 'reading',
          title: currentReadingHeading,
          lines: parsed
        });
      }
      currentReadingLines = [];
      currentReadingHeading = null;
    }
  }

  while (i < paragraphs.length) {
    const para = paragraphs[i];
    const line = para.text;

    // Check for Communion Meditation scripture
    if (/^Communion Meditation:/i.test(line)) {
      flushReading();
      const ref = parseScriptureRef(line);
      if (ref) {
        items.push({ type: 'scripture', subtype: 'communion', ref });
      }
      i++;
      continue;
    }

    // Check for Scripture Reading
    if (/^∆?\s*Scripture Reading:/i.test(line)) {
      flushReading();
      const ref = parseScriptureRef(line);
      if (ref) {
        items.push({ type: 'scripture', subtype: 'reading', ref });
      }
      i++;
      continue;
    }

    // Check for Heidelberg Catechism
    if (/^Heidelberg Catechism Q/i.test(line)) {
      flushReading();
      const heading = classifyHeading(line);

      // Split the paragraph into question (non-bold → Leader) and answer (bold → All).
      // Walk the HTML collecting non-strong text as leader, strong text as answer.
      const segHtml = para.html || '';
      const nonBoldParts = [];
      const boldParts = [];
      let pos = 0;
      const strongRe2 = /<strong\b[^>]*>([\s\S]*?)<\/strong>/gi;
      let sm2;
      while ((sm2 = strongRe2.exec(segHtml)) !== null) {
        const beforeText = cheerio.load(`<span>${segHtml.slice(pos, sm2.index)}</span>`)('span').text()
          .replace(/\s+/g, ' ').trim();
        if (beforeText) nonBoldParts.push(beforeText);
        const boldText = cheerio.load(`<span>${sm2[1]}</span>`)('span').text()
          .replace(/\s+/g, ' ').trim();
        if (boldText) boldParts.push(boldText);
        pos = sm2.index + sm2[0].length;
      }
      const trailingText = cheerio.load(`<span>${segHtml.slice(pos)}</span>`)('span').text()
        .replace(/\s+/g, ' ').trim();
      if (trailingText) nonBoldParts.push(trailingText);

      let leaderText = nonBoldParts.join(' ').replace(/^Heidelberg Catechism\s*/i, '')
        .replace(/^Q\s*[&+]\s*A\s*\d+\s*:\s*/i, '')
        .replace(/\s+/g, ' ').trim();
      const inlineAnswer = boldParts.join(' ');

      // If the heading paragraph only contains the Q&A label ("Q & A 72" with no question),
      // don't emit it as a leader line — the question comes from continuation paragraphs.
      if (/^Q\s*[&+]\s*A\s*\d+\s*:?\s*$/.test(leaderText)) leaderText = '';

      // Collect continuation paragraphs; use standard bold=All / non-bold=Leader logic.
      i++;
      const continuationLines = [];
      while (i < paragraphs.length) {
        const nextPara = paragraphs[i];
        const next = nextPara.text;
        if (!next) { i++; continue; }
        if (/^(∆\s+|Prayer |Prelude |Welcome|Introduction|Affiliated|Meeting|CCLI|Announcements|Benediction)/i.test(next)) break;
        if (looksLikeSong(next)) break;
        if (/^Communion Meditation:/i.test(next)) break;
        if (/^∆?\s*Scripture Reading:/i.test(next)) break;
        if (/^Heidelberg Catechism Q/i.test(next)) break;
        if (classifyHeading(next) !== null) break;
        continuationLines.push(nextPara);
        i++;
      }

      const contParsed = parseLeaderAllLines(continuationLines);
      const lines = [];
      if (leaderText) lines.push({ role: 'Leader', text: leaderText });
      if (inlineAnswer) lines.push({ role: 'All', text: inlineAnswer });
      for (const l of contParsed) lines.push(l);
      if (lines.length > 0) {
        items.push({ type: 'reading', title: heading, lines });
      }
      continue;
    }

    // If we're in a reading section, accumulate lines
    if (currentReadingHeading) {
      // Check if this looks like a song (would end the reading)
      if (looksLikeSong(line) && (line.startsWith('∆') || /^\(seated\)/i.test(line))) {
        flushReading();
        // Fall through to song handling below
      } else if (/^∆\s*(Prayer|Benediction|Reading of)/i.test(line)) {
        // Sub-header like "∆ Reading of the Covenant Promises" — discard junk, keep heading
        currentReadingLines = [];
        i++;
        continue;
      } else {
        // Check if a truly different section heading appears while inside a reading
        const newHeading = classifyHeading(line);
        if (newHeading && newHeading !== currentReadingHeading) {
          flushReading();
          currentReadingHeading = newHeading;
          i++;
          continue;
        }
        // Stop accumulating on service-annotation and copyright lines
        if (/^(Time of |Prayer of |Assurance of |Silent Confession|CCLI|©|Copyright)/i.test(line)) {
          flushReading();
          i++;
          continue;
        }
        currentReadingLines.push(para);
        i++;
        continue;
      }
    }

    // Check for known reading headings (section starts)
    const knownHeading = classifyHeading(line);
    if (knownHeading) {
      flushReading();
      // If the paragraph has body text beyond a bold heading (e.g. "Nicene Creed: I believe..."),
      // emit it immediately as a single-All reading rather than starting an accumulation section.
      // Only trigger when <strong> is present — plain-text headings always start an accumulation.
      const segHtml = para.html || '';
      const hasStrongHeading = /<strong\b/.test(segHtml);
      if (hasStrongHeading) {
        const afterStrongHtml = segHtml.replace(/<strong\b[^>]*>[\s\S]*?<\/strong>/gi, '');
        const bodyText = cheerio.load(`<span>${afterStrongHtml}</span>`)('span').text()
          .replace(/\s+/g, ' ').replace(/^[:\s]+/, '').trim();
        if (bodyText.length > 10) {
          items.push({ type: 'reading', title: knownHeading, lines: [{ role: 'All', text: bodyText }] });
          i++;
          continue;
        }
      }
      currentReadingHeading = knownHeading;
      i++;
      continue;
    }

    // Songs: ∆ Title (page N) or (seated) Title (page N)
    if (looksLikeSong(line)) {
      const isSitting = /^\(seated\)/i.test(line);
      const title = extractSongTitle(line);
      if (title) {
        items.push({
          type: 'song',
          title,
          rawTitle: line,
          sitting: isSitting
        });
      }
      i++;
      continue;
    }

    i++;
  }

  flushReading();

  return items;
}

module.exports = { parseBulletin, normalizeRef };
