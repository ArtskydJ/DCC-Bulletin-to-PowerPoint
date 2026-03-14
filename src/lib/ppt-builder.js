'use strict';

const AdmZip = require('adm-zip');
const path = require('path');
const fs = require('fs');

// EMU measurements (English Metric Units)
// Slide dimensions: 12192000 x 6858000 EMU (widescreen 16:9)
const SLIDE_W = 12192000;
const SLIDE_H = 6858000;

// Title shape position/size (from reference)
const TITLE_X = 763588;
const TITLE_Y = 0;
const TITLE_W = 11476037;
const TITLE_H = 844550;

// Body shape position/size (from reference)
const BODY_X = 627063;
const BODY_Y = 1004888;
const BODY_W = 10793412;
const BODY_H = 6424836;

// Hanging indent for Leader/All (from reference)
const HANGING_MAR = 1781175;
const HANGING_INDENT = -1781175;

// Normalized font sizes applied to every slide (half-points: 100 = 1pt)
const NORM_TITLE_SZ     = 3400; // 34pt — title placeholder (ph=title/ctrTitle)
const NORM_BODY_SZ      = 3800; // 38pt — lyric/content text boxes (ph=none)
const NORM_CREDITS_SZ   = 1800; // 18pt — credits/subtitle placeholder (ph=body)

// Space before a paragraph when the speaker role changes (Leader↔All), in half-points
const READING_ROLE_GAP = 1600; // 16pt

// Normalized Latin font — replaces any <a:latin> element on every slide
const NORM_LATIN = '<a:latin typeface="Arial Narrow" panose="020B0606020202030204" pitchFamily="34" charset="0"/>';
// Normalized text fill — explicit white, since background is always near-black (#010101)
const NORM_FILL  = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>';
// Slide background — #010101 instead of pure black so Zoom detects slide changes
const SLIDE_BG = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="010101"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';

// Standard xfrm strings for position normalization
const XFRM_TITLE = `<a:xfrm><a:off x="${TITLE_X}" y="${TITLE_Y}"/><a:ext cx="${TITLE_W}" cy="${TITLE_H}"/></a:xfrm>`;
const XFRM_BODY  = `<a:xfrm><a:off x="${BODY_X}" y="${BODY_Y}"/><a:ext cx="${BODY_W}" cy="${BODY_H}"/></a:xfrm>`;

// Normalize font sizes, typeface, colors, and optionally bounding-box positions.
// Three size tiers: title → NORM_TITLE_SZ, credits → NORM_CREDITS_SZ,
// all other shapes (lyric text boxes) → NORM_BODY_SZ.
// Credits shape is the lowest-bottomed non-title, non-placeholder text box on the slide,
// detected only when there are ≥2 such boxes (i.e. a body box AND a credits box).
// All <a:latin> replaced with Arial Narrow; all run fills replaced with white.
// Runs with a non-zero baseline (superscript/subscript) keep their original sz.
// normalizePositions: only set true for slides we generate (readings/scripture).
//   Song slides keep their original text-box geometry.
function normalizeSlide(slideXml, normalizePositions = false) {
  // Pre-scan: for song slides, find the credits shape (lowest non-title text box
  // when ≥2 non-title shapes are present — body + credits).
  let creditsShapeId = null;
  if (!normalizePositions) {
    const bodyShapes = [];
    for (const m of slideXml.matchAll(/<p:sp\b([\s\S]*?)<\/p:sp>/g)) {
      const s = m[1];
      if (/<p:ph\s[^>]*type="(title|ctrTitle)"/.test(s)) continue; // skip title placeholders only
      const id = (s.match(/id="(\d+)"/) || [, ''])[1];
      const off = s.match(/<a:off x="(\d+)" y="(\d+)"/);
      const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
      if (id && off && ext) {
        bodyShapes.push({ id, bottom: parseInt(off[2]) + parseInt(ext[2]) });
      }
    }
    if (bodyShapes.length >= 2) {
      bodyShapes.sort((a, b) => b.bottom - a.bottom);
      creditsShapeId = bodyShapes[0].id;
    }
  }

  return slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, shape => {
    const isTitle     = /<p:ph\s[^>]*type="(title|ctrTitle)"/.test(shape);
    const isSubTitle  = /<p:ph\s[^>]*type="subTitle"/.test(shape);
    const shapeId     = (shape.match(/id="(\d+)"/) || [, ''])[1];
    const isCredits = isSubTitle || (creditsShapeId !== null && shapeId === creditsShapeId);
    const targetSz    = isTitle ? NORM_TITLE_SZ : isCredits ? NORM_CREDITS_SZ : NORM_BODY_SZ;

    // 0. Normalize bounding-box position
    if (!isCredits) {
      if (normalizePositions) {
        // Readings/scripture: snap to exact reference coordinates
        const targetXfrm = isTitle ? XFRM_TITLE : XFRM_BODY;
        if (/<a:xfrm[\s>]/.test(shape)) {
          shape = shape.replace(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/, targetXfrm);
        } else {
          shape = shape.replace(/(<p:spPr\b[^>]*>)/, `$1${targetXfrm}`);
        }
      // Song slides: leave all shape positions as-is from the source file.
    }
    }

    // 0b. Inject <a:rPr> into bare runs that have none — they'd otherwise inherit
    //     the theme default which may be black on light-themed source files.
    shape = shape.replace(/<a:r>(<a:t>)/g,
      `<a:r><a:rPr lang="en-US" sz="${targetSz}" dirty="0">${NORM_FILL}${NORM_LATIN}</a:rPr>$1`);

    return shape
      // 1. Normalize sz in all run-property opening tags
      .replace(/(<a:(?:rPr|endParaRPr|defRPr)\b)([^>]*>)/g, (match, tag, rest) => {
        if (/\bbaseline="-?[1-9]\d*"/.test(rest)) return match; // preserve sub/superscript
        if (/\bsz="\d+"/.test(rest)) {
          return tag + rest.replace(/\bsz="\d+"/, `sz="${targetSz}"`);
        }
        return tag + rest.replace(/(\/?>)$/, ` sz="${targetSz}"$1`);
      })
      // 2. Replace existing <a:latin> elements with the normalized font
      .replace(/<a:latin\b[^/]*\/>/g, NORM_LATIN)
      // 3. Self-closing <a:rPr/> — convert to open element and inject font + fill
      .replace(/<(a:(?:rPr|endParaRPr|defRPr))(\b[^>]*)\/>/g,
        (match, tag, attrs) => `<${tag}${attrs}>${NORM_FILL}${NORM_LATIN}</${tag}>`)
      // 4. Open <a:rPr>...</a:rPr> — normalize fill and inject font if missing
      .replace(/(<a:(?:rPr|endParaRPr|defRPr)\b[^>]*>)([\s\S]*?)(<\/a:(?:rPr|endParaRPr|defRPr)>)/g,
        (match, open, content, close) => {
          // Normalize fill: replace any existing solidFill, or add if missing
          let c = content.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, NORM_FILL);
          if (!/<a:solidFill/.test(c)) c = NORM_FILL + c;
          // Inject font if missing
          if (!/<a:latin\b/.test(c)) c = c + NORM_LATIN;
          return open + c + close;
        })
      // 5. Schema order inside <a:pPr>: lnSpc → spcBef → spcAft → buNone → rest
      .replace(/<a:pPr(\b[^>]*)>([\s\S]*?)<\/a:pPr>/g, (_, attrs, inner) => {
          let stripped = inner
            .replace(/<a:buNone\/>/g, '')
            .replace(/<a:buChar\b[^/]*\/>/g, '')
            .replace(/<a:buFont\b[^/]*\/>/g, '')
            .replace(/<a:buAutoNum\b[^/]*\/>/g, '')
            .replace(/<a:buClrTx\/>/g, '')
            .replace(/<a:buSzTx\/>/g, '')
            .replace(/<a:buSzPct\b[^/]*\/>/g, '')
            .replace(/<a:buSzPts\b[^/]*\/>/g, '')
            .replace(/<a:buClr\b[\s\S]*?<\/a:buClr>/g, '');
          // Pull spcBef/spcAft out of stripped so they go before buNone
          const spcBef = (stripped.match(/<a:spcBef>[\s\S]*?<\/a:spcBef>/) || [''])[0];
          const spcAft = (stripped.match(/<a:spcAft>[\s\S]*?<\/a:spcAft>/) || [''])[0];
          const rest = stripped
            .replace(/<a:spcBef>[\s\S]*?<\/a:spcBef>/, '')
            .replace(/<a:spcAft>[\s\S]*?<\/a:spcAft>/, '');
          return `<a:pPr${attrs}>${spcBef}${spcAft}<a:buNone/>${rest}</a:pPr>`;
        })
      .replace(/<a:pPr(\b[^>]*)\/>/g, `<a:pPr$1><a:buNone/></a:pPr>`);
  })
  // Strip slide transitions — self-closing or element-with-children
  .replace(/<p:transition\b[^>]*(?:\/>|>[\s\S]*?<\/p:transition>)/g, '');
}

// Normalization for song slides: font sizes, white text, title font/position, credits bullets/bold.
// Positions and typeface of body shapes are left exactly as in the source.
function normalizeSongSlide(slideXml) {
  // Credits detection: lowest non-title shape with an explicit position.
  let creditsShapeId = null;
  const bodyShapes = [];
  for (const m of slideXml.matchAll(/<p:sp\b([\s\S]*?)<\/p:sp>/g)) {
    const s = m[1];
    if (/<p:ph\s[^>]*type="(title|ctrTitle)"/.test(s)) continue;
    const id = (s.match(/id="(\d+)"/) || [, ''])[1];
    const off = s.match(/<a:off x="(\d+)" y="(\d+)"/);
    const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
    if (id && off && ext) bodyShapes.push({ id, bottom: parseInt(off[2]) + parseInt(ext[2]) });
  }
  if (bodyShapes.length >= 2) {
    bodyShapes.sort((a, b) => b.bottom - a.bottom);
    creditsShapeId = bodyShapes[0].id;
  }

  return slideXml
    .replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, shape => {
      const isTitle     = /<p:ph\s[^>]*type="(title|ctrTitle)"/.test(shape);
      const isSubTitle  = /<p:ph\s[^>]*type="subTitle"/.test(shape);
      const shapeId     = (shape.match(/id="(\d+)"/) || [, ''])[1];
      const isCredits = isSubTitle || (creditsShapeId !== null && shapeId === creditsShapeId);
      const targetSz    = isTitle ? NORM_TITLE_SZ : isCredits ? NORM_CREDITS_SZ : NORM_BODY_SZ;

      // Pin title y=0 and vertical alignment top — prevents title jumping between source files
      if (isTitle) {
        shape = shape.replace(/(<a:off\s+x="[^"]*")\s+y="\d+"/, '$1 y="0"');
        shape = shape.replace(/<a:bodyPr\b([^>]*?)(\/?)>/g, (_, attrs, slash) => {
          const a = /\banchor=/.test(attrs)
            ? attrs.replace(/\banchor="[^"]*"/, 'anchor="t"')
            : attrs + ' anchor="t"';
          return `<a:bodyPr${a}${slash}>`;
        });
      }

      // Inject rPr into bare <a:r> runs (no existing rPr) — catches theme-colored/unformatted text
      const injectAttrs   = isCredits ? ' b="0" i="1"' : '';
      const injectContent = `${NORM_FILL}${isTitle ? NORM_LATIN : ''}`;
      shape = shape.replace(/<a:r>(<a:t>)/g,
        `<a:r><a:rPr lang="en-US" sz="${targetSz}"${injectAttrs} dirty="0">${injectContent}</a:rPr>$1`);

      // Normalize font size (and unbold credits) on all run-property opening tags
      shape = shape.replace(/(<a:(?:rPr|endParaRPr|defRPr)\b)([^>]*>)/g, (match, tag, rest) => {
        if (/\bbaseline="-?[1-9]\d*"/.test(rest)) return match; // preserve superscript
        let r = rest;
        if (/\bsz="\d+"/.test(r)) r = r.replace(/\bsz="\d+"/, `sz="${targetSz}"`);
        else r = r.replace(/(\/?>)$/, ` sz="${targetSz}"$1`);
        if (isCredits) {
          if (/\bb="[^"]*"/.test(r)) r = r.replace(/\bb="[^"]*"/, 'b="0"');
          else r = r.replace(/(\/?>)$/, ` b="0"$1`);
          if (/\bi="[^"]*"/.test(r)) r = r.replace(/\bi="[^"]*"/, 'i="1"');
          else r = r.replace(/(\/?>)$/, ` i="1"$1`);
        }
        return tag + r;
      });

      // Replace <a:latin> in title shapes with normalized font
      if (isTitle) {
        shape = shape.replace(/<a:latin\b[^/]*\/>/g, NORM_LATIN);
      }

      // Self-closing <a:rPr/> → open element with white fill (+ font for title)
      shape = shape.replace(/<(a:(?:rPr|endParaRPr|defRPr))(\b[^>]*)\/>/g,
        (match, tag, attrs) => `<${tag}${attrs}>${NORM_FILL}${isTitle ? NORM_LATIN : ''}</${tag}>`);

      // Open <a:rPr>...</a:rPr> → inject/replace fill with white (+ font for title)
      shape = shape.replace(/(<a:(?:rPr|endParaRPr|defRPr)\b[^>]*>)([\s\S]*?)(<\/a:(?:rPr|endParaRPr|defRPr)>)/g,
        (match, open, content, close) => {
          let c = content.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, NORM_FILL);
          if (!/<a:solidFill/.test(c)) c = NORM_FILL + c;
          if (isTitle && !/<a:latin\b/.test(c)) c += NORM_LATIN;
          return open + c + close;
        });

      // Remove bullet formatting from credits shapes
      if (isCredits) {
        shape = shape.replace(/<a:pPr(\b[^>]*)>([\s\S]*?)<\/a:pPr>/g, (_, attrs, inner) => {
          let s2 = inner
            .replace(/<a:buNone\/>/g, '')
            .replace(/<a:buChar\b[^/]*\/>/g, '')
            .replace(/<a:buFont\b[^/]*\/>/g, '')
            .replace(/<a:buAutoNum\b[^/]*\/>/g, '')
            .replace(/<a:buClrTx\/>/g, '')
            .replace(/<a:buSzTx\/>/g, '')
            .replace(/<a:buSzPct\b[^/]*\/>/g, '')
            .replace(/<a:buSzPts\b[^/]*\/>/g, '')
            .replace(/<a:buClr\b[\s\S]*?<\/a:buClr>/g, '');
          const spcBef = (s2.match(/<a:spcBef>[\s\S]*?<\/a:spcBef>/) || [''])[0];
          const spcAft = (s2.match(/<a:spcAft>[\s\S]*?<\/a:spcAft>/) || [''])[0];
          const rest2  = s2
            .replace(/<a:spcBef>[\s\S]*?<\/a:spcBef>/, '')
            .replace(/<a:spcAft>[\s\S]*?<\/a:spcAft>/, '');
          return `<a:pPr${attrs}>${spcBef}${spcAft}<a:buNone/>${rest2}</a:pPr>`;
        });
        shape = shape.replace(/<a:pPr(\b[^>]*)\/>/g, `<a:pPr$1><a:buNone/></a:pPr>`);
      }

      return shape;
    })
    .replace(/<p:transition\b[^>]*(?:\/>|>[\s\S]*?<\/p:transition>)/g, '');
}

function xmlEscape(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Build a run (a:r) with given properties
function makeRun(text, { bold = false, sz = NORM_BODY_SZ, italic = false } = {}) {
  const bAttr = bold   ? '1' : '0';
  const iAttr = italic ? ` i="1"` : '';
  return `<a:r><a:rPr lang="en-US" sz="${sz}" b="${bAttr}"${iAttr} dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial Narrow" panose="020B0606020202030204" pitchFamily="34" charset="0"/></a:rPr><a:t>${xmlEscape(text)}</a:t></a:r>`;
}

// Build a paragraph for a reading line (Leader: or All:).
// lineObj: { role, text, runs } — runs is [{text, italic}] or null.
// prevRole: role of the preceding paragraph (null for first).
function makeReadingParagraph(lineObj, prevRole = null) {
  const role = lineObj.role;
  const text = lineObj.text;
  const runs = lineObj.runs; // [{text, italic}] | null
  const bold = role === 'All';
  const roleChanged = prevRole !== null && prevRole !== role;
  const spaceBefore = roleChanged ? `<a:spcBef><a:spcPts val="${READING_ROLE_GAP}"/></a:spcBef>` : '';
  const label = role + ':\t';
  const contentRuns = runs
    ? runs.map(r => makeRun(r.text, { bold, italic: r.italic })).join('')
    : makeRun(text, { bold });
  return `<a:p><a:pPr marL="${HANGING_MAR}" marR="0" indent="${HANGING_INDENT}">${spaceBefore}<a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>${makeRun(label, { bold })}${contentRuns}</a:p>`;
}

// Build a plain text paragraph (for scripture lines)
// Splits a leading verse number ("10 text...") into a superscript run so it stays
// visually small even after font-size normalization.
function makePlainParagraph(text, { sz = NORM_BODY_SZ, bold = false, first = false } = {}) {
  const spaceBefore = first ? '' : `<a:spcBef><a:spcPts val="800"/></a:spcBef>`;
  const verseMatch = text.match(/^(\d+)\s+([\s\S]*)$/);
  if (verseMatch) {
    const [, num, body] = verseMatch;
    // baseline="15000" = 15% superscript; sz preserved by normalizeSlide's baseline exception
    const numRun = `<a:r><a:rPr lang="en-US" sz="${NORM_BODY_SZ}" b="0" baseline="15000" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>${NORM_LATIN}</a:rPr><a:t>${xmlEscape(num)} </a:t></a:r>`;
    return `<a:p><a:pPr>${spaceBefore}</a:pPr>${numRun}${makeRun(body, { bold, sz })}</a:p>`;
  }
  return `<a:p><a:pPr>${spaceBefore}</a:pPr>${makeRun(text, { bold, sz })}</a:p>`;
}

// Build the title text body XML
function makeTitleBody(titleText, sz = NORM_TITLE_SZ) {
  return `<p:txBody><a:bodyPr lIns="48768" tIns="48768" rIns="48768" bIns="48768" anchor="t"/><a:lstStyle/><a:p><a:pPr defTabSz="1169988" eaLnBrk="1"/>${makeRun(titleText, { bold: false, sz })}</a:p></p:txBody>`;
}

// Split reading lines into slide-sized groups.
// Prefers splitting after an All line and before a Leader line.
// If forced to split mid-section, appends "…" to the last line.
const READING_CHAR_LIMIT = 450;
function splitReadingLines(lines) {
  // Find the last sentence boundary (./?/;) at or before maxLen that is NOT inside
  // parentheses. Returns the cut position (after the punctuation), or maxLen as fallback.
  function lastSentenceBoundary(text, maxLen) {
    let last = -1;
    let depth = 0;
    for (let i = 0; i < Math.min(text.length, maxLen); i++) {
      const c = text[i];
      if (c === '(') depth++;
      else if (c === ')') depth = Math.max(0, depth - 1);
      else if (depth === 0 && (c === '.' || c === '?' || c === ';')) last = i + 1;
    }
    return last > 0 ? last : maxLen;
  }

  // Step 1: Expand any individual line whose text alone would overflow the limit,
  // splitting at sentence boundaries (.  ?  ;) and marking the fragment as continued.
  // runs are only carried on the final (unsplit) fragment; split fragments lose italic info.
  const flat = [];
  for (const line of lines) {
    let text = line.text;
    let wasSplit = false;
    while (text.length > READING_CHAR_LIMIT) {
      wasSplit = true;
      const cutAt = lastSentenceBoundary(text, READING_CHAR_LIMIT - 1); // room for ellipsis
      flat.push({ role: line.role, text: text.slice(0, cutAt).trimEnd(), runs: null, cont: true });
      text = text.slice(cutAt).trimStart();
    }
    flat.push({ role: line.role, text, runs: line.runs || null, cont: false });
  }

  // Step 2: Greedily pack lines into slides.
  const slides = [];
  let cur = [];
  const curLen = () => cur.reduce((s, l) => s + l.text.length, 0);

  for (const line of flat) {
    const addLen = line.text.length;
    if (cur.length === 0 || curLen() + addLen <= READING_CHAR_LIMIT) {
      cur.push(line);
      continue;
    }
    // Doesn't fit. Prefer a natural break: last All→Leader boundary in cur.
    let nb = -1;
    for (let j = cur.length - 2; j >= 0; j--) {
      if (cur[j].role === 'All' && cur[j + 1].role === 'Leader') { nb = j; break; }
    }
    if (nb >= 0) {
      const overflow = cur.splice(nb + 1);
      slides.push(cur);
      cur = [...overflow, line];
    } else {
      slides.push(cur);
      cur = [line];
    }
  }
  if (cur.length > 0) slides.push(cur);

  // Step 3: Add [Continued...] where a slide ends mid-section.
  return slides.map((slide, si) => {
    if (si === slides.length - 1) return slide.map(({ role, text, runs }) => ({ role, text, runs: runs || null }));
    const last = slide[slide.length - 1];
    const nextFirst = slides[si + 1][0];
    const needsCont = last.cont || last.role === nextFirst.role;
    return slide.map((l, li) => {
      const isCont = needsCont && li === slide.length - 1;
      return {
        role: l.role,
        text: isCont ? l.text + '…' : l.text,
        runs: isCont ? null : (l.runs || null), // drop runs when appending [Continued...]
      };
    });
  });
}

// Build a reading/catechism/creed slide XML
function buildReadingSlideXml(title, lines) {
  const bodyParas = lines.map((l, i) => makeReadingParagraph(l, i === 0 ? null : lines[i - 1].role)).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${SLIDE_BG}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${TITLE_X}" y="${TITLE_Y}"/><a:ext cx="${TITLE_W}" cy="${TITLE_H}"/></a:xfrm></p:spPr>${makeTitleBody(title)}</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr bwMode="auto"><a:xfrm><a:off x="${BODY_X}" y="${BODY_Y}"/><a:ext cx="${BODY_W}" cy="${BODY_H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${bodyParas}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

// Build a scripture slide XML
function buildScriptureSlideXml(title, lines) {
  // All verses flow as one paragraph; verse numbers are inline superscripts
  const runs = lines.map((line, i) => {
    const verseMatch = line.match(/^(\d+)\s+([\s\S]*)$/);
    if (verseMatch) {
      const [, num, body] = verseMatch;
      const numRun = `<a:r><a:rPr lang="en-US" sz="${NORM_BODY_SZ}" b="0" baseline="15000" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>${NORM_LATIN}</a:rPr><a:t>${xmlEscape(num)}</a:t></a:r>`;
      const space = `<a:r><a:rPr lang="en-US" sz="${NORM_BODY_SZ}" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>${NORM_LATIN}</a:rPr><a:t> </a:t></a:r>`;
      const textRun = makeRun((i < lines.length - 1 ? body + ' ' : body), { sz: NORM_BODY_SZ, bold: true });
      return numRun + space + textRun;
    }
    return makeRun((i < lines.length - 1 ? line + ' ' : line), { sz: NORM_BODY_SZ, bold: true });
  }).join('');
  const bodyParas = `<a:p><a:pPr/>${runs}</a:p>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${SLIDE_BG}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${TITLE_X}" y="${TITLE_Y}"/><a:ext cx="${TITLE_W}" cy="${TITLE_H}"/></a:xfrm></p:spPr>${makeTitleBody(title)}</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr bwMode="auto"><a:xfrm><a:off x="${BODY_X}" y="${BODY_Y}"/><a:ext cx="${BODY_W}" cy="${BODY_H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${bodyParas}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

// Build a missing song placeholder slide
function buildMissingSlideXml(songName) {
  const titleText = `⚠ MISSING: ${songName}`;
  const bodyPara = `<a:p><a:pPr/><a:r><a:rPr lang="en-US" sz="3200" dirty="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial Narrow" pitchFamily="34" charset="0"/></a:rPr><a:t>Song file not found in library. Please add it manually.</a:t></a:r></a:p>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${SLIDE_BG}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${TITLE_X}" y="${TITLE_Y}"/><a:ext cx="${TITLE_W}" cy="${TITLE_H}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr lIns="48768" tIns="48768" rIns="48768" bIns="48768" anchor="b"/><a:lstStyle/><a:p><a:pPr defTabSz="1169988"/><a:r><a:rPr lang="en-US" sz="3400" dirty="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial Narrow" pitchFamily="34" charset="0"/></a:rPr><a:t>${xmlEscape(titleText)}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr bwMode="auto"><a:xfrm><a:off x="${BODY_X}" y="${BODY_Y}"/><a:ext cx="${BODY_W}" cy="${BODY_H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${bodyPara}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

// Build the blank separator slide XML
function buildBlankSlideXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${SLIDE_BG}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

// Slide relationship pointing to layout
function makeSlideRels(layoutNum = 1) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${layoutNum}.xml"/></Relationships>`;
}

/**
 * Build the output PPTX from a template and list of slide descriptors.
 *
 * @param {string} templatePath - Path to template.pptx
 * @param {Array} slideDescriptors - Array of slide objects to add
 * @param {function} onProgress - Progress callback
 * @returns {Buffer} - The output PPTX as a Buffer
 */
async function buildPptx(referencePath, slideDescriptors, onProgress = () => {}) {
  const templateZip = new AdmZip(referencePath);
  const outZip = new AdmZip();

  // Copy reference entries, excluding content slides 2+, their notes, and the
  // files we rebuild from scratch (presentation.xml, rels, Content_Types).
  for (const entry of templateZip.getEntries()) {
    if (entry.isDirectory) continue;
    const n = entry.entryName;
    if (/^ppt\/slides\/slide\d+\.xml$/.test(n) && n !== 'ppt/slides/slide1.xml') continue;
    if (/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(n) && n !== 'ppt/slides/_rels/slide1.xml.rels') continue;
    if (/^ppt\/notesSlides\//.test(n)) continue;
    if (n === 'ppt/presentation.xml') continue;   // added below after cleaning
    if (n === 'ppt/_rels/presentation.xml.rels') continue; // added below after cleaning
    if (n === '[Content_Types].xml') continue;    // rebuilt from scratch at end
    outZip.addFile(n, entry.getData());
  }

  // Read presentation.xml and rels from reference, then clean them down to slide1 only.
  const presEntry = templateZip.getEntry('ppt/presentation.xml');
  let presXml = presEntry.getData().toString('utf8');

  const presRelsEntry = templateZip.getEntry('ppt/_rels/presentation.xml.rels');
  let presRelsXml = presRelsEntry
    ? presRelsEntry.getData().toString('utf8')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

  // Strip all slide 2+ relationships from rels (keep slide1, masters, themes, etc.)
  presRelsXml = presRelsXml.replace(
    /<Relationship\s[^>]*Target="slides\/slide\d+\.xml"[^>]*\/>/g,
    match => /Target="slides\/slide1\.xml"/.test(match) ? match : ''
  );

  // Find the rId assigned to slide1 in the (now-cleaned) rels
  const rIdMatch = presRelsXml.match(/Id="(rId\d+)"[^>]*Target="slides\/slide1\.xml"/);
  const slide1RId = rIdMatch ? rIdMatch[1] : 'rId1';

  // Clean sldIdLst to only reference slide1
  presXml = presXml.replace(
    /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst><p:sldId id="256" r:id="${slide1RId}"/></p:sldIdLst>`
  );

  // (presentation.xml and rels are written once at the end after all slides are added)

  // nextSlideNum starts at 2 (slide1 is the title/logo from the reference)
  let nextSlideNum = 2;

  // nextRId: find the highest rId already used in the cleaned rels, then go higher
  // Note: capture only the numeric part (Id="rId(\d+)") to avoid parseInt("rId3") = NaN
  const existingRIdNums = [...presRelsXml.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1]));
  let nextRId = existingRIdNums.length > 0 ? Math.max(...existingRIdNums) + 1 : 100;
  let nextMediaNum = 1;

  // Track used media names to avoid conflicts
  const existingMedia = new Set(
    templateZip.getEntries()
      .filter(e => e.entryName.startsWith('ppt/media/'))
      .map(e => path.basename(e.entryName))
  );

  // New slides/rels/media to add to presentation.xml
  const newSlideRefs = []; // { slideNum, rId }
  const newPresRels = [];  // relationship XML strings
  const newContentTypes = []; // content type override strings

  function addSlideToOutput(slideXml, relsXml, mediaFiles = {}, normalizePositions = false, preNormalized = false) {
    const slideNum = nextSlideNum++;
    const rId = `rId${nextRId++}`;
    const slidePath = `ppt/slides/slide${slideNum}.xml`;
    const slideRelsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;

    // Handle media file name conflicts
    const mediaRemap = {};
    for (const [origName, data] of Object.entries(mediaFiles)) {
      let newName = origName;
      let counter = 1;
      while (existingMedia.has(newName)) {
        const ext = path.extname(origName);
        const base = path.basename(origName, ext);
        newName = `${base}_${counter++}${ext}`;
      }
      existingMedia.add(newName);
      mediaRemap[origName] = newName;
      outZip.addFile(`ppt/media/${newName}`, data);
      // Add content type for media
      const ext = path.extname(newName).toLowerCase();
      const mimeTypes = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.wmf': 'image/x-wmf', '.emf': 'image/x-emf' };
      const mime = mimeTypes[ext] || 'application/octet-stream';
      newContentTypes.push(`<Override PartName="/ppt/media/${newName}" ContentType="${mime}"/>`);
    }

    // Update rels XML with remapped media paths
    let finalRelsXml = relsXml || makeSlideRels(1);
    for (const [orig, newName] of Object.entries(mediaRemap)) {
      finalRelsXml = finalRelsXml.replace(new RegExp(orig.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), newName);
    }

    // Strip notes slide relationships (notes weren't copied into the output)
    finalRelsXml = finalRelsXml.replace(/<Relationship\s[^>]*Target="[^"]*notesSlides[^"]*"[^>]*\/>\s*/g, '');

    // If a slideLayout referenced by this slide doesn't exist in the output, fall back to slideLayout1
    finalRelsXml = finalRelsXml.replace(
      /(<Relationship\s[^>]*Target="\.\.\/)slideLayouts\/(slideLayout\d+\.xml)("[^>]*\/>)/g,
      (match, pre, layout, post) => {
        return outZip.getEntry('ppt/slideLayouts/' + layout) ? match : pre + 'slideLayouts/slideLayout1.xml' + post;
      }
    );

    const finalXml = preNormalized ? slideXml : normalizeSlide(slideXml, normalizePositions);
    outZip.addFile(slidePath, Buffer.from(finalXml, 'utf8'));
    outZip.addFile(slideRelsPath, Buffer.from(finalRelsXml, 'utf8'));

    newSlideRefs.push({ slideNum, rId });
    newPresRels.push(`<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`);
    newContentTypes.push(`<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);

    nextMediaNum++;
    return slideNum;
  }

  // Add blank separator
  function addBlank() {
    addSlideToOutput(buildBlankSlideXml(), makeSlideRels(1));
  }

  // Process each slide descriptor.
  // Blank slides are inserted at TYPE TRANSITIONS (reading→song, song→scripture, etc.)
  // but NOT between consecutive items of the same type.
  // Always add a blank before the first content item (right after the title slide).
  const warnings = [];
  let prevType = 'title'; // virtual previous type to force blank after title
  let prevDesc = null;

  const isCatechismOrCreed = title => /^(Heidelberg Catechism|Nicene Creed|Apostles' Creed|Lord's Prayer)/i.test(title || '');

  // Helper: add content slides for one descriptor
  function addDescriptorSlides(desc) {
    if (desc.type === 'reading') {
      for (const slideLines of desc.slideGroups) {
        const slideXml = buildReadingSlideXml(desc.title, slideLines);
        addSlideToOutput(slideXml, makeSlideRels(1), {}, true);
      }
      return;
    }

    if (desc.type === 'scripture') {
      for (const slide of (desc.slides || [])) {
        const slideXml = buildScriptureSlideXml(slide.title, slide.lines);
        addSlideToOutput(slideXml, makeSlideRels(1), {}, true);
      }
      return;
    }

    if (desc.type === 'song') {
      if (desc.pptxPath) {
        try {
          const songExtractor = require('./ppt-extractor');
          const { slides: songSlides } = songExtractor.extractSlides(desc.pptxPath);
          for (const s of songSlides) {
            addSlideToOutput(normalizeSongSlide(s.slideXml), s.relsXml, s.mediaFiles, false, true);
          }
          return;
        } catch (e) {
          warnings.push({ type: 'song-error', title: desc.title, error: e.message });
        }
      } else {
        warnings.push({ type: 'missing-song', title: desc.title });
      }
      // Fallback: missing/error placeholder
      const slideXml = buildMissingSlideXml(desc.title);
      addSlideToOutput(slideXml, makeSlideRels(1));
    }
  }

  for (let itemIndex = 0; itemIndex < slideDescriptors.length; itemIndex++) {
    const desc = slideDescriptors[itemIndex];
    onProgress(itemIndex + 1, slideDescriptors.length, desc);

    if (desc.type === 'blank') {
      addBlank();
      prevType = 'blank';
      continue;
    }

    // Add blank if type changed (including first item after the title),
    // or between consecutive readings where one is a catechism/creed item.
    const currentType = desc.type;
    if (currentType !== prevType) {
      addBlank();
    } else if (currentType === 'reading' &&
               (isCatechismOrCreed(desc.title) || isCatechismOrCreed(prevDesc && prevDesc.title))) {
      addBlank();
    }

    addDescriptorSlides(desc);
    prevType = currentType;
    prevDesc = desc;
  }

  // Always end with a blank slide
  if (prevType !== 'blank') {
    addBlank();
  }

  // Update presentation.xml: add sldIdLst entries
  // Find the sldIdLst and add new entries
  let updatedPresXml = presXml;

  // slide1 was given id="256" above; start new slides at 257
  let maxSldId = 256;

  const newSldEntries = newSlideRefs.map(({ slideNum, rId }) => {
    maxSldId++;
    return `<p:sldId id="${maxSldId}" r:id="${rId}"/>`;
  }).join('');

  if (updatedPresXml.includes('</p:sldIdLst>')) {
    updatedPresXml = updatedPresXml.replace('</p:sldIdLst>', newSldEntries + '</p:sldIdLst>');
  } else if (updatedPresXml.includes('<p:sldIdLst/>')) {
    updatedPresXml = updatedPresXml.replace('<p:sldIdLst/>', `<p:sldIdLst>${newSldEntries}</p:sldIdLst>`);
  }

  outZip.addFile('ppt/presentation.xml', Buffer.from(updatedPresXml, 'utf8'));

  // Update presentation rels
  let updatedPresRels = presRelsXml;
  const relsToAdd = newPresRels.join('\n  ');
  updatedPresRels = updatedPresRels.replace('</Relationships>', `  ${relsToAdd}\n</Relationships>`);
  outZip.addFile('ppt/_rels/presentation.xml.rels', Buffer.from(updatedPresRels, 'utf8'));

  // Rebuild [Content_Types].xml from scratch based on what's actually in the output ZIP.
  const CT_OVERRIDES = [
    [/\/ppt\/presentation\.xml$/,                 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
    [/\/ppt\/slides\/slide\d+\.xml$/,             'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'],
    [/\/ppt\/slideLayouts\/slideLayout\d+\.xml$/, 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],
    [/\/ppt\/slideMasters\/slideMaster\d+\.xml$/, 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
    [/\/ppt\/theme\/theme\d+\.xml$/,              'application/vnd.openxmlformats-officedocument.theme+xml'],
    [/\/ppt\/notesMasters\/notesMaster\d+\.xml$/, 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'],
    [/\/ppt\/presProps\.xml$/,                    'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'],
    [/\/ppt\/viewProps\.xml$/,                    'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'],
    [/\/ppt\/tableStyles\.xml$/,                  'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'],
    [/\/docProps\/core\.xml$/,                    'application/vnd.openxmlformats-package.core-properties+xml'],
    [/\/docProps\/app\.xml$/,                     'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
  ];
  const CT_EXT_DEFAULTS = {
    rels: 'application/vnd.openxmlformats-package.relationships+xml',
    xml:  'application/xml',
    png:  'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    gif:  'image/gif', wmf: 'image/x-wmf', emf: 'image/x-emf',
  };
  const usedExts = new Set();
  const ctOverrides = [];
  for (const e of outZip.getEntries()) {
    if (e.isDirectory) continue;
    const partPath = '/' + e.entryName;
    const ext = e.entryName.split('.').pop().toLowerCase();
    const hit = CT_OVERRIDES.find(([re]) => re.test(partPath));
    if (hit) ctOverrides.push(`<Override PartName="${partPath}" ContentType="${hit[1]}"/>`);
    else if (CT_EXT_DEFAULTS[ext]) usedExts.add(ext);
  }
  const ctDefaults = [...usedExts].map(ext =>
    `<Default Extension="${ext}" ContentType="${CT_EXT_DEFAULTS[ext]}"/>`
  );
  const freshCt = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    ctDefaults.join('') + ctOverrides.join('') + `</Types>`;
  outZip.addFile('[Content_Types].xml', Buffer.from(freshCt, 'utf8'));

  return { buffer: outZip.toBuffer(), warnings };
}

module.exports = {
  buildPptx,
  splitReadingLines,
  buildReadingSlideXml,
  buildScriptureSlideXml,
  buildBlankSlideXml,
  buildMissingSlideXml
};
