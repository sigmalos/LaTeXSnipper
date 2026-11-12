// Word host: native OMML or SVG picture formulas.
//
// Word for Mac lays an equation out as a display equation only when it sits
// directly in its own paragraph; any content control around it forces inline
// layout. So formulas use two anchors:
//   - inline formulas: an inline content control tagged `latexsnipper-eq-<id>`;
//   - display formulas: a centred paragraph holding a hidden bookmark `_LS<id>`.
// Full metadata lives in the add-in's document settings; picture formulas also
// carry it in their alternative text so a copy in another document can still
// be loaded.
/* global Office, Word */
import {MetadataError, newId, validateMetadata} from './metadata.mjs';

export const TAG_PREFIX = 'latexsnipper-eq-';
export const BOOKMARK_PREFIX = '_LS';
const SETTINGS_PREFIX = 'LaTeXSnipper.Formula.';
const DOCUMENT_ID_KEY = 'LaTeXSnipper.DocumentId';
const PICTURE_TITLE_PREFIX = 'LaTeXSnipper ';
const EMU_PER_POINT = 12700;

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
};

export class WordFormulaError extends Error {}

const escapeXml = value => String(value).replace(/[<>&"']/g, character =>
  ({'<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;'})[character]);

const tagOf = equationId => TAG_PREFIX + equationId;
const bookmarkOf = equationId => BOOKMARK_PREFIX + equationId;
const idFromTag = tag => typeof tag === 'string' && tag.startsWith(TAG_PREFIX) ? tag.slice(TAG_PREFIX.length) : null;
const idFromBookmark = name => typeof name === 'string' && /^_LS[0-9a-f]{32}$/.test(name) ? name.slice(BOOKMARK_PREFIX.length) : null;

// Document settings ----------------------------------------------------------------

function saveSettings() {
  return new Promise((resolve, reject) => Office.context.document.settings.saveAsync(result =>
    result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error('无法保存公式元数据。'))));
}

export async function documentId() {
  const settings = Office.context.document.settings;
  let id = settings.get(DOCUMENT_ID_KEY);
  if (typeof id !== 'string' || !id) {
    id = newId();
    settings.set(DOCUMENT_ID_KEY, id);
    await saveSettings();
  }
  return id;
}

function storeMetadata(metadata) {
  Office.context.document.settings.set(SETTINGS_PREFIX + metadata.equationId, JSON.stringify(metadata));
}

function forgetMetadata(equationId) {
  Office.context.document.settings.remove(SETTINGS_PREFIX + equationId);
}

function storedMetadata(equationId) {
  const raw = Office.context.document.settings.get(SETTINGS_PREFIX + equationId);
  if (typeof raw !== 'string') return null;
  try { return validateMetadata(JSON.parse(raw)); } catch { return null; }
}

// OOXML ----------------------------------------------------------------------------

function flatPackage(bodyXml, media = []) {
  const relationships = media.map(item =>
    `<Relationship Id="${item.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${item.name}"/>`).join('');
  const parts = media.map(item =>
    `<pkg:part pkg:name="/word/media/${item.name}" pkg:contentType="${item.type}" pkg:compression="store"><pkg:binaryData>${item.base64}</pkg:binaryData></pkg:part>`).join('');
  return '<?xml version="1.0" standalone="yes"?>'
    + '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships></pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>'
    + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`
    + '</pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    + `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:m="${NS.m}">`
    + `<w:body>${bodyXml}</w:body></w:document></pkg:xmlData></pkg:part>${parts}</pkg:package>`;
}

/**
 * Body XML for one formula. `insert` ends a display formula's package with an
 * empty paragraph: Word merges the package's last paragraph into the
 * destination and keeps the destination's properties, so without it the
 * formula paragraph would lose its centring. Replacing a display formula in
 * place merges into its own centred paragraph and needs no extra paragraph.
 */
function placed(metadata, content, {insert}) {
  const id = metadata.equationId;
  if (metadata.displayMode !== 'Display') {
    return '<w:p><w:sdt><w:sdtPr><w:alias w:val="LaTeXSnipper"/>'
      + `<w:tag w:val="${escapeXml(tagOf(id))}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
  }
  return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:bookmarkStart w:id="0" w:name="${bookmarkOf(id)}"/>`
    + `${content}<w:bookmarkEnd w:id="0"/></w:p>${insert ? '<w:p/>' : ''}`;
}

/** OMML from the converter, always as an `m:oMath` (see the file comment). */
export function ommlContent(omml) {
  return String(omml).replace(/^<\?xml[^>]*\?>\s*/, '')
    .replace(/^<m:oMathPara([^>]*)>\s*<m:oMath>/, '<m:oMath$1>').replace(/<\/m:oMathPara>\s*$/, '');
}

export function ommlPackage(metadata, omml, options) {
  return flatPackage(placed(metadata, ommlContent(omml), options));
}

function pictureTitle(metadata) {
  const {latex: _latex, ...rest} = metadata;
  return PICTURE_TITLE_PREFIX + JSON.stringify(rest);
}

export function picturePackage(metadata, rendered, pngBase64, svgBase64, options) {
  const cx = Math.max(1, Math.round(rendered.widthPoints * EMU_PER_POINT));
  const cy = Math.max(1, Math.round(rendered.heightPoints * EMU_PER_POINT));
  // w:position is in half-points; negative lowers inline pictures to the text baseline.
  const lower = metadata.displayMode === 'Display' ? 0 : Math.round(rendered.baselinePoints * 2);
  const drawing = `<w:r><w:rPr><w:position w:val="${-lower}"/></w:rPr><w:drawing>`
    + `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>`
    + `<wp:docPr id="1" name="LaTeXSnipper formula" descr="${escapeXml(metadata.latex)}" title="${escapeXml(pictureTitle(metadata))}"/>`
    + '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    + `<a:graphic><a:graphicData uri="${NS.pic}"><pic:pic>`
    + '<pic:nvPicPr><pic:cNvPr id="0" name="latexsnipper-formula.svg"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill><a:blip r:embed="rIdLaTeXSnipperPng"><a:extLst>'
    + '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
    + '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdLaTeXSnipperSvg"/>'
    + '</a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  return flatPackage(placed(metadata, drawing, options), [
    {id: 'rIdLaTeXSnipperPng', name: 'latexsnipper-formula.png', type: 'image/png', base64: pngBase64},
    {id: 'rIdLaTeXSnipperSvg', name: 'latexsnipper-formula.svg', type: 'image/svg+xml', base64: svgBase64},
  ]);
}

// Anchors --------------------------------------------------------------------------

/** One managed formula location: {kind: 'control'|'bookmark', equationId, range, control?}. */
async function anchorsIn(context, range, includeAdjacent) {
  const controls = range.contentControls;
  controls.load('items/id,items/tag');
  const names = range.getBookmarks(true, includeAdjacent);
  await context.sync();
  const anchors = [];
  for (const control of controls.items) {
    const equationId = idFromTag(control.tag);
    if (equationId) anchors.push({kind: 'control', equationId, control, range: control.getRange('Whole')});
  }
  for (const name of names.value) {
    const equationId = idFromBookmark(name);
    if (equationId) anchors.push({kind: 'bookmark', equationId, name, range: context.document.getBookmarkRange(name)});
  }
  return anchors;
}

/** Managed formulas touching the selection; a caret inside a formula counts. */
async function selectedAnchors(context) {
  const selection = context.document.getSelection();
  const parent = selection.parentContentControlOrNullObject;
  parent.load('id,tag');
  const paragraphs = selection.paragraphs;
  paragraphs.load('items');
  await context.sync();
  const anchors = [];
  if (!parent.isNullObject && idFromTag(parent.tag))
    anchors.push({kind: 'control', equationId: idFromTag(parent.tag), control: parent, range: parent.getRange('Whole')});
  // Display formulas own their paragraph, so a caret anywhere in it selects
  // them. Adjacent lookup also reports the previous paragraph's bookmark, so
  // keep only bookmarks whose paragraph really overlaps the selected ones.
  const scope = paragraphs.items.length
    ? paragraphs.items[0].getRange('Whole').expandTo(paragraphs.items[paragraphs.items.length - 1].getRange('Whole'))
    : selection;
  anchors.push(...(await anchorsIn(context, selection, false)).filter(anchor => anchor.kind === 'control'));
  const candidates = (await anchorsIn(context, scope, true)).filter(anchor => anchor.kind === 'bookmark');
  const relations = candidates.map(anchor => anchor.range.paragraphs.getFirst().getRange('Whole').compareLocationWith(scope));
  await context.sync();
  const overlapping = new Set(['Equal', 'Inside', 'InsideStart', 'InsideEnd', 'Contains', 'OverlapsBefore', 'OverlapsAfter']);
  candidates.forEach((anchor, index) => { if (overlapping.has(relations[index].value)) anchors.push(anchor); });
  const seen = new Set();
  return anchors.filter(anchor => {
    const key = anchor.kind === 'control' ? `c${anchor.control.id}` : `b${anchor.name}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function pictureMetadata(context, range) {
  const picture = range.inlinePictures.getFirstOrNullObject();
  picture.load('altTextTitle,altTextDescription');
  await context.sync();
  if (picture.isNullObject || !String(picture.altTextTitle || '').startsWith(PICTURE_TITLE_PREFIX)) return null;
  try {
    const rest = JSON.parse(picture.altTextTitle.slice(PICTURE_TITLE_PREFIX.length));
    return validateMetadata({...rest, latex: picture.altTextDescription || ''});
  } catch { return null; }
}

async function readAnchor(context, anchor) {
  const metadata = storedMetadata(anchor.equationId) || await pictureMetadata(context, anchor.range);
  if (!metadata) throw new WordFormulaError('所选公式的元数据缺失。跨文档复制的 Word 原生公式无法恢复源码，请重新插入。');
  return {...metadata, equationId: anchor.equationId};
}

// Reading --------------------------------------------------------------------------

export async function selectedFormulas() {
  return Word.run(async context => {
    const formulas = [];
    for (const anchor of await selectedAnchors(context)) formulas.push(await readAnchor(context, anchor));
    return formulas;
  });
}

export async function selectedFormula() {
  const formulas = await selectedFormulas();
  if (formulas.length !== 1) throw new WordFormulaError(formulas.length ? '请只选择一个 LaTeXSnipper 公式。' : '请先选择一个 LaTeXSnipper 公式。');
  return formulas[0];
}

export async function allFormulas() {
  return Word.run(async context => {
    const formulas = [];
    for (const anchor of await anchorsIn(context, context.document.body.getRange('Whole'), false)) {
      try { formulas.push(await readAnchor(context, anchor)); }
      catch (error) { if (!(error instanceof WordFormulaError || error instanceof MetadataError)) throw error; }
    }
    return formulas;
  });
}

export async function selectionFontSize() {
  return Word.run(async context => {
    const font = context.document.getSelection().font;
    font.load('size');
    await context.sync();
    return typeof font.size === 'number' && font.size > 0 ? font.size : null;
  });
}

// Writing --------------------------------------------------------------------------

/**
 * Where a new formula goes (never inside another formula), plus the alignment
 * for the paragraph Word leaves after a display formula: it continues the
 * user's text, so it keeps their alignment rather than the formula's centring.
 */
async function insertionTarget(context, display) {
  const selection = context.document.getSelection();
  const paragraph = selection.paragraphs.getFirst();
  const pictures = paragraph.inlinePictures;
  const controls = paragraph.contentControls;
  paragraph.load('text');
  pictures.load('items');
  controls.load('items');
  await context.sync();
  const occupied = (await selectedAnchors(context)).length > 0;
  paragraph.load('alignment');
  await context.sync();
  const trailingAlignment = occupied ? 'Left' : paragraph.alignment;
  if (!display) {
    if (!occupied) return {target: selection, trailingAlignment};
    const next = paragraph.insertParagraph('', 'After');
    next.alignment = trailingAlignment;
    return {target: next.getRange('Start'), trailingAlignment};
  }
  // A display formula owns its paragraph, like a Word display equation.
  // Pictures and controls have no text, so check them as well.
  const empty = !paragraph.text.trim() && !pictures.items.length && !controls.items.length;
  return {target: occupied || !empty ? paragraph.insertParagraph('', 'After') : paragraph, trailingAlignment};
}

async function anchorFor(context, equationId) {
  const controls = context.document.contentControls.getByTag(tagOf(equationId));
  controls.load('items/id');
  const bookmark = context.document.getBookmarkRangeOrNullObject(bookmarkOf(equationId));
  await context.sync();
  if (!bookmark.isNullObject) return {kind: 'bookmark', range: bookmark, copies: 1};
  if (!controls.items.length) return null;
  // A copy inside the same document shares the tag: edit the selected copy.
  let control = controls.items[0];
  if (controls.items.length > 1) {
    const selected = (await selectedAnchors(context)).find(anchor => anchor.equationId === equationId && anchor.control);
    if (selected) control = selected.control;
  }
  return {kind: 'control', control, range: control.getRange('Whole'), copies: controls.items.length};
}

/**
 * Inserts at the selection, or replaces the formula `existing` (an equation id)
 * in place. `buildPackage(metadata, {insert})` returns the Flat OPC to insert;
 * the metadata it receives carries a fresh id when an in-document copy is edited.
 */
export async function writeFormula(metadata, buildPackage, existing = null) {
  await Word.run(async context => {
    let trailingAlignment = null;
    if (existing) {
      const anchor = await anchorFor(context, existing);
      if (!anchor) throw new WordFormulaError('原公式已不存在，请重新选择后再试。');
      if (anchor.copies > 1) metadata = {...metadata, equationId: newId()};
      if (anchor.kind === 'control') {
        anchor.control.cannotEdit = false;
        anchor.control.cannotDelete = false;
      }
      anchor.range.insertOoxml(buildPackage(metadata, {insert: false}), 'Replace');
      await context.sync();
      if (metadata.equationId !== existing && anchor.copies === 1) forgetMetadata(existing);
    } else {
      const placement = await insertionTarget(context, metadata.displayMode === 'Display');
      trailingAlignment = placement.trailingAlignment;
      placement.target.insertOoxml(buildPackage(metadata, {insert: true}), 'Replace');
      await context.sync();
    }
    const placedAnchor = await anchorFor(context, metadata.equationId);
    if (!placedAnchor) throw new WordFormulaError('Word 未能插入公式。');
    if (trailingAlignment && placedAnchor.kind === 'bookmark') {
      // Leave somewhere to keep typing after a display formula.
      const formulaParagraph = placedAnchor.range.paragraphs.getFirst();
      const next = formulaParagraph.getNextOrNullObject();
      await context.sync();
      if (next.isNullObject) formulaParagraph.insertParagraph('', 'After').alignment = trailingAlignment;
    }
    if (placedAnchor.kind === 'control') {
      const controls = context.document.contentControls.getByTag(tagOf(metadata.equationId));
      controls.load('items');
      await context.sync();
      for (const control of controls.items) {
        control.title = 'LaTeXSnipper';
        // Pictures are replaced as a whole; native equations stay editable in Word.
        control.cannotEdit = metadata.renderEngine !== 'Omml';
      }
      await context.sync();
    }
  });
  storeMetadata(metadata);
  await saveSettings();
  return metadata;
}

export async function deleteSelected() {
  const removed = await Word.run(async context => {
    const anchors = await selectedAnchors(context);
    for (const anchor of anchors) {
      if (anchor.kind === 'control') {
        anchor.control.cannotDelete = false;
        anchor.control.cannotEdit = false;
        anchor.control.delete(false);
      } else {
        // The display paragraph belongs to the formula; remove it with the bookmark.
        anchor.range.paragraphs.getFirst().delete();
        context.document.deleteBookmark(anchor.name);
      }
    }
    await context.sync();
    return anchors.map(anchor => anchor.equationId);
  });
  if (!removed.length) throw new WordFormulaError('请先选择 LaTeXSnipper 公式。');
  await Word.run(async context => {
    for (const id of new Set(removed)) {
      if (!await anchorFor(context, id)) forgetMetadata(id);
    }
  });
  await saveSettings();
  return removed.length;
}
