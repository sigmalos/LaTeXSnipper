// Schema 3 formula metadata, mirroring FormulaMetadata / FormulaTypographyFields
// and the PowerPoint shape-tag layout of PowerPointFormulaMetadataStore.
export const SCHEMA_VERSION = 3;
export const TYPOGRAPHY_VERSION = 1;
export const RENDER_ENGINES = ['Omml', 'Image', 'MathJaxSvg'];
export const DISPLAY_MODES = ['Inline', 'Display'];
const MATH_STYLES = ['Automatic', 'Upright', 'Bold', 'Italic', 'BoldItalic', 'SansSerif', 'SansSerifBold',
  'SansSerifItalic', 'SansSerifBoldItalic', 'Monospace', 'Calligraphic', 'BoldCalligraphic', 'Script',
  'BoldScript', 'Fraktur', 'BoldFraktur', 'Blackboard'];

export class MetadataError extends Error {}

export function newId() {
  return crypto.randomUUID().replaceAll('-', '');
}

export function validateTypography(fields) {
  const fail = message => { throw new MetadataError(message); };
  if (!fields || typeof fields !== 'object') fail('缺少公式样式。');
  for (const name of ['typographyVersion', 'symbolFontId', 'numberFontFamily', 'cjkFontFamily',
    'defaultMathStyle', 'fontSizePoints', 'color'])
    if (fields[name] === undefined || fields[name] === null) fail('缺少公式样式字段：' + name);
  if (String(fields.typographyVersion) !== String(TYPOGRAPHY_VERSION)) fail('不支持的公式样式版本。');
  if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(fields.symbolFontId)) fail('数学字体必须使用稳定的字体方案 ID。');
  if (!MATH_STYLES.includes(fields.defaultMathStyle)) fail('无效的数学样式。');
  const points = Number(fields.fontSizePoints);
  if (!(points >= 1 && points <= 1638)) fail('公式字号必须在 1–1638 pt 之间。');
  if (!/^#[0-9a-fA-F]{6}$/.test(fields.color)) fail('颜色必须为 #RRGGBB 格式。');
  if (!String(fields.cjkFontFamily).trim()) fail('字体家族不能为空。');
  return {
    typographyVersion: TYPOGRAPHY_VERSION,
    symbolFontId: fields.symbolFontId,
    numberFontFamily: String(fields.numberFontFamily || '').trim(),
    cjkFontFamily: String(fields.cjkFontFamily).trim(),
    defaultMathStyle: fields.defaultMathStyle,
    fontSizePoints: points,
    color: fields.color.toUpperCase(),
  };
}

export function createMetadata({documentId, latex, display, renderEngine, typography, naturalWidthPoints,
  naturalHeightPoints, equationId = newId()}) {
  return validateMetadata({
    schemaVersion: SCHEMA_VERSION,
    documentId,
    equationId,
    latex,
    displayMode: display ? 'Display' : 'Inline',
    renderEngine,
    typography,
    naturalWidthPoints,
    naturalHeightPoints,
  });
}

export function validateMetadata(value) {
  const fail = () => { throw new MetadataError('所选公式的元数据无效或已损坏。'); };
  if (!value || typeof value !== 'object' || value.schemaVersion !== SCHEMA_VERSION) fail();
  if (!value.documentId || !value.equationId || typeof value.latex !== 'string') fail();
  if (!DISPLAY_MODES.includes(value.displayMode) || !RENDER_ENGINES.includes(value.renderEngine)) fail();
  for (const name of ['naturalWidthPoints', 'naturalHeightPoints'])
    if (value[name] !== undefined && !(Number(value[name]) > 0)) fail();
  return {...value, typography: validateTypography(value.typography)};
}

// PowerPoint shape tags -------------------------------------------------------

export const TAG = {
  equationId: 'LaTeXSnipperEquationId',
  documentId: 'LaTeXSnipperDocumentId',
  displayMode: 'LaTeXSnipperDisplayMode',
  schemaVersion: 'LaTeXSnipperSchemaVersion',
  renderEngine: 'LaTeXSnipperRenderEngine',
  naturalWidth: 'LaTeXSnipperNaturalWidthPoints',
  naturalHeight: 'LaTeXSnipperNaturalHeightPoints',
  imagePath: 'LaTeXSnipperImagePath',
  // macOS only: whether an Image formula is an SVG or PNG picture. Ignored on Windows.
  imageFormat: 'LaTeXSnipperImageFormat',
  latexPrefix: 'LaTeXSnipperLatex',
  typographyPrefix: 'LaTeXSnipperTypography',
};
const CHUNK = 200;
const MAX_CHUNKS = 10000;

function encodeChunks(prefix, text) {
  const bytes = new TextEncoder().encode(text || '');
  let hex = '';
  for (const byte of bytes) hex += byte.toString(16).toUpperCase().padStart(2, '0');
  const payload = hex || '0';
  const count = Math.ceil(payload.length / CHUNK);
  const tags = [[prefix + 'Bytes', String(bytes.length)], [prefix + 'Chunks', String(count)]];
  for (let index = 0; index < count; index++)
    tags.push([prefix + String(index).padStart(4, '0'), payload.slice(index * CHUNK, (index + 1) * CHUNK)]);
  return tags;
}

function decodeChunks(prefix, read) {
  const length = Number.parseInt(read(prefix + 'Bytes'), 10);
  const count = Number.parseInt(read(prefix + 'Chunks'), 10);
  if (!(length >= 0) || !(count > 0) || count > MAX_CHUNKS) throw new MetadataError('公式数据不完整。');
  let hex = '';
  for (let index = 0; index < count; index++) hex += read(prefix + String(index).padStart(4, '0'));
  if (length === 0) {
    if (hex !== '0') throw new MetadataError('公式数据不完整。');
    return '';
  }
  if (hex.length !== length * 2 || !/^[0-9A-Fa-f]*$/.test(hex)) throw new MetadataError('公式数据不完整。');
  const bytes = new Uint8Array(length);
  for (let index = 0; index < length; index++) bytes[index] = Number.parseInt(hex.slice(index * 2, index * 2 + 2), 16);
  return new TextDecoder('utf-8', {fatal: true}).decode(bytes);
}

/** All tag key/value pairs that describe a formula picture. */
export function metadataToTags(metadata) {
  return [
    [TAG.documentId, metadata.documentId],
    [TAG.equationId, metadata.equationId],
    [TAG.schemaVersion, String(SCHEMA_VERSION)],
    [TAG.renderEngine, metadata.renderEngine],
    ...encodeChunks(TAG.latexPrefix, metadata.latex),
    [TAG.displayMode, metadata.displayMode],
    ...encodeChunks(TAG.typographyPrefix, JSON.stringify(metadata.typography)),
    [TAG.naturalWidth, String(metadata.naturalWidthPoints)],
    [TAG.naturalHeight, String(metadata.naturalHeightPoints)],
    [TAG.imagePath, ''],
    [TAG.imageFormat, metadata.imageFormat === 'Png' ? 'Png' : 'Svg'],
  ];
}

/** Reads tags case-insensitively: PowerPoint normalizes tag names to upper case. */
export function metadataFromTags(entries) {
  const values = new Map(entries.map(([key, value]) => [String(key).toUpperCase(), String(value ?? '')]));
  const read = name => {
    const value = values.get(name.toUpperCase());
    if (value === undefined) throw new MetadataError('所选公式的元数据不完整。');
    return value;
  };
  let typography;
  try { typography = JSON.parse(decodeChunks(TAG.typographyPrefix, read)); }
  catch (error) { throw error instanceof MetadataError ? error : new MetadataError('公式样式数据无效。'); }
  const imageFormat = values.get(TAG.imageFormat.toUpperCase()) === 'Svg' ? 'Svg' : 'Png';
  return validateMetadata({
    imageFormat,
    schemaVersion: Number.parseInt(read(TAG.schemaVersion), 10),
    documentId: read(TAG.documentId),
    equationId: read(TAG.equationId),
    latex: decodeChunks(TAG.latexPrefix, read),
    displayMode: read(TAG.displayMode),
    renderEngine: read(TAG.renderEngine),
    typography,
    naturalWidthPoints: Number(read(TAG.naturalWidth)),
    naturalHeightPoints: Number(read(TAG.naturalHeight)),
  });
}

export function isFormulaTagSet(entries) {
  return entries.some(([key]) => String(key).toUpperCase() === TAG.equationId.toUpperCase());
}
