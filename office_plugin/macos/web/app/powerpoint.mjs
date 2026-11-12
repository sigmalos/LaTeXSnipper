// PowerPoint host: SVG or PNG pictures carrying the Windows shape-tag schema,
// so a formula inserted on either platform can be loaded on the other.
/* global Office, PowerPoint */
import {isFormulaTagSet, metadataFromTags, metadataToTags, newId} from './metadata.mjs';

const DOCUMENT_ID_KEY = 'LaTeXSnipper.DocumentId';

export class PowerPointFormulaError extends Error {}

function saveSettings() {
  return new Promise((resolve, reject) => Office.context.document.settings.saveAsync(result =>
    result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error('无法保存演示文稿标识。'))));
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

async function tagsOf(context, shape) {
  shape.tags.load('items/key,items/value');
  await context.sync();
  return shape.tags.items.map(tag => [tag.key, tag.value]);
}

async function currentSlide(context) {
  const slides = context.presentation.getSelectedSlides();
  slides.load('items/id');
  await context.sync();
  if (!slides.items.length) throw new PowerPointFormulaError('请先在普通视图中选择一张幻灯片。');
  return slides.items[0];
}

async function selectedFormulaShapes(context) {
  const slides = context.presentation.getSelectedSlides();
  slides.load('items/id');
  const shapes = context.presentation.getSelectedShapes();
  shapes.load('items/id,items/name,items/left,items/top,items/width,items/height');
  await context.sync();
  const result = [];
  for (const shape of shapes.items) {
    const tags = await tagsOf(context, shape);
    if (isFormulaTagSet(tags)) result.push({shape, tags, slideId: slides.items[0]?.id});
  }
  return result;
}

export async function selectedFormulas() {
  return PowerPoint.run(async context => (await selectedFormulaShapes(context))
    .map(({shape, tags, slideId}) => ({...metadataFromTags(tags), slideId, shapeId: shape.id})));
}

export async function selectedFormula() {
  const formulas = await selectedFormulas();
  if (formulas.length !== 1) throw new PowerPointFormulaError(formulas.length ? '请只选择一个 LaTeXSnipper 公式。' : '请先选择一个 LaTeXSnipper 公式。');
  return formulas[0];
}

export async function allFormulas() {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items/id');
    await context.sync();
    const formulas = [];
    for (const slide of slides.items) {
      slide.shapes.load('items/id');
      await context.sync();
      for (const shape of slide.shapes.items) {
        const tags = await tagsOf(context, shape);
        if (!isFormulaTagSet(tags)) continue;
        try { formulas.push({...metadataFromTags(tags), slideId: slide.id, shapeId: shape.id}); } catch { /* skipped */ }
      }
    }
    return formulas;
  });
}

export async function selectionFontSize() {
  return PowerPoint.run(async context => {
    const range = context.presentation.getSelectedTextRangeOrNullObject();
    range.load('font/size');
    await context.sync();
    if (range.isNullObject) return null;
    const size = range.font.size;
    return typeof size === 'number' && size > 0 ? size : null;
  });
}

/** Without a position PowerPoint centres the picture on the slide, like the Windows plugin. */
function setSelectedImage(data, coercionType, geometry) {
  const options = {coercionType, imageWidth: geometry.width, imageHeight: geometry.height};
  if (geometry.left !== undefined) Object.assign(options, {imageLeft: geometry.left, imageTop: geometry.top});
  return new Promise((resolve, reject) => Office.context.document.setSelectedDataAsync(data, options, result => result.status === Office.AsyncResultStatus.Succeeded
    ? resolve() : reject(new PowerPointFormulaError(result.error?.message || 'PowerPoint 未能插入公式图片。'))));
}

async function shapeIds(context, slide) {
  slide.shapes.load('items/id');
  await context.sync();
  return new Set(slide.shapes.items.map(shape => shape.id));
}

/**
 * Inserts a picture formula, or replaces the selected formula `target`. A
 * replacement keeps its position and the user's resize scale unless
 * `naturalSize` is set (the Format commands reset it).
 * `image` is {svg} or {base64}, plus natural {widthPoints, heightPoints}.
 */
export async function writeFormula(metadata, image, target = null, {naturalSize = false} = {}) {
  return PowerPoint.run(async context => {
    let geometry = {scale: 1};
    let previous = null;
    if (target) {
      const selected = (await selectedFormulaShapes(context))
        .filter(({tags}) => metadataFromTags(tags).equationId === target.equationId);
      if (selected.length !== 1) throw new PowerPointFormulaError('原公式已不在选区中，请重新选择后再试。');
      previous = selected[0].shape;
      const naturalWidth = Number(target.naturalWidthPoints);
      geometry = {left: previous.left, top: previous.top,
        scale: !naturalSize && naturalWidth > 0 ? previous.width / naturalWidth : 1};
    }
    const slide = await currentSlide(context);
    const before = await shapeIds(context, slide);
    const width = image.widthPoints * geometry.scale;
    const height = image.heightPoints * geometry.scale;
    const placement = {left: geometry.left, top: geometry.top, width, height};
    // Insert the replacement first: a failed insert must leave the original intact.
    if (image.svg) await setSelectedImage(image.svg, Office.CoercionType.XmlSvg, placement);
    else await setSelectedImage(image.base64, Office.CoercionType.Image, placement);
    slide.shapes.load('items/id');
    await context.sync();
    const inserted = slide.shapes.items.filter(shape => !before.has(shape.id));
    if (inserted.length !== 1) throw new PowerPointFormulaError('PowerPoint 未能定位新插入的公式图片。');
    const shape = inserted[0];
    // SVG graphics keep a locked aspect ratio, so each size assignment
    // rescales the other side; set both, then position, then verify.
    shape.width = width;
    shape.height = height;
    if (placement.left !== undefined) {
      shape.left = placement.left;
      shape.top = placement.top;
    }
    shape.name = 'LaTeXSnipper formula ' + metadata.equationId;
    for (const [key, value] of metadataToTags(metadata)) shape.tags.add(key, value);
    previous?.delete();
    await context.sync();
    return metadata;
  });
}

export async function deleteSelected() {
  return PowerPoint.run(async context => {
    const shapes = await selectedFormulaShapes(context);
    if (!shapes.length) throw new PowerPointFormulaError('请先选择 LaTeXSnipper 公式。');
    for (const {shape} of shapes) shape.delete();
    await context.sync();
    return shapes.length;
  });
}

/** Selects one formula shape so it can be rewritten in place by writeFormula. */
export async function selectShape(slideId, shapeId) {
  await PowerPoint.run(async context => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
    context.presentation.slides.getItem(slideId).setSelectedShapes([shapeId]);
    await context.sync();
  });
}
