// Run from the repository root: node --test office_plugin/macos/tests/*.test.mjs
import assert from 'node:assert/strict';
import test from 'node:test';
import {createMetadata, metadataFromTags, metadataToTags, MetadataError, TAG} from '../web/app/metadata.mjs';

const typography = {
  typographyVersion: 1, symbolFontId: 'mathjax-tex', numberFontFamily: '', cjkFontFamily: 'PingFang SC',
  defaultMathStyle: 'Automatic', fontSizePoints: 12, color: '#000000',
};

function formula(latex = '\\frac{a}{b}') {
  return createMetadata({documentId: 'doc', equationId: 'eq', latex, display: true, renderEngine: 'Image',
    typography, naturalWidthPoints: 20.5, naturalHeightPoints: 30});
}

test('LaTeX is stored as upper-case UTF-8 hex in 200-character chunks, like PowerPointFormulaMetadataStore', () => {
  const latex = 'x'.repeat(150) + '中';
  const tags = new Map(metadataToTags(formula(latex)));
  const bytes = new TextEncoder().encode(latex);
  assert.equal(tags.get('LaTeXSnipperLatexBytes'), String(bytes.length));
  assert.equal(tags.get('LaTeXSnipperLatexChunks'), '2');
  const hex = [...bytes].map(b => b.toString(16).toUpperCase().padStart(2, '0')).join('');
  assert.equal(tags.get('LaTeXSnipperLatex0000'), hex.slice(0, 200));
  assert.equal(tags.get('LaTeXSnipperLatex0001'), hex.slice(200));
  assert.equal(tags.get(TAG.schemaVersion), '3');
  assert.equal(tags.get(TAG.renderEngine), 'Image');
  assert.equal(tags.get(TAG.displayMode), 'Display');
});

test('round trips through PowerPoint, which upper-cases tag names', () => {
  const original = formula('\\int_0^1 \\text{面积}\\,dx');
  const shouted = metadataToTags(original).map(([key, value]) => [key.toUpperCase(), value]);
  const restored = metadataFromTags(shouted);
  assert.equal(restored.latex, original.latex);
  assert.deepEqual(restored.typography, typography);
  assert.equal(restored.naturalWidthPoints, 20.5);
  assert.equal(restored.imageFormat, 'Svg');
});

test('reads tags written by the Windows plugin, which has no image-format tag', () => {
  const windows = metadataToTags(formula()).filter(([key]) => key !== TAG.imageFormat);
  assert.equal(metadataFromTags(windows).imageFormat, 'Png');
});

test('an empty source is stored as a single "0" chunk', () => {
  const tags = new Map(metadataToTags({...formula(), latex: ''}));
  assert.equal(tags.get('LaTeXSnipperLatexBytes'), '0');
  assert.equal(tags.get('LaTeXSnipperLatex0000'), '0');
});

test('rejects truncated, corrupt or unsupported tag sets', () => {
  const tags = metadataToTags(formula());
  const without = name => tags.filter(([key]) => key !== name);
  const replace = (name, value) => tags.map(([key, old]) => [key, key === name ? value : old]);
  for (const broken of [without('LaTeXSnipperLatex0000'), replace('LaTeXSnipperLatexBytes', '999'),
    replace('LaTeXSnipperLatex0000', 'ZZ'), replace(TAG.schemaVersion, '2'), replace(TAG.renderEngine, 'Ole2')])
    assert.throws(() => metadataFromTags(broken), MetadataError);
});

test('validates typography like FormulaTypographyFields.Read', () => {
  for (const change of [{typographyVersion: 2}, {color: 'red'}, {fontSizePoints: 0}, {defaultMathStyle: 'Nope'}])
    assert.throws(() => createMetadata({documentId: 'd', latex: 'x', display: false, renderEngine: 'Omml',
      typography: {...typography, ...change}}), MetadataError);
});
