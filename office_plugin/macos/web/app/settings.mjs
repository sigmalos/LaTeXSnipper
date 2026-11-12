// Settings dialog. Field handling follows the Windows settings.js; values are
// stored by the desktop app in the cross-platform settings.json schema.
/* global Office */
import {populateTypographyOptions} from '/shared/settings-typography.mjs';
import {api, catalog, loadSettings} from './api.mjs';
import {isChinese} from './strings.mjs';

const params = new URLSearchParams(location.search);
const host = params.get('host') === 'powerpoint' ? 'powerpoint' : 'word';
const locale = String(params.get('locale') || navigator.language || 'zh').toLowerCase();
const zh = isChinese(locale);
const hostName = host === 'word' ? 'Word' : 'PowerPoint';

const BACKENDS = {
  word: {
    values: ['Omml', 'Image'],
    zh: ['Word 原生公式', 'SVG 图片', '原生公式可在 Word 中直接编辑；SVG 图片与其他平台显示一致。'],
    en: ['Native equation', 'SVG picture', 'Native equations stay editable in Word; SVG pictures look the same on every platform.'],
  },
  powerpoint: {
    values: ['Image', 'Png'],
    zh: ['SVG 图片', 'PNG 图片', '默认插入矢量 SVG 图片，也可切换为 PNG 图片。'],
    en: ['SVG picture', 'PNG picture', 'Vector SVG pictures are the default. PNG pictures are also available.'],
  },
};

const TEXT = {
  zh: {
    title: `LaTeXSnipper ${hostName} 插件设置`,
    backendTitle: '公式插入方式',
    formulaDefaultsTitle: '公式默认属性',
    formulaDefaultsHint: '这些设置应用于新插入公式和格式化命令。',
    colorLabel: '字体颜色', resetColor: '恢复黑色', fontStyleLabel: '默认字形', fontSizeLabel: '公式字号',
    symbolFontLabel: '符号字体', numberFontLabel: '数字字体', cjkFontLabel: '汉字字体',
    presetTitle: '全局公式预设',
    presetHint: '导入会替换当前宿主的公式默认属性；导出时用 JSON 文件名命名预设。编辑器内调整只作用于当前公式。',
    importPreset: '导入 JSON', exportPreset: '导出 JSON', importedPreset: '已导入：', exportedPreset: '已导出：',
    followHostSize: '新建时跟随文字字号（无有效选区时使用上述字号）',
    editorTitle: '编辑器键盘行为', acceptShortcut: '插入或更新当前公式', fractionShortcut: '插入分式',
    rootShortcut: '插入根号', superscriptShortcut: '插入上标', subscriptShortcut: '插入下标', scriptsShortcut: '插入上下标',
    saved: '已保存。',
  },
  en: {
    title: `LaTeXSnipper ${hostName} Plugin Settings`,
    backendTitle: 'Formula Insertion',
    formulaDefaultsTitle: 'Default Formula Properties',
    formulaDefaultsHint: 'These settings apply to new formulas and formatting commands.',
    colorLabel: 'Font color', resetColor: 'Reset to black', fontStyleLabel: 'Default math style', fontSizeLabel: 'Formula size',
    symbolFontLabel: 'Symbol font', numberFontLabel: 'Number font', cjkFontLabel: 'CJK font',
    presetTitle: 'Global formula preset',
    presetHint: "Import replaces this host's formula defaults. The JSON filename names the preset. Editor changes apply to the current formula.",
    importPreset: 'Import JSON', exportPreset: 'Export JSON', importedPreset: 'Imported: ', exportedPreset: 'Exported: ',
    followHostSize: 'Follow text size for new formulas (use the size above when unavailable)',
    editorTitle: 'Editor Keyboard Behavior', acceptShortcut: 'insert or update the current formula',
    fractionShortcut: 'insert a fraction', rootShortcut: 'insert a square root', superscriptShortcut: 'insert a superscript',
    subscriptShortcut: 'insert a subscript', scriptsShortcut: 'insert superscript and subscript',
    saved: 'Saved.',
  },
};

const strings = zh ? TEXT.zh : TEXT.en;
const backend = BACKENDS[host];
const [firstLabel, secondLabel, hint] = zh ? backend.zh : backend.en;
Object.assign(strings, {backendFirst: firstLabel, backendSecond: secondLabel, backendHint: hint});

const element = id => document.getElementById(id);
const inputs = {
  formulaColor: element('formulaColor'),
  formulaMathStyle: element('formulaMathStyle'),
  formulaFontSizePoints: element('formulaFontSizePoints'),
  symbolFontId: element('symbolFontId'),
  numberFontFamily: element('numberFontFamily'),
  cjkFontFamily: element('cjkFontFamily'),
  followHostFontSize: element('followHostFontSize'),
};
const backendButtons = [...document.querySelectorAll('[data-backend]')];
const presetStatus = element('presetStatus');
let state = null;

function applyText() {
  document.title = strings.title;
  document.documentElement.lang = zh ? 'zh-CN' : 'en';
  for (const node of document.querySelectorAll('[data-i18n]')) node.textContent = strings[node.dataset.i18n] ?? node.textContent;
  backendButtons.forEach((button, index) => { button.dataset.backend = backend.values[index]; });
}

function render() {
  for (const button of backendButtons) button.classList.toggle('active', button.dataset.backend === state.insertionBackend);
  inputs.formulaColor.value = state.formulaColor.toLowerCase();
  inputs.formulaMathStyle.value = state.formulaMathStyle;
  inputs.formulaFontSizePoints.value = String(state.formulaFontSizePoints);
  inputs.symbolFontId.value = state.symbolFontId;
  inputs.numberFontFamily.value = state.numberFontFamily;
  inputs.cjkFontFamily.value = state.cjkFontFamily;
  inputs.followHostFontSize.checked = state.followHostFontSize;
}

function showStatus(message, error = false) {
  presetStatus.textContent = message;
  presetStatus.classList.toggle('error', error);
}

async function populate(settings) {
  const options = await catalog();
  state = settings;
  populateTypographyOptions({...options, namedSizes: Object.entries(options.namedSizes)
    .map(([name, points]) => ({name, points}))}, {
    symbolFontId: settings.symbolFontId,
    numberFontFamily: settings.numberFontFamily,
    cjkFontFamily: settings.cjkFontFamily,
    formulaFontSizePoints: settings.formulaFontSizePoints,
  }, locale);
  render();
}

async function save() {
  try {
    const saved = await api(`/api/settings/${host}`, {method: 'PUT', body: {
      insertionBackend: state.insertionBackend,
      formulaColor: state.formulaColor,
      formulaMathStyle: state.formulaMathStyle,
      formulaFontSizePoints: state.formulaFontSizePoints,
      followHostFontSize: state.followHostFontSize,
      symbolFontId: state.symbolFontId,
      numberFontFamily: state.numberFontFamily,
      cjkFontFamily: state.cjkFontFamily,
    }});
    state = saved;
    render();
    showStatus(strings.saved);
    try { Office.context.ui.messageParent(JSON.stringify({type: 'settingsSaved'})); } catch { /* opened outside Office */ }
  } catch (error) {
    showStatus(error.message, true);
    state = await loadSettings(host);
    render();
  }
}

function bind() {
  for (const button of backendButtons) {
    button.addEventListener('click', () => { state.insertionBackend = button.dataset.backend; render(); save(); });
  }
  inputs.formulaColor.addEventListener('change', () => { state.formulaColor = inputs.formulaColor.value.toUpperCase(); save(); });
  element('resetFormulaColor').addEventListener('click', () => { state.formulaColor = '#000000'; render(); save(); });
  inputs.formulaMathStyle.addEventListener('change', () => { state.formulaMathStyle = inputs.formulaMathStyle.value; save(); });
  inputs.formulaFontSizePoints.addEventListener('change', () => {
    state.formulaFontSizePoints = Number(inputs.formulaFontSizePoints.value);
    save();
  });
  for (const name of ['symbolFontId', 'numberFontFamily', 'cjkFontFamily'])
    inputs[name].addEventListener('change', () => { state[name] = inputs[name].value; save(); });
  inputs.followHostFontSize.addEventListener('change', () => {
    state.followHostFontSize = inputs.followHostFontSize.checked;
    save();
  });
  for (const [id, kind] of [['importTypography', 'import'], ['exportTypography', 'export']]) {
    element(id).addEventListener('click', async () => {
      try {
        const result = await api(`/api/settings/${host}/${kind}`, {method: 'POST', body: {}});
        if (result.canceled) return;
        if (result.settings) await populate(result.settings);
        showStatus(`${strings[kind === 'import' ? 'importedPreset' : 'exportedPreset']}${result.name || ''}`);
      } catch (error) {
        showStatus(error.message, true);
      }
    });
  }
}

applyText();
bind();
Office.onReady().then(() => loadSettings(host)).then(populate).catch(error => showStatus(error.message, true));
