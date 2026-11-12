// Ribbon commands, task pane, dialogs and OCR for Word and PowerPoint on Mac.
// Runs in the shared runtime that taskpane.html hosts for the whole session.
/* global Office */
import {api, ApiError, catalog, loadSettings} from './api.mjs';
import {createMetadata, MetadataError} from './metadata.mjs';
import {base64Utf8, rasterize, renderMathMl, renderSvg, startRenderer} from './renderer.mjs';
import * as word from './word.mjs';
import * as powerpoint from './powerpoint.mjs';
import {text} from './strings.mjs';

const OCR_TIMEOUT_SECONDS = 120;
const OCR_POLL_MS = 250;
const DIALOG_CLOSE_SETTLE_MS = 1000;

export class Controller {
  constructor(hostName, pane) {
    this.host = hostName === Office.HostType.Word ? 'word' : 'powerpoint';
    this.adapter = this.host === 'word' ? word : powerpoint;
    this.locale = String(Office.context.displayLanguage || 'zh-CN');
    this.t = key => text(this.locale, this.host, key);
    this.pane = pane;
    this.busy = false;
    this.ocr = null;
    this.dialog = null;
    this.session = 0;
    this.pendingEdit = null;
  }

  start() {
    const commands = {
      insertFormula: () => this.openEditor(),
      screenshotOcr: () => this.toggleOcr(),
      editSelected: () => this.editSelected(),
      deleteSelected: () => this.deleteSelected(),
      formatSelected: () => this.format(false),
      formatAll: () => this.format(true),
      openSettings: () => this.openDialog('settings.html', 50, 70),
      openHelp: () => this.openDialog('help.html', 55, 75, false),
    };
    for (const [name, command] of Object.entries(commands)) {
      Office.actions.associate(name, event => {
        // Commands only start work; long operations report through the pane.
        Promise.resolve().then(command).catch(error => this.fail(error));
        event.completed();
      });
    }
    this.pane.configure({
      strings: this.paneStrings(),
      locale: this.locale,
      onInsert: state => this.run(() => this.insertDirect(state)),
      onConnect: () => this.run(() => this.testConnection()),
      onOcr: () => this.toggleOcr(),
    });
    this.status('info', this.t('ReadyStatus'));
    startRenderer().catch(error => this.fail(error));
  }

  paneStrings() {
    const keys = {officePlugin: 'OfficePluginLabel', connect: 'ConnectButton', equation: 'EquationLabel',
      display: 'DisplayLabel', screenshotOcr: 'ScreenshotOcrButton', cancelOcr: 'CancelOcrButton', insert: 'EditorInsert'};
    return Object.fromEntries(Object.entries(keys).map(([key, id]) => [key, this.t(id)]));
  }

  // Status -------------------------------------------------------------------------

  status(kind, message) {
    this.pane.status({kind, message, busy: this.busy, ocrActive: Boolean(this.ocr)});
  }

  fail(error) {
    const message = error instanceof ApiError || error instanceof MetadataError
      || error instanceof word.WordFormulaError || error instanceof powerpoint.PowerPointFormulaError
      ? error.message : (error?.message ? `${this.t('UnexpectedError')}${error.message}` : this.t('UnexpectedError'));
    this.status('error', message);
    Office.addin.showAsTaskpane().catch(() => {});
  }

  /** Runs one document operation at a time; errors propagate to the caller. */
  async exclusive(operation) {
    if (this.busy) throw new MetadataError(this.t('BusyStatus'));
    this.busy = true;
    this.status('info', this.t('WorkingStatus'));
    try {
      return await operation();
    } finally {
      this.busy = false;
      // Never leave "Working..." behind if the operation reported nothing else.
      if (this.pane.lastStatus.message === this.t('WorkingStatus')) this.status('info', this.t('ReadyStatus'));
      else this.pane.status({busy: false, ocrActive: Boolean(this.ocr)});
    }
  }

  /** Like exclusive(), but reports failures in the task pane. */
  async run(operation) {
    try {
      await this.exclusive(operation);
    } catch (error) {
      this.fail(error);
    }
  }

  // Formula construction -----------------------------------------------------------

  async defaults() {
    const settings = await loadSettings(this.host);
    return {backend: settings.insertionBackend, typography: settings.typography, follow: settings.followHostFontSize};
  }

  /** New-formula typography, following the host text size when configured. */
  async newTypography() {
    const {typography, follow} = await this.defaults();
    if (!follow) return typography;
    const size = await this.adapter.selectionFontSize().catch(() => null);
    return size && size >= 1 && size <= 1638 ? {...typography, fontSizePoints: size} : typography;
  }

  async build({latex, display, typography, existing}) {
    if (!String(latex || '').trim()) throw new MetadataError(this.t('LatexRequired'));
    const documentId = existing?.documentId || await this.adapter.documentId();
    const {backend} = await this.defaults();
    const equationId = existing?.equationId;
    if (this.host === 'word') return this.buildWord({latex, display, typography, documentId, equationId, backend, existing});
    return this.buildPowerPoint({latex, typography, documentId, equationId, backend, existing});
  }

  async buildWord({latex, display, typography, documentId, equationId, backend, existing}) {
    // An edit keeps the formula's own kind; only new formulas follow the setting.
    const engine = existing ? existing.renderEngine : backend === 'Omml' ? 'Omml' : 'Image';
    if (engine === 'Omml') {
      const mathml = await renderMathMl(latex, display, typography);
      const {omml} = await api('/api/omml', {method: 'POST', body: {mathml, typography}});
      const metadata = createMetadata({documentId, equationId, latex, display, renderEngine: 'Omml', typography});
      return {metadata, write: target => word.writeFormula(metadata, (current, options) => word.ommlPackage(current, omml, options), target)};
    }
    const rendered = await renderSvg(latex, display, typography);
    const png = await rasterize(rendered);
    const metadata = createMetadata({documentId, equationId, latex, display, renderEngine: 'Image', typography,
      naturalWidthPoints: rendered.widthPoints, naturalHeightPoints: rendered.heightPoints});
    return {metadata, rendered,
      write: target => word.writeFormula(metadata,
        (current, options) => word.picturePackage(current, rendered, png, base64Utf8(rendered.svg), options), target)};
  }

  async buildPowerPoint({latex, typography, documentId, equationId, backend, existing}) {
    const rendered = await renderSvg(latex, true, typography);
    // An edit keeps the formula's own picture format; only new formulas follow the setting.
    const svg = existing ? existing.imageFormat !== 'Png' : backend !== 'Png';
    const image = {svg: svg ? rendered.svg : null, base64: svg ? null : await rasterize(rendered),
      widthPoints: rendered.widthPoints, heightPoints: rendered.heightPoints};
    const metadata = {...createMetadata({documentId, equationId, latex, display: true, renderEngine: 'Image',
      typography, naturalWidthPoints: rendered.widthPoints, naturalHeightPoints: rendered.heightPoints}),
    imageFormat: svg ? 'Svg' : 'Png'};
    return {metadata, rendered,
      write: (target, options) => powerpoint.writeFormula(metadata, image, target, options)};
  }

  // Commands -----------------------------------------------------------------------

  async insertDirect(state) {
    const typography = await this.newTypography();
    const formula = await this.build({latex: state.latex, display: this.host === 'powerpoint' || state.display, typography});
    await formula.write(null);
    this.status('success', this.t('InsertedStatus'));
  }

  async editSelected() {
    const metadata = await this.adapter.selectedFormula();
    this.pane.load(metadata.latex, metadata.displayMode === 'Display');
    await this.openEditor(metadata);
    this.status('success', this.t('LoadedStatus'));
  }

  async deleteSelected() {
    await this.run(async () => {
      const count = await this.adapter.deleteSelected();
      this.status('success', count > 1 ? this.t('DeletedManyStatus').replace('{count}', count) : this.t('DeletedStatus'));
    });
  }

  async format(all) {
    await this.run(async () => {
      const formulas = all ? await this.adapter.allFormulas() : await this.adapter.selectedFormulas();
      if (!formulas.length) {
        this.status(all ? 'info' : 'error', all ? this.t('NoFormattingNeededStatus') : this.t('SelectedFormulaRequired'));
        return;
      }
      const {typography} = await this.defaults();
      let succeeded = 0;
      let firstError = '';
      for (const [index, existing] of formulas.entries()) {
        this.status('info', this.t('BatchFormattingStatus').replace('{processed}', index).replace('{total}', formulas.length));
        try {
          const formula = await this.build({latex: existing.latex, display: existing.displayMode === 'Display',
            typography, existing});
          if (this.host === 'powerpoint') {
            await powerpoint.selectShape(existing.slideId, existing.shapeId);
            await formula.write(existing, {naturalSize: true});
          } else {
            await formula.write(existing.equationId);
          }
          succeeded++;
        } catch (error) {
          firstError ||= error?.message || String(error);
        }
      }
      const failed = formulas.length - succeeded;
      this.status(failed ? 'error' : 'success', failed
        ? this.t('FormattedWithFailuresStatus').replace('{total}', formulas.length).replace('{succeeded}', succeeded)
          .replace('{failed}', failed).replace('{reason}', firstError)
        : this.t('FormattedStatus').replace('{count}', succeeded));
    });
  }

  async testConnection() {
    await api('/api/recognition/config');
    this.status('success', this.t('ConnectedAutomationStatus'));
  }

  // OCR ----------------------------------------------------------------------------

  async toggleOcr() {
    if (this.ocr) {
      const {jobId, controller} = this.ocr;
      this.ocr = null;
      controller.abort();
      if (jobId) await api(`/api/recognition/jobs/${encodeURIComponent(jobId)}`, {method: 'DELETE'}).catch(() => {});
      this.status('info', this.t('OcrCanceledStatus'));
      return;
    }
    const state = {jobId: null, controller: new AbortController()};
    this.ocr = state;
    this.status('info', this.t('OcrWaitingStatus'));
    try {
      const latex = await this.waitForOcr(state);
      if (this.ocr !== state) return;
      this.ocr = null;
      this.pane.load(latex, this.host === 'powerpoint');
      this.status('success', this.t('OcrLoadedStatus'));
      Office.addin.showAsTaskpane().catch(() => {});
    } catch (error) {
      if (this.ocr !== state || error?.name === 'AbortError') return;
      this.ocr = null;
      this.fail(error);
    }
  }

  async waitForOcr(state) {
    let created;
    try {
      created = await api('/api/recognition/next', {method: 'POST', body: {timeout: OCR_TIMEOUT_SECONDS}, signal: state.controller.signal});
    } catch (error) {
      if (error?.code !== 'next_result_busy') throw error;
      throw new ApiError(409, error.code, this.t('AutomationOcrAlreadyWaiting'));
    }
    state.jobId = created?.job?.id;
    if (!state.jobId) throw new ApiError(502, 'invalid_response', this.t('InvalidResponse'));
    const deadline = Date.now() + (OCR_TIMEOUT_SECONDS + 10) * 1000;
    let announced = false;
    while (Date.now() < deadline) {
      await new Promise(resolve => setTimeout(resolve, OCR_POLL_MS));
      if (this.ocr !== state) throw new DOMException('canceled', 'AbortError');
      const {job} = await api(`/api/recognition/jobs/${encodeURIComponent(state.jobId)}`, {signal: state.controller.signal});
      if (job.state === 'completed') {
        const item = (job.items || []).find(entry => String(entry.text || '').trim());
        if (!item) throw new ApiError(422, 'empty_formula', this.t('OcrEmpty'));
        return item.text;
      }
      if (job.state === 'failed' || job.state === 'canceled')
        throw new ApiError(422, job.error?.code || job.state, job.error?.message || this.t('OcrFailed'));
      if (!announced && (job.state === 'queued' || job.state === 'running')) {
        announced = true;
        this.status('info', this.t('OcrRecognizingStatus'));
      }
    }
    throw new ApiError(504, 'timeout', this.t('OcrTimeout'));
  }

  // Dialogs ------------------------------------------------------------------------

  async openEditor(existing = null) {
    if (this.busy) throw new MetadataError(this.t('BusyStatus'));
    const typography = existing ? existing.typography : await this.newTypography();
    const display = this.host === 'powerpoint' || (existing ? existing.displayMode === 'Display' : this.pane.state().display);
    this.session += 1;
    this.pendingEdit = existing;
    const init = {
      type: 'init',
      session: this.session,
      typography,
      catalog: await catalog(),
      referencePreview: existing ? existing.renderEngine === 'Omml'
        : this.host === 'word' && (await this.defaults()).backend === 'Omml',
      latex: existing ? existing.latex : this.pane.state().latex,
      display,
      forceDisplay: this.host === 'powerpoint',
      mode: existing ? 'update' : 'insert',
      locale: this.locale,
    };
    await this.openDialog('editor.html', 62, 78, true, init);
  }

  async onEditorAccept(dialog, message) {
    if (message.session !== this.session) return;
    const existing = this.pendingEdit;
    try {
      await this.exclusive(async () => {
        const formula = await this.build({latex: message.latex, display: message.display,
          typography: message.typography, existing});
        if (this.host === 'powerpoint') await formula.write(existing);
        else await formula.write(existing ? existing.equationId : null);
      });
    } catch (error) {
      this.status('error', error?.message || this.t('UnexpectedError'));
      dialog.messageChild(JSON.stringify({type: 'submitFailed', session: message.session,
        message: error?.message || this.t('UnexpectedError')}));
      return;
    }
    this.pane.load(message.latex, message.display);
    this.status('success', this.t(existing ? 'UpdatedStatus' : 'InsertedStatus'));
    this.closeDialog();
  }

  async openDialog(page, width, height, interactive = true, init = null) {
    const replacing = Boolean(this.dialog);
    if (replacing) {
      // Office for Mac reports no close event and may accept a new dialog
      // before the old one is gone, orphaning it; give it time to close.
      this.closeDialog();
      await new Promise(resolve => setTimeout(resolve, DIALOG_CLOSE_SETTLE_MS));
    }
    const url = new URL(`./${page}`, location.href);
    url.search = new URLSearchParams({host: this.host, locale: this.locale}).toString();
    // Office needs a moment to tear down a dialog before it accepts another
    // (error 12007), so retry briefly instead of failing the command.
    const DIALOG_ALREADY_OPEN = 12007;
    let result;
    for (let attempt = 0; ; attempt++) {
      result = await new Promise(resolve =>
        Office.context.ui.displayDialogAsync(url.href, {width, height, displayInIframe: false}, resolve));
      if (result.status === Office.AsyncResultStatus.Succeeded) break;
      if (result.error?.code !== DIALOG_ALREADY_OPEN || attempt >= (replacing ? 20 : 3))
        throw new MetadataError(this.t('DialogFailed'));
      await new Promise(resolve => setTimeout(resolve, 150));
    }
    const dialog = result.value;
    this.dialog = dialog;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
      let message;
      try { message = JSON.parse(event.message); } catch { return; }
      if (message.type === 'ready' && init) dialog.messageChild(JSON.stringify(init));
      else if (message.type === 'accept') this.onEditorAccept(dialog, message);
      else if (message.type === 'cancel' || message.type === 'close') this.closeDialog();
      else if (message.type === 'settingsSaved') this.status('success', this.t('SettingsSavedStatus'));
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
      if (this.dialog === dialog) this.dialog = null;
    });
    if (!interactive) this.status('info', this.t('HelpStatus'));
  }

  closeDialog() {
    const dialog = this.dialog;
    this.dialog = null;
    try { dialog?.close(); } catch { /* already closed */ }
  }
}
