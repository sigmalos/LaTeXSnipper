// Hosts the shared task pane (taskpane-app.mjs) without WebView2: the page's
// window.chrome.webview.postMessage calls arrive here and replies go through
// window.LaTeXSnipperTaskPane.apply, exactly as on Windows.
export class Pane {
  constructor() {
    this.current = {latex: '', display: false};
    this.lastStatus = {kind: 'info', message: '', busy: false, ocrActive: false};
    this.handlers = {};
    globalThis.chrome = {...(globalThis.chrome || {}), webview: {postMessage: message => this.receive(message)}};
  }

  async mount() {
    const {startTaskPane} = await import('/shared/taskpane-app.mjs');
    startTaskPane();
  }

  configure({strings, locale, ...handlers}) {
    this.handlers = handlers;
    this.apply({type: 'state', strings, locale, latex: this.current.latex, display: this.current.display});
  }

  receive(message) {
    if (!message || typeof message !== 'object') return;
    const {type, ...state} = message;
    this.current = {...this.current, latex: state.latex ?? this.current.latex,
      display: state.display ?? this.current.display};
    if (type === 'insert') this.handlers.onInsert?.(this.state());
    else if (type === 'connect') this.handlers.onConnect?.();
    else if (type === 'ocr') this.handlers.onOcr?.();
  }

  apply(payload) {
    if (globalThis.LaTeXSnipperTaskPane) globalThis.LaTeXSnipperTaskPane.apply(payload);
    else (globalThis.__latexSnipperTaskPanePending ||= []).push(payload);
  }

  state() {
    return {...this.current};
  }

  load(latex, display) {
    this.current = {latex, display: Boolean(display)};
    this.apply({type: 'state', latex, display: Boolean(display)});
  }

  /** Partial updates keep the previous message, so clearing busy does not erase it. */
  status(update) {
    this.lastStatus = {...this.lastStatus, ...update};
    this.apply({type: 'status', ...this.lastStatus});
  }
}
