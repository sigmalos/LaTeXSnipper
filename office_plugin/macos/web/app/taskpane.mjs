// Task pane page and shared runtime: ribbon commands run here even while the
// pane is hidden (manifest Runtime lifetime="long").
/* global Office */
import {Controller} from './controller.mjs';
import {Pane} from './pane.mjs';

const pane = new Pane();

Office.onReady(async info => {
  const word = info.host === Office.HostType.Word;
  document.body.dataset.host = word ? 'word' : 'powerpoint';
  for (const element of document.querySelectorAll('.word-only')) element.hidden = !word;
  await pane.mount();
  new Controller(info.host, pane).start();
});
