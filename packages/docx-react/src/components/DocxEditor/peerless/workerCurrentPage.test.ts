import { expect, mock, test } from 'bun:test';
import { createElement } from 'react';
import { renderToString } from 'react-dom/server';
import { readWorkerSection, type WorkerShellController } from './workerProperties';
import { useWorkerRulerControls } from './workerPropertyControls';


function controllerOf(status: string) {
  const getCurrentPage = mock(() => 3);
  const pages = Array.from({ length: 40 }, (_, pageIndex) => ({ page: { pageIndex }, displayPage: { sectionIndex: 0 } }));
  const snapshot = { status, view: { geometry: { pages } } };
  const controller = { getCurrentPage, value: {
    query: async (request: { kind: string }) => request.kind === 'partSnapshot' ? { value: { metadata: { body: { sections: [{}] } } } } : { value: request },
    ref: { getCurrentPage }, getSnapshot: () => snapshot, subscribe: () => () => {},
    environment: () => ({ readOnly: true, mode: 'viewing' }),
  } as unknown as WorkerShellController };
  return controller;
}

test('reading the current section measures the current page once', async () => {
  const { value, getCurrentPage } = controllerOf('ready');
  await readWorkerSection(value);
  expect(getCurrentPage).toHaveBeenCalledTimes(1);
});

test('a ruler render measures the current page at most once', () => {
  const { value, getCurrentPage } = controllerOf('loading');
  renderToString(createElement(function Ruler() { useWorkerRulerControls(value); return null; }));
  expect(getCurrentPage.mock.calls.length).toBeLessThanOrEqual(1);
});
