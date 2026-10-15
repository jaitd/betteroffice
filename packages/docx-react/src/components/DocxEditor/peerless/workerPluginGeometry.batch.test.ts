import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import * as display from '@betteroffice/docx/layout/render';
import type { ProposalGeometryTarget } from '@betteroffice/docx/yrs';
import type { DocxGeometryTarget } from '../../../plugins/types';
import { createWorkerPluginGeometry } from './workerPluginGeometry';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const PAGE = { width: 300, height: 500 };
const GAP = 20;
const rect = (x: number, y: number, width: number, height: number) => new DOMRect(x, y, width, height);
const restore: { mockRestore(): void }[] = [];
afterEach(() => { for (const spy of restore.splice(0)) spy.mockRestore(); });

function target(paraId: string): DocxGeometryTarget {
  return { kind: 'range', version: 'v1', range: { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 1 }, view: 'accepted' } };
}

/** A two-page worker shell whose pages are built per `built`, its anchor reads held until `answer()`. */
async function shell(built: [boolean, boolean] = [true, true]) {
  const container = document.createElement('div');
  const host = document.createElement('div');
  host.setAttribute('data-peerless-pages-host', '');
  container.appendChild(host);
  host.getBoundingClientRect = () => rect(0, 0, PAGE.width, 2 * PAGE.height + GAP);
  container.getBoundingClientRect = host.getBoundingClientRect;
  const layer = document.createElement('div');
  layer.getBoundingClientRect = container.getBoundingClientRect;
  const identity = { session: 's', incarnation: 1, epoch: 1 };
  const frameOf = (epoch: number, unbuilt: readonly boolean[]) => ({ identity, coversThroughSeq: 0, frameEpoch: epoch, layoutRevision: epoch, caret: null, selectionRects: [],
    pages: unbuilt.map((pending, pageIndex) => ({ page: { pageId: String(pageIndex), pageIndex }, transform: [1, 0, 0, 1, 0, 0],
      bounds: { x: 0, y: pageIndex * (PAGE.height + GAP), ...PAGE },
      displayPage: { width: PAGE.width, height: PAGE.height, positionSpan: [pageIndex * 100, pageIndex * 100 + 99], ...(pending ? { unbuilt: true } : {}) } })) });
  const view = { identity, version: 0, throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, editState: undefined,
    geometry: frameOf(1, built.map(value => !value)) };
  const review = { version: 'v1', previewVersion: 0 };
  const reads: { targets: { range: { start: { paraId: string } } }[]; resolve(): void }[] = [];
  const answerOf = (paraId: string): ProposalGeometryTarget => {
    const from = paraId === 'late' ? 150 : 10;
    return { ok: true, ranges: [{ from, to: from + 5 }], paragraph: from } as ProposalGeometryTarget;
  };
  const controller = { getSnapshot: () => ({ status: 'ready', view }), review: async () => review, reviewSnapshot: () => review,
    readAt: (request: { targets: (typeof reads)[number]['targets'] }) => new Promise(resolve => {
      reads.push({ targets: request.targets, resolve: () => resolve(request.targets.map(({ range }) => answerOf(range.start.paraId))) });
    }),
    ref: { getZoom: () => 1 } };
  restore.push(spyOn(display, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never));
  restore.push(spyOn(display, 'createDisplayListQueries').mockImplementation(((list: { pages: { unbuilt?: boolean }[] }) => ({
    displayList: list, pageCount: () => 2, pageSize: () => PAGE,
    pageBounds: (pageIndex: number) => pageIndex < 2 ? { pageIndex, x: 0, y: 0, ...PAGE } : null,
    rangeRects: (from: number, to: number) => { const pageIndex = Math.floor(from / 100); return list.pages[pageIndex]?.unbuilt ? [] : [{ pageIndex, x: from % 100, y: 10, width: to - from, height: 12 }]; },
    anchorRect: (position: number) => ({ pageIndex: Math.floor(position / 100), x: 0, y: 0, width: 0, height: 12 }),
    dispose() {},
  })) as never));
  let published = 0;
  const plugin = createWorkerPluginGeometry(controller as never, () => container, () => layer, () => { published += 1; });
  await plugin.update();
  return {
    plugin, reads,
    published: () => published,
    async answer() { for (const read of reads.splice(0)) read.resolve(); await new Promise(resolve => setTimeout(resolve, 0)); },
    async rebuild(built: readonly boolean[]) { view.geometry = frameOf(view.geometry.frameEpoch + 1, built.map(value => !value)); await plugin.update(); },
  };
}

test('readAnchorGeometry resolves each target once the worker answers', async () => {
  const s = await shell();
  const read = s.plugin.geometry()!.readAnchorGeometry([target('a'), target('b')]);
  await Promise.resolve();
  await s.answer();
  const results = await read;
  expect(results.map(result => result.ok)).toEqual([true, true]);
  expect(results[0]).toMatchObject({ rects: [{ pageIndex: 0, x: 10, y: 10, width: 5, height: 12 }] });
});

test('a target whose worker read is in flight is pending, not layout-unavailable', async () => {
  const s = await shell();
  expect(s.plugin.geometry()!.getAnchorGeometry(target('a'))).toEqual({ ok: false, failure: { code: 'pending', message: 'The target geometry is being read' } });
  await Promise.resolve();
  await s.answer();
  expect(s.plugin.geometry()!.getAnchorGeometry(target('a')).ok).toBe(true);
});

test('targets asked in one tick share one worker read and one publish', async () => {
  const s = await shell();
  const geometry = s.plugin.geometry()!;
  for (const paraId of ['a', 'b', 'c', 'b']) geometry.getAnchorGeometry(target(paraId));
  await Promise.resolve();
  expect(s.reads.map(read => read.targets.map(({ range }) => range.start.paraId))).toEqual([['a', 'b', 'c']]);
  const before = s.published();
  await s.answer();
  expect(s.published() - before).toBe(1);
  const read = geometry.readAnchorGeometry([target('b'), target('d'), target('e')]);
  await Promise.resolve();
  expect(s.reads.map(read => read.targets.map(({ range }) => range.start.paraId))).toEqual([['d', 'e']]);
  await s.answer();
  expect((await read).every(result => result.ok)).toBe(true);
  expect(s.published() - before).toBe(2);
});

test('answers outlive a new frame at the same version and stay provisional until their page builds', async () => {
  const s = await shell([true, false]);
  s.plugin.geometry()!.getAnchorGeometry(target('late'));
  await Promise.resolve();
  await s.answer();
  expect(s.plugin.geometry()!.getAnchorGeometry(target('late'))).toMatchObject({ ok: true, rects: [], unbuiltPages: [1] });
  await s.rebuild([true, true]);
  const built = s.plugin.geometry()!.getAnchorGeometry(target('late'));
  await Promise.resolve();
  expect(s.reads).toHaveLength(0);
  expect(built).toMatchObject({ ok: true, rects: [{ pageIndex: 1 }] });
  expect(built.ok && built.unbuiltPages).toBeUndefined();
});

test('getPageRect answers in overlay pixels and says whether the page is built', async () => {
  const s = await shell([true, false]);
  const geometry = s.plugin.geometry()!;
  expect(geometry.getPageRect(0)).toEqual({ pageIndex: 0, x: 0, y: 0, ...PAGE, built: true });
  expect(geometry.getPageRect(1)).toEqual({ pageIndex: 1, x: 0, y: PAGE.height + GAP, ...PAGE, built: false });
  expect(geometry.getPageRect(2)).toBeNull();
});
