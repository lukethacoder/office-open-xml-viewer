import type { MatchRunSlice } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer.js';
import type { DocxStorySource } from './types.js';

/** Exact UTF-16 code-unit position within one rendered DOCX body paragraph. */
export interface DocxSourceTextPosition {
  readonly paragraphId: string;
  readonly offset: number;
}

/** Application-owned highlight bound to one immutable DOCX source range. */
export interface DocxSourceRangeHighlight {
  readonly id: string;
  readonly start: Readonly<DocxSourceTextPosition>;
  readonly end: Readonly<DocxSourceTextPosition>;
  readonly exact: string;
  readonly color?: string;
  readonly label?: string;
  readonly onActivate?: (id: string) => void;
}

/** Resolution outcome for one application-owned source range. */
export type DocxSourceRangeHighlightResolution =
  | { readonly id: string; readonly status: 'resolved' }
  | {
      readonly id: string;
      readonly status: 'unresolved';
      readonly reason: 'source-not-found' | 'text-mismatch';
    };

/** Render-ready source-range slices grouped by zero-based page index. */
export type DocxSourceRangeProjection = {
  readonly pages: ReadonlyMap<number, readonly DocxSourceRangePageMatch[]>;
  readonly resolutions: readonly DocxSourceRangeHighlightResolution[];
};

/** One resolved source range on one rendered page. */
export type DocxSourceRangePageMatch = {
  readonly id: string;
  readonly slices: MatchRunSlice[];
  readonly color?: string;
  readonly label?: string;
  readonly onActivate?: (id: string) => void;
};

type PageRuns = {
  readonly page: number;
  readonly runs: readonly Readonly<DocxTextRunInfo>[];
};

type LocatedRun = {
  readonly page: number;
  readonly runIndex: number;
  readonly run: Readonly<DocxTextRunInfo>;
  readonly start: number;
  readonly end: number;
};

/** Resolves source ranges against rendered runs without retaining vendor DOM identities. */
export function projectDocxSourceRangeHighlights(
  pages: readonly PageRuns[],
  highlights: readonly Readonly<DocxSourceRangeHighlight>[],
): DocxSourceRangeProjection {
  const locatedRuns = locateRuns(pages);
  const projectedPages = new Map<number, DocxSourceRangePageMatch[]>();
  const resolutions = highlights.map((highlight) => {
    const candidateRuns = locatedRuns.filter(({ run }) => run.source?.story === 'body');
    const startIndex = candidateRuns.findIndex(({ run, start, end }) =>
      positionInRun(highlight.start, run, start, end, false));
    const endIndex = candidateRuns.findIndex(({ run, start, end }, index) =>
      index >= startIndex && positionInRun(highlight.end, run, start, end, true));
    if (startIndex < 0 || endIndex < startIndex) {
      return unresolved(highlight.id, 'source-not-found');
    }

    const selected = candidateRuns.slice(startIndex, endIndex + 1).flatMap((located, index) => {
      const start = index === 0 ? highlight.start.offset - located.start : 0;
      const end = index === endIndex - startIndex
        ? highlight.end.offset - located.start
        : located.run.text.length;
      return end > start ? [{ located, start, end }] : [];
    });
    if (selected.length === 0 ||
      selected.map(({ located, start, end }) => located.run.text.slice(start, end)).join('') !== highlight.exact) {
      return unresolved(highlight.id, 'text-mismatch');
    }

    const slicesByPage = new Map<number, MatchRunSlice[]>();
    for (const { located, start, end } of selected) {
      const slices = slicesByPage.get(located.page) ?? [];
      slices.push({ runIndex: located.runIndex, start, end });
      slicesByPage.set(located.page, slices);
    }
    for (const [page, slices] of slicesByPage) {
      const matches = projectedPages.get(page) ?? [];
      matches.push({
        id: highlight.id,
        slices,
        color: highlight.color,
        label: highlight.label,
        onActivate: highlight.onActivate,
      });
      projectedPages.set(page, matches);
    }
    return { id: highlight.id, status: 'resolved' } as const;
  });
  return { pages: projectedPages, resolutions };
}

/** Assigns paragraph-relative offsets to rendered run fragments in document order. */
function locateRuns(pages: readonly PageRuns[]): LocatedRun[] {
  const offsets = new Map<string, number>();
  return pages.flatMap(({ page, runs }) => runs.flatMap((run, runIndex) => {
    if (!run.source || !run.paragraphId || run.text.length === 0) return [];
    const key = positionKey(run.source, run.paragraphId);
    const start = offsets.get(key) ?? 0;
    const end = start + run.text.length;
    offsets.set(key, end);
    return [{ page, runIndex, run, start, end }];
  }));
}

/** Checks whether one exclusive or inclusive endpoint belongs to a rendered run. */
function positionInRun(
  position: Readonly<DocxSourceTextPosition>,
  run: Readonly<DocxTextRunInfo>,
  start: number,
  end: number,
  inclusiveEnd: boolean,
): boolean {
  return run.source?.story === 'body' && run.paragraphId === position.paragraphId &&
    position.offset >= start &&
    (position.offset < end || (inclusiveEnd && position.offset === end));
}

/** Returns stable key for one source paragraph. */
function positionKey(source: Readonly<DocxStorySource>, paragraphId: string): string {
  return `${source.story}\u0000${source.storyInstance}\u0000${source.path.join('.')}\u0000${paragraphId}`;
}

/** Creates one typed unresolved outcome. */
function unresolved(
  id: string,
  reason: 'source-not-found' | 'text-mismatch',
): DocxSourceRangeHighlightResolution {
  return { id, status: 'unresolved', reason };
}
