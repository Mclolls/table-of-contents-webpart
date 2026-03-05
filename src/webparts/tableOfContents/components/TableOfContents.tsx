import * as React from 'react';
import { useEffect, useState, useCallback, useRef } from 'react';
import { getTheme } from '@fluentui/react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './TableOfContents.module.scss';

type HeadingItem = { id: string; text: string; level: number; children: HeadingItem[]; };
type SectionGroup = { sectionElement: HTMLElement; sectionIndex: number; sectionTitle: string; items: HeadingItem[]; };

export interface ITableOfContentsProps {
  levels: number[];
  listType: 'bulleted' | 'numbered';
  listStyle: string;
  title?: string;
  columns: number;
  minColumnWidth: number;
  themeVariant?: IReadonlyTheme;
  debugHeadings?: boolean;
  scope?: 'page' | 'section';
  groupBySection?: boolean;
  showSectionHeadings?: boolean;
  titleColor?: string;
  itemsColor?: string;
  titleLevel?: 'h2' | 'h3' | 'h4';
}

function slugify(text: string): string {
  return (text ?? '').toLowerCase().trim().replace(/[^\w\s-]/g, '').replace(/\s+/g, '-').replace(/-+/g, '-');
}

function buildHierarchy(flat: HeadingItem[]): HeadingItem[] {
  const root: HeadingItem[] = [];
  const stack: HeadingItem[] = [];
  flat.forEach((h) => {
    while (stack.length && stack[stack.length - 1].level >= h.level) stack.pop();
    if (stack.length === 0) { root.push(h); stack.push(h); }
    else { stack[stack.length - 1].children.push(h); stack.push(h); }
  });
  return root;
}

function ensureUniqueId(baseText: string): string {
  const base = slugify(baseText) || 'heading';
  let id = base;
  let i = 1;
  while (document.getElementById(id)) {
    id = `${base}-${i++}`;
  }
  return id;
}

export const TableOfContents: React.FC<ITableOfContentsProps> = ({
  levels, listType, listStyle, title, columns, minColumnWidth, themeVariant,
  debugHeadings, scope, groupBySection, showSectionHeadings, titleColor, itemsColor, titleLevel
}) => {
  const [toc, setToc] = useState<HeadingItem[]>([]);
  const [groups, setGroups] = useState<SectionGroup[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);

  const navRef = useRef<HTMLElement | null>(null);
  const scrollObserverRef = useRef<IntersectionObserver | null>(null);

  const theme = getTheme();
  const sc = themeVariant?.semanticColors ?? theme.semanticColors;

  // Cast styles to indexable record to avoid missing-class typing errors from generated d.ts
  const s = styles as unknown as Record<string, string>;

  // Build inline style object that sets CSS variables used by the module.scss
  // Use a type that allows setting custom properties
  const containerStyle = {} as React.CSSProperties & Record<string, string>;

  // Prefer explicit user colors; if not provided, fall back to theme tokens or semantic colors
  const linkFallback = itemsColor ?? themeVariant?.palette?.themePrimary ?? sc.link ?? '';
  const titleFallback = titleColor ?? themeVariant?.semanticColors?.bodyText ?? themeVariant?.palette?.themePrimary ?? sc.bodyText ?? '';

  containerStyle['--toc-columns'] = String(columns ?? 1);
  containerStyle['--toc-column-width'] = `${minColumnWidth ?? 150}px`;
  containerStyle['--toc-column-gap'] = '16px';

  if (linkFallback) {
    containerStyle['--toc-link-color'] = linkFallback;
  }
  if (titleFallback) {
    containerStyle['--toc-title-color'] = titleFallback;
  }

  const onNavigate = (id: string): void => {
    const el = document.getElementById(id);
    if (el) el.scrollIntoView({ behavior: 'smooth' });
  };

  const renderItems = useCallback((items: HeadingItem[], depth = 0): JSX.Element => {
    const Tag = listType === 'numbered' ? 'ol' : 'ul';
    const isHierarchical = listStyle === 'decimal' && listType === 'numbered';

    const hierarchicalClass = isHierarchical ? (s.hierarchical ?? '') : '';

    const listClassName = depth === 0
      ? [s.listRoot ?? '', hierarchicalClass].filter(Boolean).join(' ')
      : (s.listNested ?? s.listRoot ?? '');

    // Inline style only for the top-level list to guarantee column behavior and to preserve bullets when needed
    const topLevelInlineStyle: React.CSSProperties | undefined = depth === 0 ? {
      columnCount: columns ?? 1,
      columnGap: (containerStyle && (containerStyle['--toc-column-gap'] as string)) ?? '16px',
      columnWidth: `${minColumnWidth ?? 150}px`,
      color: itemsColor ?? undefined,
      // For non-hierarchical lists, explicitly set the chosen list style so bullets appear.
      listStyleType: !isHierarchical ? (listStyle as React.CSSProperties['listStyleType']) : 'none'
    } : undefined;

    return (
      <Tag
        className={listClassName}
        style={topLevelInlineStyle}
      >
        {items.map(item => (
          <li key={item.id}>
            <a
              href={`#${item.id}`}
              className={[s.link ?? '', activeId === item.id ? (s.active ?? '') : ''].filter(Boolean).join(' ')}
              aria-current={activeId === item.id ? 'true' : undefined}
              data-toc-active={activeId === item.id ? 'true' : undefined}
              onClick={(e) => { e.preventDefault(); onNavigate(item.id); }}
            >
              {item.text}
            </a>
            {item.children.length > 0 && renderItems(item.children, depth + 1)}
          </li>
        ))}
      </Tag>
    );
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeId, itemsColor, listType, listStyle, columns, minColumnWidth]);

  useEffect(() => {
    let timeoutId: number;

    type HeadingEntry = { el: HTMLElement; id: string; text: string; level: number; };

    const setupIntersectionObserver = (items: HeadingItem[]): void => {
      if (scrollObserverRef.current) scrollObserverRef.current.disconnect();
      scrollObserverRef.current = new IntersectionObserver((entries) => {
        entries.forEach(entry => { if (entry.isIntersecting) setActiveId(entry.target.id); });
      }, { rootMargin: '0px 0px -75% 0px', threshold: 0.1 });

      items.forEach(i => {
        const el = document.getElementById(i.id);
        if (el) scrollObserverRef.current?.observe(el);
      });
    };

    const isVisible = (el: HTMLElement): boolean => {
      try {
        const style = window.getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 && rect.height === 0) return false;
        return true;
      } catch {
        return false;
      }
    };

    const scan = (): void => {
      if (!navRef.current) return;
      const allowedLevels = new Set<number>(levels.length ? levels : [2, 3, 4]);

      const getCleanText = (el: HTMLElement): string => (el.innerText || el.textContent || '').trim();

      const collectFromZone = (zone: HTMLElement): HeadingEntry[] => {
        const nodeList = zone.querySelectorAll('h2, h3, h4, [role="heading"]');
        const entries: HeadingEntry[] = [];
        (Array.from(nodeList) as HTMLElement[]).forEach(h => {
          if (navRef.current?.contains(h)) return;
          if (!isVisible(h)) {
            if (debugHeadings) console.debug('[TOC] skipping invisible heading', h);
            return;
          }
          if (h.classList.contains('ms-srOnly') || h.getAttribute('aria-hidden') === 'true' || h.closest('.ms-Editor-host')) return;

          const text = getCleanText(h);
          if (!text) return;

          const tag = h.tagName.toLowerCase();
          const lvl = /^h[2-4]$/.test(tag) ? parseInt(tag[1], 10) : parseInt(h.getAttribute('aria-level') || '0', 10);
          if (!allowedLevels.has(lvl)) return;

          if (!h.id) {
            h.id = ensureUniqueId(text);
          } else {
            const existing = document.querySelectorAll(`#${CSS.escape(h.id)}`);
            if (existing.length > 1) {
              h.id = ensureUniqueId(text);
            }
          }

          entries.push({ el: h, id: h.id, text, level: lvl });
        });
        return entries;
      };

      const sectionSelector = '[data-automation-id*="CanvasZone"], [data-automation-id="CanvasSection"], .CanvasZone';
      const mySection = navRef.current.closest(sectionSelector) as HTMLElement | null;

      let allEntries: HeadingEntry[] = [];
      if (scope === 'section' && mySection) {
        const entries = collectFromZone(mySection);
        allEntries = entries;
        setGroups([{ sectionElement: mySection, sectionIndex: 0, sectionTitle: 'Current Section', items: entries.map(e => ({ id: e.id, text: e.text, level: e.level, children: [] })) }]);
      } else {
        const zones = Array.from(document.querySelectorAll(sectionSelector)) as HTMLElement[];
        const groupsCollected = zones.map((z, i) => {
          const entries = collectFromZone(z);
          return { sectionElement: z, sectionIndex: i, sectionTitle: `Section ${i + 1}`, entries };
        }).filter(g => g.entries.length > 0);

        groupsCollected.forEach(g => allEntries.push(...g.entries));

        const dedupedGroups: SectionGroup[] = groupsCollected.map(g => {
          const seen = new Set<HTMLElement>();
          const dedupedItems = g.entries.filter(e => {
            if (seen.has(e.el)) return false;
            seen.add(e.el);
            return true;
          }).map(e => ({ id: e.id, text: e.text, level: e.level, children: [] as HeadingItem[] }));
          return { sectionElement: g.sectionElement, sectionIndex: g.sectionIndex, sectionTitle: g.sectionTitle, items: dedupedItems };
        }).filter(g => g.items.length > 0);

        setGroups(dedupedGroups);
      }

      // Deduplicate by element reference
      const seenElements = new Set<HTMLElement>();
      const uniqueEntries: HeadingEntry[] = [];
      allEntries.forEach(e => {
        if (!seenElements.has(e.el)) {
          seenElements.add(e.el);
          uniqueEntries.push(e);
        } else {
          if (debugHeadings) console.debug('[TOC] duplicate element ignored', e.el);
        }
      });

      const finalItems: HeadingItem[] = uniqueEntries.map(e => ({ id: e.id, text: e.text, level: e.level, children: [] }));

      setToc(buildHierarchy(finalItems));
      setupIntersectionObserver(finalItems);
    };

    const debouncedScan = (): void => {
      window.clearTimeout(timeoutId);
      timeoutId = window.setTimeout(scan, 400);
    };

    const mutObserver = new MutationObserver((mutations) => {
      const isExternal = mutations.some(m => !navRef.current?.contains(m.target));
      if (isExternal) debouncedScan();
    });

    mutObserver.observe(document.body, { childList: true, subtree: true });
    scan();

    return () => {
      window.clearTimeout(timeoutId);
      mutObserver.disconnect();
      if (scrollObserverRef.current) scrollObserverRef.current.disconnect();
    };
  }, [levels, scope, groupBySection, showSectionHeadings, title, debugHeadings]);

  return (
    <nav
      ref={navRef}
      className={s.root ?? ''}
      style={containerStyle}
    >
      {title && (
        <>
          {titleLevel === 'h2' ? (
            <h2 className={s.title ?? ''}>{title}</h2>
          ) : titleLevel === 'h3' ? (
            <h3 className={s.title ?? ''}>{title}</h3>
          ) : titleLevel === 'h4' ? (
            <h4 className={s.title ?? ''}>{title}</h4>
          ) : (
            <strong className={s.title ?? ''}>{title}</strong>
          )}
        </>
      )}

      {groupBySection ? (
        groups.map(g => (
          <div key={g.sectionIndex} className={s.sectionWrapper ?? ''}>
            {showSectionHeadings && <strong>{g.sectionTitle}</strong>}
            {renderItems(buildHierarchy(g.items))}
          </div>
        ))
      ) : (
        toc.length > 0 ? renderItems(toc) : <div>No headings found.</div>
      )}
    </nav>
  );
};