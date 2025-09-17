'use client';
import { MagnifyingGlassIcon } from '@heroicons/react/24/outline';
import { usePathname, useRouter, useSearchParams } from 'next/navigation';
import { useDebouncedCallback } from 'use-debounce';
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';

type SuggestItem = {
    id: number;
    type: 'class' | 'instructor' | 'department';
    title: string;
    subtitle?: string;
    href: string;
};

type SuggestResponse = {
    status: 'ok';
    data: {
        classes: SuggestItem[];
        instructors: SuggestItem[];
        departments: SuggestItem[];
    };
} | { status: 'error' };

export default function Search({ placeholder }: { placeholder: string }) {
    const searchParams = useSearchParams();
    const { push } = useRouter();
    const pathname = usePathname();
    const inputRef = useRef<HTMLInputElement>(null);
    const abortRef = useRef<AbortController | null>(null);
    const [open, setOpen] = useState(false);
    const [loading, setLoading] = useState(false);
    const [activeIndex, setActiveIndex] = useState<number>(-1);
    const [query, setQuery] = useState<string>(searchParams.get('query')?.toString() || '');
    const [groups, setGroups] = useState<{ key: string; label: string; items: SuggestItem[] }[]>([]);

    const handleSearch = useDebouncedCallback((term: string) => {
        console.log(`Searching... ${term}`);
        const params = new URLSearchParams(searchParams);
        params.set('page', '1');
        if (term) {
            params.set('query', term);
        } else {
            params.delete('query');
        }
        push(`${pathname}?${params.toString()}`);
    }, 300);

    const fetchSuggestions = useDebouncedCallback(async (term: string) => {
        if (!term) {
            setGroups([]);
            setOpen(false);
            return;
        }
        try {
            abortRef.current?.abort();
            const controller = new AbortController();
            abortRef.current = controller;
            setLoading(true);
            const res = await fetch(`/api/suggest?q=${encodeURIComponent(term)}&limit=5`, {
                signal: controller.signal,
            });
            if (!res.ok) throw new Error('Network error');
            const data: SuggestResponse = await res.json();
            if (data.status === 'ok') {
                const g = [
                    { key: 'classes', label: 'Classes', items: data.data.classes },
                    { key: 'instructors', label: 'Instructors', items: data.data.instructors },
                    { key: 'departments', label: 'Departments', items: data.data.departments },
                ].filter(gr => gr.items.length > 0);
                setGroups(g);
                setActiveIndex(g.length ? 0 : -1);
                setOpen(g.length > 0);
            } else {
                setGroups([]);
                setOpen(false);
            }
        } catch (e) {
            if ((e as any)?.name !== 'AbortError') {
                console.error('suggest error', e);
                setGroups([]);
                setOpen(false);
            }
        } finally {
            setLoading(false);
        }
    }, 200);

    const handleKeyDown = useCallback((event: React.KeyboardEvent<HTMLInputElement>) => {
        if (!open) {
            if (event.key === 'Enter') {
                const searchTerm = inputRef.current?.value || '';
                push(`/search?query=${encodeURIComponent(searchTerm)}`);
            }
            return;
        }
        const flat = groups.flatMap(g => g.items);
        if (event.key === 'ArrowDown') {
            event.preventDefault();
            setActiveIndex(i => (i + 1) % flat.length);
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            setActiveIndex(i => (i - 1 + flat.length) % flat.length);
        } else if (event.key === 'Enter') {
            event.preventDefault();
            const item = flat[activeIndex];
            if (item) {
                push(item.href);
                setOpen(false);
            } else {
                const searchTerm = inputRef.current?.value || '';
                push(`/search?query=${encodeURIComponent(searchTerm)}`);
                setOpen(false);
            }
        } else if (event.key === 'Escape') {
            setOpen(false);
        }
    }, [groups, activeIndex, open, push]);

    useEffect(() => {
        setQuery(searchParams.get('query')?.toString() || '');
    }, [searchParams]);

    const highlight = useCallback((text: string, q: string) => {
        const i = text.toLowerCase().indexOf(q.toLowerCase());
        if (i === -1) return <>{text}</>;
        const before = text.slice(0, i);
        const match = text.slice(i, i + q.length);
        const after = text.slice(i + q.length);
        return (<>
            {before}
            <span className="bg-red-800/20 dark:bg-red-400/20 font-semibold">{match}</span>
            {after}
        </>);
    }, []);


    return (
        <div className="relative flex flex-1 w-full">
            <label htmlFor="search" className="sr-only">
                Search
            </label>
            <input
                ref={inputRef}
                className="peer block w-full h-16 rounded-lg bg-[#f6f6ef] dark:bg-zinc-900 px-6 outline-none ring-red-900 dark:ring-red-400 transition focus:ring-2 focus:bg-[#f6f6ef] dark:focus:bg-zinc-900 text-black dark:text-white placeholder:text-gray-500 dark:placeholder:text-gray-400 py-[9px] pl-10 border border-black/10 dark:border-white/10"
                placeholder={placeholder}
                value={query}
                onChange={(e) => {
                    const val = e.target.value;
                    setQuery(val);
                    handleSearch(val);
                    fetchSuggestions(val);
                }}
                onKeyDown={handleKeyDown}
            />
            <MagnifyingGlassIcon className="absolute left-3 top-1/2 h-[18px] w-[18px] -translate-y-1/2 text-gray-500 dark:text-gray-400 peer-focus:text-gray-900 dark:peer-focus:text-gray-100" />

            {open && (
                <div className="absolute z-50 top-[110%] left-0 w-full rounded-lg border border-black/10 dark:border-white/10 bg-white/90 dark:bg-black/80 backdrop-blur p-2 shadow-xl">
                    {groups.map((group, gi) => (
                        <div key={group.key} className="mb-2 last:mb-0">
                            <div className="px-2 py-1 text-xs uppercase tracking-wide text-gray-600 dark:text-gray-400">{group.label}</div>
                            <ul role="listbox" className="max-h-80 overflow-auto">
                                {group.items.map((item, idx) => {
                                    const flatIndex = groups.slice(0, gi).reduce((acc, g) => acc + g.items.length, 0) + idx;
                                    const isActive = flatIndex === activeIndex;
                                    return (
                                        <li
                                            key={`${item.type}-${item.id}`}
                                            role="option"
                                            aria-selected={isActive}
                                            className={`px-3 py-2 rounded-md cursor-pointer transition ${isActive ? 'bg-red-800/15 dark:bg-red-400/20' : 'hover:bg-black/5 dark:hover:bg-white/10'}`}
                                            onMouseEnter={() => setActiveIndex(flatIndex)}
                                            onMouseDown={(e) => { e.preventDefault(); }}
                                            onClick={() => { setOpen(false); push(item.href); }}
                                        >
                                            <div className="text-sm text-gray-900 dark:text-gray-100">{highlight(item.title, query)}</div>
                                            {item.subtitle && (
                                                <div className="text-xs text-gray-600 dark:text-gray-400">{highlight(item.subtitle, query)}</div>
                                            )}
                                        </li>
                                    );
                                })}
                            </ul>
                        </div>
                    ))}
                    <div className="mt-1 pt-1 border-t border-black/10 dark:border-white/10 text-xs">
                        <button
                            className="w-full text-left px-2 py-2 text-gray-700 dark:text-gray-300 hover:underline"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => { setOpen(false); push(`/search?query=${encodeURIComponent(query)}`); }}
                        >
                            View all results for "{query}"
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
}