'use client';
import { MagnifyingGlassIcon } from '@heroicons/react/24/outline';
import { usePathname, useRouter, useSearchParams } from 'next/navigation';
import { useDebouncedCallback } from 'use-debounce';
import { useCallback, useRef } from 'react';

export default function Search({ placeholder }: { placeholder: string }) {
    const searchParams = useSearchParams();
    const { push } = useRouter();
    const pathname = usePathname();
    const inputRef = useRef<HTMLInputElement>(null);

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

    const handleKeyPress = useCallback((event: React.KeyboardEvent<HTMLInputElement>) => {
        if (event.key === 'Enter') {
            const searchTerm = inputRef.current?.value || '';
            const params = new URLSearchParams();
            params.set('query', searchTerm);
            push(`/search?${params.toString()}`);
        }
    }, [push]);


    return (
        <div className="relative flex flex-1 w-full">
            <label htmlFor="search" className="sr-only">
                Search
            </label>
            <input
                ref={inputRef}
                className="peer block w-full h-16 rounded-lg bg-[#f6f6ef] dark:bg-zinc-900 px-6 outline-none ring-red-900 dark:ring-red-400 transition focus:ring-2 focus:bg-[#f6f6ef] dark:focus:bg-zinc-900 text-black dark:text-white placeholder:text-gray-500 dark:placeholder:text-gray-400 py-[9px] pl-10 border border-black/10 dark:border-white/10"
                placeholder={placeholder}
                onChange={(e) => {
                    handleSearch(e.target.value);
                }}
                onKeyPress={handleKeyPress}
                defaultValue={searchParams.get('query')?.toString()}
            />
            <MagnifyingGlassIcon className="absolute left-3 top-1/2 h-[18px] w-[18px] -translate-y-1/2 text-gray-500 dark:text-gray-400 peer-focus:text-gray-900 dark:peer-focus:text-gray-100" />
        </div>
    );
}