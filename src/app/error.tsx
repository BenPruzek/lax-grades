"use client";

import { useEffect } from "react";

export default function Error({
    error,
    reset,
}: {
    error: Error & { digest?: string };
    reset: () => void;
}) {
    useEffect(() => {
        // Log the error to an error reporting service
        console.error(error);
    }, [error]);

    return (
        <main className="text-center py-24 h-screen bg-white dark:bg-transparent">
            {/* <H1 className="mb-5">Something went wrong!</H1> */}
            <h1>{error.message}</h1>
            <button
                onClick={
                    // Attempt to recover by trying to re-render the segment
                    reset
                }
            >
                Try again
            </button>
        </main>
    );
}
