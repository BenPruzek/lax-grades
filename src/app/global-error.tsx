"use client";

export default function GlobalError({
    error,
    reset,
}: {
    error: Error & { digest?: string };
    reset: () => void;
}) {
    return (
        <html>
            <body>
                <main className="text-center py-24 h-screen bg-white dark:bg-transparent flex flex-col items-center justify-center px-4">
                    <h1 className="text-2xl md:text-3xl font-bold mb-4 text-gray-900 dark:text-gray-100">
                        Sorry, I ran out of money to pay for this.
                    </h1>
                    <p className="text-xl opacity-75 mb-8 text-gray-900 dark:text-gray-100">
                        Contact your school to consider supporting this service.
                    </p>
                    <button
                        onClick={() => reset()}
                        className="px-4 py-2 bg-gray-900 text-white rounded hover:bg-gray-800 dark:bg-gray-100 dark:text-gray-900 dark:hover:bg-gray-200 transition-colors"
                    >
                        Try again
                    </button>
                </main>
            </body>
        </html>
    );
}
