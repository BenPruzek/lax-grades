"use client";

import { useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Wallet } from "lucide-react";

export default function Error({
    error,
    reset,
}: {
    error: Error & { digest?: string };
    reset: () => void;
}) {
    useEffect(() => {
        console.error(error);
    }, [error]);

    return (
        <main className="flex min-h-screen flex-col items-center justify-center bg-gray-50 p-4 dark:bg-gray-900">
            <Card className="w-full max-w-md text-center border-2 shadow-lg">
                <CardHeader className="flex flex-col items-center gap-4 pb-2">
                    <div className="rounded-full bg-red-100 p-3 dark:bg-red-900/20">
                        <Wallet className="h-10 w-10 text-red-600 dark:text-red-400" />
                    </div>
                    <CardTitle className="text-xl sm:text-2xl">
                        Sorry, I ran out of money to pay for this.
                    </CardTitle>
                </CardHeader>
                <CardContent>
                    <p className="text-muted-foreground text-lg">
                        Contact your school to consider supporting this service.
                    </p>
                </CardContent>
                <CardFooter className="flex justify-center pb-8">
                    <Button onClick={reset} size="lg" className="font-semibold">
                        Try again
                    </Button>
                </CardFooter>
            </Card>
        </main>
    );
}
