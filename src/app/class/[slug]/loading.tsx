import { Skeleton } from "@/components/ui/skeleton";

export default function Loading() {
    return (
        <div className="bg-white dark:bg-transparent p-8">
            <Skeleton className="bg-gray-200 h-12 w-full mb-4" />
            <div className="border-b border-red-800 pb-4 pt-6">
                <Skeleton className="bg-gray-200 h-12 w-1/2 mb-2" />
                <Skeleton className="bg-gray-200 h-8 w-1/3" />
            </div>
            <div className="lg:grid lg:grid-cols-4 gap-16 mt-4">
                <div className="lg:col-span-1">
                    <div className="mb-8">
                        <Skeleton className="bg-gray-200 h-6 w-1/2 mb-1" />
                        <Skeleton className="bg-gray-200 h-12 w-full" />
                    </div>
                    <div>
                        <Skeleton className="bg-gray-200 h-6 w-1/2 mb-1" />
                        <Skeleton className="bg-gray-200 h-12 w-full" />
                    </div>
                </div>
                <div className="mt-6 lg:mt-0 lg:col-span-3">
                    <div className="col-span-2">
                        <div className="flex items-center mb-4">
                            <Skeleton className="bg-gray-200 h-6 w-1/2" />
                            <Skeleton className="bg-gray-200 h-6 w-1/4 ml-2" />
                        </div>
                        <Skeleton className="bg-gray-200 h-[500px] w-full" />
                    </div>
                    <div className="grid gap-4 sm:grid-cols-1 lg:grid-cols-3 mt-8">
                        <Skeleton className="bg-gray-200 h-28 rounded-lg" />
                        <Skeleton className="bg-gray-200 h-28 rounded-lg" />
                        <Skeleton className="bg-gray-200 h-28 rounded-lg" />
                    </div>
                </div>
            </div>
        </div>
    );
}