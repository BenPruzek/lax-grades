import { Skeleton } from "@/components/ui/skeleton";

export default function Loading() {
    return (
        <div className="bg-white dark:bg-transparent">
            <div className="p-8">
                <Skeleton className="bg-gray-200 h-12 w-full mb-4" />
                <div className="border-b border-red-800 pb-4 pt-6">
                    <Skeleton className="bg-gray-200 h-12 w-1/2 mb-2" />
                    <Skeleton className="bg-gray-200 h-8 w-1/3" />
                </div>
                <div className="mt-6 flex flex-col gap-4">
                    <div className="col-span-2">
                        <Skeleton className="bg-gray-200 h-6 w-1/3 mb-4" />
                        <Skeleton className="bg-gray-200 h-[300px] w-full" />
                    </div>
                    <div className="grid gap-4 sm:grid-cols-1 lg:grid-cols-3">
                        <Skeleton className="bg-gray-200 h-20 rounded-lg" />
                        <Skeleton className="bg-gray-200 h-20 rounded-lg" />
                        <Skeleton className="bg-gray-200 h-20 rounded-lg" />
                    </div>
                </div>
            </div>
            <div className="px-6">
                <Skeleton className="bg-gray-200 h-8 w-1/4 mb-4" />
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    <Skeleton className="bg-gray-200 h-40 rounded-lg" />
                    <Skeleton className="bg-gray-200 h-40 rounded-lg" />
                    <Skeleton className="bg-gray-200 h-40 rounded-lg" />
                </div>
            </div>
            <div className="px-6 mt-8">
                <Skeleton className="bg-gray-200 h-8 w-1/4 mb-4" />
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    <Skeleton className="bg-gray-200 h-40 rounded-lg" />
                    <Skeleton className="bg-gray-200 h-40 rounded-lg" />
                    <Skeleton className="bg-gray-200 h-40 rounded-lg" />
                </div>
            </div>
        </div>
    );
}