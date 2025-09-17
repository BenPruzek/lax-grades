import { Suspense } from "react";
import { Skeleton } from "@/components/ui/skeleton";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import Search from "@/components/component/search-components/search";
import { getSearch } from "@/lib/data";
import SearchList from "@/components/component/search-components/search-list";

export default async function Page({
    searchParams,
}: {
    searchParams?: {
        query?: string;
        page?: string;
    };
}) {
    const query = searchParams?.query || "";
    const page = Number(searchParams?.page) || 1;
    const limit = 10;

    const { classes, instructors, departments, classCount, instructorCount, departmentCount } = await getSearch(
        query,
        page,
        page,
        page,
        limit
    );

    const totalPages = {
        classes: Math.ceil(classCount / limit),
        instructors: Math.ceil(instructorCount / limit),
        departments: Math.ceil(departmentCount / limit),
    };

    return (
        <div className="flex flex-col items-center justify-start p-8 gap-8 bg-white dark:bg-transparent min-h-[85vh]">
            <div className="flex flex-col items-center w-full">
                <Search placeholder="Search for classes, instructors, or departments" />
            </div>
            <Tabs defaultValue="Class" className="w-full">
                <TabsList className="flex justify-center items-center mb-4">
                    <TabsTrigger value="Class">Class</TabsTrigger>
                    <TabsTrigger value="Instructor">Instructor</TabsTrigger>
                    <TabsTrigger value="Department">Department</TabsTrigger>
                </TabsList>
                <TabsContent value="Class">
                    <Suspense fallback={<Skeleton />}>
                        <SearchList
                            results={classes}
                            currentPage={page}
                            totalPages={totalPages.classes}
                            entity="classes"
                        />
                    </Suspense>
                </TabsContent>
                <TabsContent value="Instructor">
                    <Suspense fallback={<Skeleton />}>
                        <SearchList
                            results={instructors}
                            currentPage={page}
                            totalPages={totalPages.instructors}
                            entity="instructors"
                        />
                    </Suspense>
                </TabsContent>
                <TabsContent value="Department">
                    <Suspense fallback={<Skeleton />}>
                        <SearchList
                            results={departments}
                            currentPage={page}
                            totalPages={totalPages.departments}
                            entity="departments"
                        />
                    </Suspense>
                </TabsContent>
            </Tabs>
        </div>
    );
}