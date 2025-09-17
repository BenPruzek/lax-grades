"use client";
import { useRouter } from "next/navigation";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../ui/select";
import { useEffect, useState } from "react";
import { getUniqueInstructors, getUniqueSemesters, gradesOrder } from "@/lib/utils";
import ClassBarChart from "./class-bar-chart";
import Link from "next/link";
import ClassDataCards from "./class-data-cards";
import { GradePercentages } from "@/lib/types";

export default function ClassFilterSelect({ classData, distributions }: { classData: any; distributions: any[] }) {
    const router = useRouter();
    const [selectedInstructor, setSelectedInstructor] = useState<number | null>(null);
    const [selectedSemester, setSelectedSemester] = useState<string | null>(null);

    useEffect(() => {
        const params = new URLSearchParams();
        if (selectedInstructor) params.set('instructor', selectedInstructor.toString());
        if (selectedSemester) params.set('semester', selectedSemester);
        const url = `/class/${classData.code}?${params.toString()}`;
        router.push(url, { scroll: false });
    }, [selectedInstructor, selectedSemester, classData.code, router]);

    const instructors = getUniqueInstructors(distributions);
    const semesters = selectedInstructor
        ? getUniqueSemesters(distributions.filter(dist => dist.instructor?.id === selectedInstructor))
        : getUniqueSemesters(distributions);

    const filteredDistributions = distributions.filter((dist) => {
        const instructorMatch =
            selectedInstructor === null || dist.instructor?.id === selectedInstructor;
        const semesterMatch = selectedSemester === null || dist.term === selectedSemester;
        return instructorMatch && semesterMatch;
    });

    const filteredNumStudents = filteredDistributions.reduce((acc, curr) => acc + curr.studentHeadcount, 0);
    const filteredAverageGPA = filteredDistributions.reduce((acc, curr) => acc + curr.avgCourseGrade * curr.studentHeadcount, 0) / filteredNumStudents;
    const filteredGradePercentages: GradePercentages = gradesOrder.reduce((acc, grade) => {
        const totalGradeStudents = filteredDistributions.reduce((sum, dist) => sum + (dist.grades[grade] / 100 * dist.studentHeadcount), 0);
        acc[grade] = (totalGradeStudents / filteredNumStudents) * 100;
        return acc;
    }, {} as GradePercentages);

    const cumulativeNumStudents = distributions.reduce((sum, dist) => sum + dist.studentHeadcount, 0)
    const cumulativeAverageGPA = distributions.reduce((acc, curr) => acc + curr.avgCourseGrade * curr.studentHeadcount, 0) / cumulativeNumStudents;
    const cumulativeGradePercentages: GradePercentages = gradesOrder.reduce((acc, grade) => {
        const totalGradeStudents = distributions.reduce((sum, dist) => sum + (dist.grades[grade] / 100 * dist.studentHeadcount), 0);
        acc[grade] = (totalGradeStudents / cumulativeNumStudents) * 100;
        return acc;
    }, {} as GradePercentages);

    let chartData = gradesOrder.map(grade => {
        const data: any = {
            grade,
            cumulative: cumulativeGradePercentages[grade],
        };
        if (selectedInstructor) {
            data.instructor = filteredGradePercentages[grade];
        } else if (selectedSemester) {
            data.semester = filteredGradePercentages[grade];
        }
        return data;
    }).filter(entry => entry.cumulative > 0 || entry.instructor > 0 || entry.semester > 0);

    return (
        <>
            <div className="lg:col-span-1">
                <div className="mb-8">
                    <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1" htmlFor="instructors">
                        Instructors
                    </label>
                    <Select
                        value={selectedInstructor !== null ? selectedInstructor.toString() : undefined}
                        onValueChange={(value) => setSelectedInstructor(value ? Number(value) : null)}
                    >
                        <SelectTrigger id="instructors" className="w-full bg-[#f6f6ef] border-[#f6f6ef] text-black dark:bg-zinc-900 dark:border-zinc-900 dark:text-white">
                            <SelectValue placeholder="All Instructors" />
                        </SelectTrigger>
                        <SelectContent position="popper" className='bg-[#f6f6ef] border-[#f6f6ef] dark:bg-zinc-900 dark:border-zinc-900'>
                            {/** @ts-ignore */}
                            <SelectItem value={null}>All Instructors</SelectItem>
                            {instructors.map((instructor) => (
                                <SelectItem key={instructor.id} value={instructor.id.toString()}>
                                    {instructor.name}
                                </SelectItem>
                            ))}
                        </SelectContent>
                    </Select>
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1" htmlFor="semesters">
                        Semesters
                    </label>
                    <Select
                        value={selectedSemester !== null ? selectedSemester : undefined}
                        onValueChange={(value) => setSelectedSemester(value)}
                    >
                        <SelectTrigger id="semesters" className="w-full bg-[#f6f6ef] border-[#f6f6ef] text-black dark:bg-zinc-900 dark:border-zinc-900 dark:text-white">
                            <SelectValue placeholder="All Semesters" />
                        </SelectTrigger>
                        <SelectContent position="popper" className='bg-[#f6f6ef] border-[#f6f6ef] dark:bg-zinc-900 dark:border-zinc-900'>
                            {/** @ts-ignore */}
                            <SelectItem value={null}>All Semesters</SelectItem>
                            {semesters.map(semester => (
                                <SelectItem key={semester} value={semester}>
                                    {semester}
                                </SelectItem>
                            ))}
                        </SelectContent>
                    </Select>
                </div>
            </div>
            <div className="mt-6 lg:mt-0 lg:col-span-3">
                <div className="col-span-2">
                    <div className="flex items-center mb-4">
                        <div className="text-lg font-semibold text-gray-900 dark:text-gray-100">
                            {classData.name}: {selectedInstructor === null && selectedSemester === null && "Cumulative"} {selectedInstructor !== null && (
                                <Link href={`/instructor/${selectedInstructor}`} className="text-red-800 dark:text-red-400 hover:underline">
                                    {instructors.find((instructor) => instructor.id === selectedInstructor)?.name}
                                </Link>
                            )}
                        </div>
                        {selectedSemester !== null && <span className="ml-2 text-gray-600 dark:text-gray-300">({selectedSemester})</span>}
                    </div>
                    <ClassBarChart className="w-full h-[500px]"
                        data={chartData}
                        selectedInstructor={selectedInstructor !== null ? instructors.find((instructor) => instructor.id === selectedInstructor)?.name || null : null}
                        selectedSemester={selectedSemester}
                    />
                </div>
                <ClassDataCards
                    cumulativeNumStudents={cumulativeNumStudents}
                    cumulativeAverageGPA={cumulativeAverageGPA}
                    cumulativeGradePercentages={cumulativeGradePercentages}
                    filteredNumStudents={filteredNumStudents}
                    filteredAverageGPA={filteredAverageGPA}
                    filteredGradePercentages={filteredGradePercentages}
                    isFiltered={selectedInstructor !== null || selectedSemester !== null}
                />
            </div>
        </>
    );
}