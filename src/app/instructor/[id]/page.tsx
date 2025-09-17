import InstructorBarChart from "@/components/component/instructor-components/instructor-bar-chart";
import InstructorHoverCards from "@/components/component/instructor-components/instructor-hover-cards";
import InstructorDataCards from "@/components/component/instructor-components/intructor-data-cards";
import Search from "@/components/component/search-components/search";
import { getInstructorById, fetchInstructorClasses } from "@/lib/data";
import { gradesOrder } from "@/lib/utils";

export default async function InstructorPage({ params }: { params: { id: string } }) {
    const instructorId = parseInt(params.id);
    const instructor = await getInstructorById(instructorId);
    const instructorData = await fetchInstructorClasses(instructorId);

    const totalStudents = instructorData.reduce((acc, curr) => acc + curr.studentHeadcount, 0);
    const averageGPA = instructorData.reduce((acc, curr) => acc + curr.avgCourseGrade * curr.studentHeadcount, 0) / totalStudents;
    const gradePercentages: { [key: string]: number } = {};
    gradesOrder.forEach(grade => {
        const totalGradeStudents = instructorData.reduce((acc, curr) => acc + (curr.gradePercentages[grade] / 100 * curr.studentHeadcount), 0);
        gradePercentages[grade] = (totalGradeStudents / totalStudents) * 100;
    });

    const chartData = gradesOrder.map(grade => ({
        grade,
        percentage: gradePercentages[grade],
    })).filter(entry => entry.percentage > 0);

    const percentageA = gradePercentages["A"];

    return (
        <div className="bg-white dark:bg-transparent">
            <div className="p-8">
                <Search placeholder="Search for classes, instructors, or departments" />
                <div className="border-b border-red-800 pb-4 pt-6">
                    <h1 className="text-4xl font-bold text-gray-900">{instructor!.name}</h1>
                    <p className="text-xl text-gray-600">Instructor</p>
                </div>
                <div className="mt-6 flex flex-col gap-4">
                    <div className="col-span-2">
                        <h2 className="text-lg font-semibold text-gray-900 mb-4">Overall Grades Given</h2>
                        <div className="w-full h-[300px]">
                            <InstructorBarChart data={chartData} />
                        </div>
                    </div>
                    <InstructorDataCards totalStudents={totalStudents} averageGPA={averageGPA} percentageA={percentageA} />
                </div>
            </div>
            <InstructorHoverCards instructorClasses={instructorData} />
        </div >
    );
}