import InstructorBarChart from "@/components/component/instructor-components/instructor-bar-chart";
import InstructorHoverCards from "@/components/component/instructor-components/instructor-hover-cards";
import InstructorDataCards from "@/components/component/instructor-components/intructor-data-cards";
import Search from "@/components/component/search-components/search";
// 1. Add getInstructorAggregates to imports
import { getInstructorById, fetchInstructorClasses, getInstructorAggregates } from "@/lib/data";
import { gradesOrder } from "@/lib/utils";
// 2. Add AnimatedScore
import { AnimatedScore } from "@/components/ui/animated-score";

export default async function InstructorPage({ params }: { params: { id: string } }) {
    const instructorId = parseInt(params.id);
    const instructor = await getInstructorById(instructorId);
    
    if (!instructor) return <div>Instructor not found</div>;

    const instructorData = await fetchInstructorClasses(instructorId);

    // 3. Fetch the Efficient Aggregates
    const aggregates = await getInstructorAggregates(instructorId);

    // 4. Calculate the Big Numbers
    const qualityScore = aggregates.count > 0 
        ? ((aggregates.avgClarity + aggregates.avgSupport) / 2).toFixed(1) 
        : "N/A";
        
    const intensityScore = aggregates.count > 0 
        ? ((aggregates.avgWorkload + aggregates.avgDifficulty) / 2).toFixed(1) 
        : "N/A";

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
                    <h1 className="text-4xl font-bold text-gray-900 dark:text-gray-100">{instructor.name}</h1>
                    <p className="text-xl text-gray-600 dark:text-gray-300">Instructor</p>
                </div>
                
                <div className="mt-6 flex flex-col gap-4">
                    <div className="col-span-2">
                        <h2 className="text-lg font-semibold text-gray-900 dark:text-gray-100 mb-4">Overall Grades Given</h2>
                        <div className="w-full h-[300px]">
                            <InstructorBarChart data={chartData} />
                        </div>
                    </div>

                    {/* --- 5. NEW: INSTRUCTOR SCORE CARDS --- */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        {/* Quality Card */}
                        <div className="p-4 bg-[#f6f6ef] dark:bg-zinc-900 border border-transparent rounded-lg">
                            <div className="flex justify-between items-start">
                                <div>
                                    <h3 className="text-sm font-semibold text-gray-700 dark:text-gray-200">Instructor Quality</h3>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">Clarity & Support</p>
                                </div>
                                <div className="text-right">
                                    <AnimatedScore value={qualityScore} />
                                    <span className="text-xs text-gray-400 ml-1">/5</span>
                                </div>
                            </div>
                        </div>

                        {/* Intensity Card */}
                        <div className="p-4 bg-[#f6f6ef] dark:bg-zinc-900 border border-transparent rounded-lg">
                            <div className="flex justify-between items-start">
                                <div>
                                    <h3 className="text-sm font-semibold text-gray-700 dark:text-gray-200">Course Intensity</h3>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">Avg. Difficulty & Workload</p>
                                </div>
                                <div className="text-right">
                                    <AnimatedScore value={intensityScore} />
                                    <span className="text-xs text-gray-400 ml-1">/5</span>
                                </div>
                            </div>
                        </div>
                    </div>
                    {/* ------------------------------------- */}

                    <InstructorDataCards totalStudents={totalStudents} averageGPA={averageGPA} percentageA={percentageA} />
                </div>
            </div>
            <InstructorHoverCards instructorClasses={instructorData} />
        </div >
    );
}
