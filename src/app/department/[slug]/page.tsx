import DepartmentBarChart from "@/components/component/department-components/department-bar-chart";
import DepartmentDataCards from "@/components/component/department-components/department-data-cards";
import DepartmentHoverCards from "@/components/component/department-components/department-hover-cards";
import Search from "@/components/component/search-components/search";
// 1. Add getDepartmentAggregates to imports
import { getDepartmentByCode, fetchDepartmentClasses, fetchDepartmentInstructors, fetchDepartmentGrades, getDepartmentAggregates } from "@/lib/data";
import { gradesOrder } from "@/lib/utils";
// 2. Add AnimatedScore import
import { AnimatedScore } from "@/components/ui/animated-score";

export default async function DepartmentPage({ params }: { params: { slug: string } }) {
    const departmentCode = params.slug;
    const department = await getDepartmentByCode(departmentCode);
    
    // Handle case where department doesn't exist (optional safety)
    if (!department) return <div>Department not found</div>;

    const departmentClasses = await fetchDepartmentClasses(department.id);
    const departmentGrades = await fetchDepartmentGrades(department.id);
    const departmentInstructors = await fetchDepartmentInstructors(department.name);

    // 3. Fetch the Efficient Aggregates
    const aggregates = await getDepartmentAggregates(department.id);

    // 4. Calculate the Big Numbers
    const qualityScore = aggregates.count > 0 
        ? ((aggregates.avgClarity + aggregates.avgSupport) / 2).toFixed(1) 
        : "N/A";
        
    const intensityScore = aggregates.count > 0 
        ? ((aggregates.avgWorkload + aggregates.avgDifficulty) / 2).toFixed(1) 
        : "N/A";

    const totalStudents = departmentGrades.reduce((acc, curr) => acc + curr.studentHeadcount, 0);
    const averageGPA = departmentGrades.reduce((acc, curr) => acc + curr.avgCourseGrade * curr.studentHeadcount, 0) / totalStudents;
    const gradePercentages: { [key: string]: number } = {};

    gradesOrder.forEach(grade => {
        const totalGradeStudents = departmentGrades.reduce((acc, curr) => acc + (curr.gradePercentages[grade] / 100 * curr.studentHeadcount), 0);
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
                    <h1 className="text-4xl font-bold text-gray-900 dark:text-gray-100">{department.name}</h1>
                    <p className="text-xl text-gray-600 dark:text-gray-300">{department.code}</p>
                </div>
                
                <div className="mt-6 flex flex-col gap-4">
                    <div className="col-span-2">
                        <h2 className="text-lg font-semibold text-gray-900 dark:text-gray-100 mb-4">Overall Grades in Department</h2>
                        <div className="w-full h-[300px]">
                            <DepartmentBarChart data={chartData} />
                        </div>
                    </div>

                    {/* --- 5. NEW: DEPARTMENT SCORE CARDS --- */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        {/* Quality Card */}
                        <div className="p-4 bg-[#f6f6ef] dark:bg-zinc-900 border border-transparent rounded-lg">
                            <div className="flex justify-between items-start">
                                <div>
                                    <h3 className="text-sm font-semibold text-gray-700 dark:text-gray-200">Avg. Instructor Quality</h3>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">Based on all classes in department</p>
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
                                    <h3 className="text-sm font-semibold text-gray-700 dark:text-gray-200">Avg. Course Intensity</h3>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">Difficulty & Workload</p>
                                </div>
                                <div className="text-right">
                                    <AnimatedScore value={intensityScore} />
                                    <span className="text-xs text-gray-400 ml-1">/5</span>
                                </div>
                            </div>
                        </div>
                    </div>
                    {/* ------------------------------------- */}

                    <DepartmentDataCards totalStudents={totalStudents} averageGPA={averageGPA} percentageA={percentageA} />
                </div>
            </div>
            <DepartmentHoverCards departmentInstructors={departmentInstructors} departmentClasses={departmentClasses} />
        </div>
    );
}
