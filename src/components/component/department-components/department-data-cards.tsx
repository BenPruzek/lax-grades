
export default function DepartmentDataCards({ totalStudents, averageGPA, percentageA }: { totalStudents: number, averageGPA: number, percentageA: number }) {
    return (
        <div className="grid gap-4 sm:grid-cols-1 lg:grid-cols-3">
            <div className="bg-[#f6f6ef] dark:bg-zinc-900 p-4 rounded-lg">
                <h3 className="text-base sm:text-lg font-semibold text-gray-900 dark:text-gray-100">Total Students</h3>
                <p className="text-2xl sm:text-3xl font-bold text-red-800 dark:text-red-400">{totalStudents.toFixed(0)}</p>
            </div>
            <div className="bg-[#f6f6ef] dark:bg-zinc-900 p-4 rounded-lg">
                <h3 className="text-base sm:text-lg font-semibold text-gray-900 dark:text-gray-100">Average GPA</h3>
                <p className="text-2xl sm:text-3xl font-bold text-red-800 dark:text-red-400">{averageGPA.toFixed(2)}</p>
            </div>
            <div className="bg-[#f6f6ef] dark:bg-zinc-900 p-4 rounded-lg">
                <h3 className="text-base sm:text-lg font-semibold text-gray-900 dark:text-gray-100">Percentage A</h3>
                <p className="text-2xl sm:text-3xl font-bold text-red-800 dark:text-red-400">{percentageA.toFixed(1)}%</p>
            </div>
        </div>
    );
}
