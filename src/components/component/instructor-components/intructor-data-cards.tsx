export default function InstructorDataCards({ totalStudents, averageGPA, percentageA }: { totalStudents: number, averageGPA: number, percentageA: number }) {
    return (
        <div className="grid gap-4 sm:grid-cols-1 lg:grid-cols-3" >
            <div className="bg-[#f6f6ef] dark:bg-zinc-900 p-4 rounded-lg">
                <h3 className="text-base sm:text-lg font-semibold text-gray-900 dark:text-gray-100">Students taught</h3>
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
            {/*<Link href={instructor.RMP_link} className="bg-white p-4 rounded-lg">
            <h3 className="text-lg font-semibold text-gray-900">Rate My Professor</h3>
            <p className="text-3xl font-bold text-red-800">{instructor.RMP_score}</p>
        </Link>
        <Link href={instructor.RMP_link} className="bg-white p-4 rounded-lg">
            <h3 className="text-lg font-semibold text-gray-900">Difficulty</h3>
            <p className="text-3xl font-bold text-red-800">{instructor.RMP_diff}</p>
        </Link>*/}
        </div>
    );
}