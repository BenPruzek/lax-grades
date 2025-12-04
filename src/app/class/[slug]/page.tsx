import ClassFilterSelect from "@/components/component/class-components/class-filter-select";
import Search from "@/components/component/search-components/search";
import ReviewsSection from "@/components/component/reviews/reviews-section";
import { fetchGPADistributions, getClassByCode, fetchClassReviews } from "@/lib/data";
import { ClassData } from "@/lib/types";
import Link from "next/link";

export default async function ClassPage({ params, searchParams }: {
  params: { slug: string };
  searchParams?: { instructor?: string; semester?: string };
}) {
  const slug = decodeURIComponent(params.slug);
  const instructor = searchParams?.instructor;
  const semester = searchParams?.semester;

  const classData: ClassData | null = await getClassByCode(slug);

  if (!classData) {
    return <div className="bg-white dark:bg-transparent h-screen flex flex-center justify-center text-h1">NO DATA</div>;
  }

  const distributions = await fetchGPADistributions(
    classData.id,
    classData.department.id,
  );

  // Fetch reviews for this class
  const reviews = await fetchClassReviews(classData.id);

  // --- NEW LOGIC START ---
  // 1. Extract all instructors from the distribution data
  const allInstructors = distributions
    .map((d) => d.instructor)
    .filter((i): i is NonNullable<typeof i> => i !== null && i !== undefined);

  // 2. Remove duplicates (instructors teach multiple semesters, so they appear multiple times)
  const uniqueInstructors = Array.from(
    new Map(allInstructors.map((item) => [item.id, item])).values()
  );

  // 3. Format for the dropdown
  const availableInstructors = uniqueInstructors.map((i) => ({
    id: i.id,
    name: i.name,
  })).sort((a, b) => a.name.localeCompare(b.name)); // Sort alphabetically
  // --- NEW LOGIC END ---

  // Default instructor ID (keep existing logic for fallback)
  const instructorId = distributions.find((distribution) => distribution.instructor)?.instructor?.id ?? 0;

  return (
    <>
      <div className="bg-white dark:bg-transparent p-8">
        <Search placeholder="Search for classes, instructors, or departments" />
        <div className="border-b border-red-800 pb-4 pt-6">
          <h1 className="text-4xl font-bold text-gray-900 dark:text-gray-100">{classData.name}</h1>
          <p className="text-xl text-gray-600 dark:text-gray-300">
            <Link href={`/department/${classData.code.slice(0, classData.code.search(/\d/))}`}>
              <span className="text-red-800 dark:text-red-400 hover:underline">{classData.code.slice(0, classData.code.search(/\d/))}</span>
            </Link>
            {classData.code.slice(classData.code.search(/\d/))}
          </p>
        </div>
        <div className="lg:grid lg:grid-cols-4 gap-16 mt-4">
          <ClassFilterSelect classData={classData} distributions={distributions} />
        </div>
        
        {/* Reviews Section */}
        <div className="mt-12">
          <ReviewsSection
            classId={classData.id}
            instructorId={instructorId}
            departmentId={classData.department.id}
            classCode={classData.code}
            initialReviews={reviews}
            // PASS THE LIST HERE
            availableInstructors={availableInstructors}
          />
        </div>
      </div>
    </>
  );
}
