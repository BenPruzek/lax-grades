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

  // 1. Fetch reviews
  const reviews = await fetchClassReviews(classData.id);

  // 2. Calculate the Big Numbers
  // FIX: Convert both IDs to String to ensure "101" matches 101
  const activeReviews = instructor
    ? reviews.filter((r) => String(r.instructor.id) === String(instructor))
    : reviews;

  // FIX: Robust Filtering to ignore "Zombie" (Old) Reviews
  // We check if values exist AND are greater than 0
  const validQualityReviews = activeReviews.filter((r) => (r.clarity || 0) > 0 && (r.support || 0) > 0);
  const validIntensityReviews = activeReviews.filter((r) => (r.workload || 0) > 0 && (r.difficulty || 0) > 0);

  // Calculate Averages
  const avgClarity = validQualityReviews.reduce((sum, r) => sum + (r.clarity || 0), 0) / (validQualityReviews.length || 1);
  const avgSupport = validQualityReviews.reduce((sum, r) => sum + (r.support || 0), 0) / (validQualityReviews.length || 1);
  
  const avgWorkload = validIntensityReviews.reduce((sum, r) => sum + (r.workload || 0), 0) / (validIntensityReviews.length || 1);
  const avgDifficulty = validIntensityReviews.reduce((sum, r) => sum + (r.difficulty || 0), 0) / (validIntensityReviews.length || 1);

  // Create Big Numbers
  // If no valid reviews, explicit "N/A"
  const qualityScore = validQualityReviews.length > 0 ? ((avgClarity + avgSupport) / 2).toFixed(1) : "N/A";
  const intensityScore = validIntensityReviews.length > 0 ? ((avgWorkload + avgDifficulty) / 2).toFixed(1) : "N/A";

  // Debugging Log (Check your VS Code Terminal to see this!)
  console.log(`--- DEBUG SCORES for ${instructor ? "Instructor " + instructor : "All"} ---`);
  console.log(`Total Reviews: ${activeReviews.length}`);
  console.log(`Valid Quality Reviews: ${validQualityReviews.length}`);
  console.log(`Quality Score: ${qualityScore}`);

  // 3. Logic for the Review Section
  const allInstructors = distributions
    .map((d) => d.instructor)
    .filter((i): i is NonNullable<typeof i> => i !== null && i !== undefined);

  const uniqueInstructors = Array.from(
    new Map(allInstructors.map((item) => [item.id, item])).values()
  );

  const availableInstructors = uniqueInstructors.map((i) => ({
    id: i.id,
    name: i.name,
  })).sort((a, b) => a.name.localeCompare(b.name));

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
          <ClassFilterSelect 
             classData={classData} 
             distributions={distributions} 
             qualityScore={qualityScore}
             intensityScore={intensityScore}
          />
        </div>
        
        <div className="mt-12">
          <ReviewsSection
            classId={classData.id}
            instructorId={instructorId}
            departmentId={classData.department.id}
            classCode={classData.code}
            initialReviews={reviews}
            availableInstructors={availableInstructors}
          />
        </div>
      </div>
    </>
  );
}