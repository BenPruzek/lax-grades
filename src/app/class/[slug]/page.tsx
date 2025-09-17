import ClassFilterSelect from "@/components/component/class-components/class-filter-select";
import Search from "@/components/component/search-components/search";
import { fetchGPADistributions, getClassByCode } from "@/lib/data";
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
      </div>
    </>
  );
}