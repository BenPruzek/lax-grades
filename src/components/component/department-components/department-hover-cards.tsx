"use client";
import { HoverEffect } from "@/components/ui/card-hover-effect";

export default function DepartmentHoverCards({ departmentClasses, departmentInstructors }: { departmentClasses: any[], departmentInstructors: any[] }) {
    return (
        <div>
            <div className="px-6">
                <h3 className="text-2xl font-semibold text-gray-900 dark:text-gray-100 px-2">Classes</h3>
                <HoverEffect items={departmentClasses.map(classData => ({
                    title: classData.name,
                    description: classData.code,
                    link: `/class/${classData.code}`,
                }))} />
            </div>
            <div className="px-6 mt-8">
                <h3 className="text-2xl font-semibold text-gray-900 dark:text-gray-100 px-2">Instructors</h3>
                <HoverEffect items={departmentInstructors.map((instructor: { id: number; name: string; department: string }) => ({
                    title: instructor.name,
                    description: instructor.department,
                    link: `/instructor/${instructor.id}`,
                }))} />
            </div>

        </div>
    );
}