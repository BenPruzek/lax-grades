"use client";
import { HoverEffect } from "@/components/ui/card-hover-effect";

export default function InstructorHoverCards({ instructorClasses }: { instructorClasses: any[] }) {
    return (
        <div className="px-6">
            <h3 className="text-2xl font-semibold text-gray-900 dark:text-gray-100 px-2">Classes Taught</h3>
            <HoverEffect items={instructorClasses.map(classData => ({
                title: classData.class.name,
                description: classData.class.code,
                link: `/class/${classData.class.code}`,
            }))} />
        </div>
    );
}