import { prisma } from '@/lib/prisma';
import type { Prisma } from '@prisma/client';
import type { Review as ReviewType } from './types';

const reviewInclude = {
    user: {
        select: {
            id: true,
            name: true,
            email: true,
        },
    },
    class: {
        select: {
            id: true,
            code: true,
            name: true,
        },
    },
    instructor: {
        select: {
            id: true,
            name: true,
        },
    },
    department: {
        select: {
            id: true,
            code: true,
            name: true,
        },
    },
} satisfies Prisma.ReviewInclude;

type ReviewWithRelations = Prisma.ReviewGetPayload<{ include: typeof reviewInclude }>;

function serializeReview(review: ReviewWithRelations): ReviewType {
    return {
        id: review.id,
        title: review.title,
        rating: review.rating,
        content: review.content,
        courseCode: review.courseCode,
        isOnlineCourse: review.isOnlineCourse,
        difficulty: review.difficulty ?? null,
        
        // --- ADDED NEW METRICS HERE ---
        // Note: You might get a red squiggly line here if you haven't updated 'types.ts' yet.
        // We will fix that in the next step.
        clarity: review.clarity,
        workload: review.workload,
        support: review.support,
        // ------------------------------

        wouldTakeAgain: review.wouldTakeAgain ?? null,
        attendanceMandatory: review.attendanceMandatory ?? null,
        grade: review.grade ?? null,
        tags: review.tags as ReviewType['tags'],
        createdAt: review.createdAt.toISOString(),
        updatedAt: review.updatedAt.toISOString(),
        user: {
            id: review.user.id,
            name: review.user.name,
            email: review.user.email,
        },
        class: {
            id: review.class.id,
            code: review.class.code,
            name: review.class.name,
        },
        instructor: {
            id: review.instructor.id,
            name: review.instructor.name,
        },
        department: {
            id: review.department.id,
            code: review.department.code,
            name: review.department.name,
        },
    };
}

export async function fetchClassReviews(classId: number) {
    try {
        const reviews = await prisma.review.findMany({
            where: { classId },
            include: reviewInclude,
            orderBy: { createdAt: 'desc' },
        });
        return reviews.map(serializeReview);
    } catch (error) {
        console.error('Failed to fetch class reviews:', error);
        throw new Error('Failed to fetch class reviews');
    }
}

export async function fetchInstructorReviews(instructorId: number) {
    try {
        const reviews = await prisma.review.findMany({
            where: { instructorId },
            include: reviewInclude,
            orderBy: { createdAt: 'desc' },
        });
        return reviews.map(serializeReview);
    } catch (error) {
        console.error('Failed to fetch instructor reviews:', error);
        throw new Error('Failed to fetch instructor reviews');
    }
}

export async function fetchDepartmentReviews(departmentId: number) {
    try {
        const reviews = await prisma.review.findMany({
            where: { departmentId },
            include: reviewInclude,
            orderBy: [
                { score: 'desc' },
                { createdAt: 'desc' }
            ],
        });
        return reviews.map(serializeReview);
    } catch (error) {
        console.error('Failed to fetch department reviews:', error);
        throw new Error('Failed to fetch department reviews');
    }
}

// --- UPDATED CREATE FUNCTION ---
export async function createReview({
    classId,
    instructorId,
    departmentId,
    userId,
    rating,
    content,
    title,
    courseCode,
    isOnlineCourse,
    difficulty,
    // New Arguments
    clarity,
    workload,
    support,
    // -------------
    wouldTakeAgain,
    attendanceMandatory,
    grade,
    tags,
}: {
    classId: number;
    instructorId: number;
    departmentId: number;
    userId: number;
    rating: number;
    content: string;
    title?: string | null;
    courseCode: string;
    isOnlineCourse: boolean;
    difficulty?: number | null;
    // New Types
    clarity: number;
    workload: number;
    support: number;
    // ---------
    wouldTakeAgain?: boolean | null;
    attendanceMandatory?: boolean | null;
    grade?: string | null;
    tags?: string[];
}): Promise<ReviewType> {
    try {
        const review = await prisma.review.create({
            data: {
                classId,
                instructorId,
                departmentId,
                userId,
                rating,
                content,
                title: title ?? null,
                courseCode,
                isOnlineCourse,
                difficulty: difficulty ?? null,
                
                // Saving new metrics to Database
                clarity,
                workload,
                support,
                
                wouldTakeAgain: wouldTakeAgain ?? null,
                attendanceMandatory: attendanceMandatory ?? null,
                grade: grade ?? null,
                tags: tags ?? [],
            },
            include: reviewInclude,
        });
        return serializeReview(review);
    } catch (error) {
        console.error('Failed to create review:', error);
        throw new Error('Failed to create review');
    }
}

export async function fetchGPADistributions(
    classId: number,
    departmentId: number,
) {
    try {
        const data = await prisma.distribution.findMany({
            where: {
                classId,
                class: {
                    departmentId,
                },
            },
            select: {
                class: {
                    select: {
                        code: true,
                        name: true,
                        department: {
                            select: {
                                code: true,
                                name: true,
                            },
                        },
                    },
                },
                instructor: true,
                term: true,
                studentHeadcount: true,
                avgCourseGrade: true,
                grades: true,
            },
            orderBy: {
                term: 'asc',
            },
        });

        // Parse the grades data as { [key: string]: number }
        const parsedData = data.map((item: (typeof data)[number]) => ({
            ...item,
            grades: item.grades as { [key: string]: number },
        }));

        return parsedData;
    } catch (error) {
        console.error('Database Error:', error);
        throw new Error('Failed to fetch GPA distributions.');
    }
}

export const getSearch = async (search: string, classPage: number, instructorPage: number, departmentPage: number, limit: number) => {
    const classSkip = (classPage - 1) * limit;
    const instructorSkip = (instructorPage - 1) * limit;
    const departmentSkip = (departmentPage - 1) * limit;

    try {
        const [classResults, instructorResults, departmentResults, classCount, instructorCount, departmentCount] = await Promise.all([
            prisma.class.findMany({
                where: search ? {
                    OR: [
                        {
                            code: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                        {
                            name: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                    ],
                } : undefined,
                select: {
                    id: true,
                    code: true,
                    name: true,
                    department: {
                        select: {
                            code: true,
                            name: true,
                        },
                    },
                },
                orderBy: {
                    code: 'asc',
                },
                skip: classSkip,
                take: limit,
            }),
            prisma.instructor.findMany({
                where: search ? {
                    name: {
                        contains: search,
                        mode: 'insensitive',
                    },
                } : undefined,
                select: {
                    id: true,
                    name: true,
                },
                orderBy: {
                    name: 'asc',
                },
                skip: instructorSkip,
                take: limit,
            }),
            prisma.department.findMany({
                where: search ? {
                    OR: [
                        {
                            code: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                        {
                            name: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                    ],
                } : undefined,
                select: {
                    id: true,
                    code: true,
                    name: true,
                },
                orderBy: {
                    code: 'asc',
                },
                skip: departmentSkip,
                take: limit,
            }),
            prisma.class.count({
                where: search ? {
                    OR: [
                        {
                            code: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                        {
                            name: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                    ],
                } : undefined,
            }),
            prisma.instructor.count({
                where: search ? {
                    name: {
                        contains: search,
                        mode: 'insensitive',
                    },
                } : undefined,
            }),
            prisma.department.count({
                where: search ? {
                    OR: [
                        {
                            code: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                        {
                            name: {
                                contains: search,
                                mode: 'insensitive',
                            },
                        },
                    ],
                } : undefined,
            }),
        ]);

        return {
            classes: classResults,
            instructors: instructorResults,
            departments: departmentResults,
            classCount,
            instructorCount,
            departmentCount,
        };
    } catch (error) {
        console.error('Database Error:', error);
        throw new Error('Failed to perform search.');
    }
};

export async function getClassByCode(code: string) {
    try {
        const classData = await prisma.class.findFirst({
            where: {
                OR: [
                    { code: code },
                    { code: { contains: code, mode: 'insensitive' } },
                ],
            },
            include: {
                department: true,
            },
        });
        return classData;
    } catch (error) {
        console.error('Failed to fetch class data:', error);
        throw new Error('Failed to fetch class data');
    }
}

export async function getInstructorById(instructorId: number) {
    try {
        const instructor = await prisma.instructor.findUnique({
            where: { id: instructorId },
        });
        return instructor;
    } catch (error) {
        console.error('Failed to fetch instructor data:', error);
        throw new Error('Failed to fetch instructor data');
    }
}

export async function fetchInstructorClasses(instructorId: number) {
    try {
        const instructorClasses = await prisma.distribution.findMany({
            where: {
                instructorId,
            },
            select: {
                class: {
                    select: {
                        code: true,
                        name: true,
                    },
                },
                term: true,
                studentHeadcount: true,
                avgCourseGrade: true,
                grades: true,
            },
            orderBy: {
                term: 'asc',
            },
        });

        const parsedData = instructorClasses.map((item: (typeof instructorClasses)[number]) => ({
            ...item,
            gradePercentages: item.grades as { [key: string]: number },
        }));

        return parsedData;
    } catch (error) {
        console.error('Failed to fetch instructor classes:', error);
        throw new Error('Failed to fetch instructor classes');
    }
}

export async function getDepartmentByCode(code: string) {
    try {
        const department = await prisma.department.findUnique({
            where: { code },
        });
        return department;
    } catch (error) {
        console.error('Failed to fetch department data:', error);
        throw new Error('Failed to fetch department data');
    }
}

export async function fetchDepartmentClasses(departmentId: number) {
    try {
        const departmentClasses = await prisma.class.findMany({
            where: {
                departmentId,
            },
            select: {
                code: true,
                name: true,
            },
        });
        return departmentClasses;
    } catch (error) {
        console.error('Failed to fetch department classes:', error);
        throw new Error('Failed to fetch department classes');
    }
}

export async function fetchDepartmentInstructors(departmentName: string) {
    try {
        const departmentInstructors = await prisma.instructor.findMany({
            where: {
                department: departmentName,
            },
            select: {
                id: true,
                name: true,
                department: true,
            },
        });
        return departmentInstructors;
    } catch (error) {
        console.error('Failed to fetch department instructors:', error);
        throw new Error('Failed to fetch department instructors');
    }
}

// --- NEW: Efficient Aggregation for Instructors ---
export async function getInstructorAggregates(instructorId: number) {
    try {
        const aggregates = await prisma.review.aggregate({
            where: { 
                instructorId,
                clarity: { gt: 0 }, // Only count new reviews
            },
            _avg: {
                clarity: true,
                support: true,
                workload: true,
                difficulty: true
            },
            _count: {
                _all: true
            }
        });

        return {
            avgClarity: aggregates._avg.clarity || 0,
            avgSupport: aggregates._avg.support || 0,
            avgWorkload: aggregates._avg.workload || 0,
            avgDifficulty: aggregates._avg.difficulty || 0,
            count: aggregates._count._all
        };
    } catch (error) {
        console.error('Failed to fetch instructor aggregates:', error);
        return { avgClarity: 0, avgSupport: 0, avgWorkload: 0, avgDifficulty: 0, count: 0 };
    }
}

export async function fetchDepartmentGrades(departmentId: number) {
    try {
        const departmentGrades = await prisma.distribution.findMany({
            where: {
                class: {
                    departmentId,
                },
            },
            select: {
                class: {
                    select: {
                        code: true,
                        name: true,
                    },
                },
                grades: true,
                studentHeadcount: true,
                avgCourseGrade: true,
            },
        });

        const parsedData = departmentGrades.map((item: (typeof departmentGrades)[number]) => ({
            ...item,
            gradePercentages: item.grades as { [key: string]: number },
        }));

        return parsedData;
    } catch (error) {
        console.error('Failed to fetch department grades:', error);
        throw new Error('Failed to fetch department grades');
    }
}

// --- NEW FUNCTION GOES HERE (OUTSIDE THE ONE ABOVE) ---
export async function getDepartmentAggregates(departmentId: number) {
    try {
        const aggregates = await prisma.review.aggregate({
            where: { 
                departmentId,
                clarity: { gt: 0 },
            },
            _avg: {
                clarity: true,
                support: true,
                workload: true,
                difficulty: true
            },
            _count: {
                _all: true
            }
        });

        return {
            avgClarity: aggregates._avg.clarity || 0,
            avgSupport: aggregates._avg.support || 0,
            avgWorkload: aggregates._avg.workload || 0,
            avgDifficulty: aggregates._avg.difficulty || 0,
            count: aggregates._count._all
        };
    } catch (error) {
        console.error('Failed to fetch department aggregates:', error);
        return { avgClarity: 0, avgSupport: 0, avgWorkload: 0, avgDifficulty: 0, count: 0 };
    }
}