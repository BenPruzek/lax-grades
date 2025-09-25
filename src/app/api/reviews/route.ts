import { getServerAuthSession } from '@/lib/auth';
import { createReview } from '@/lib/data';
import { NextResponse } from 'next/server';
import { z } from 'zod';

const TAG_OPTIONS = [
    'Tough Grader',
    'Get Ready To Read',
    'Participation Matters',
    'Extra Credit',
    'Group Projects',
    'Amazing Lectures',
    'Clear Grading Criteria',
    'Gives Good Feedback',
    'Inspirational',
    'Lots Of Homework',
    'Hilarious',
    'Beware Of Pop Quizzes',
    'So Many Papers',
    'Caring',
    'Respected',
    'Lecture Heavy',
    'Test Heavy',
    'Graded By Few Things',
    'Accessible Outside Class',
    'Online Savvy',
];

const createReviewSchema = z.object({
    classId: z.number().int().positive(),
    instructorId: z.number().int().positive(),
    departmentId: z.number().int().positive(),
    title: z.string().max(100).nullable().optional(),
    rating: z.number().int().min(1).max(5),
    content: z.string().min(1).max(350),
    courseCode: z.string().min(1).max(20),
    isOnlineCourse: z.boolean(),
    difficulty: z.number().int().min(1).max(5),
    wouldTakeAgain: z.boolean(),
    attendanceMandatory: z.boolean().nullable(),
    grade: z.string().max(5).nullable(),
    tags: z
        .array(z.enum(TAG_OPTIONS as [string, ...string[]]))
        .max(3)
        .optional()
        .default([]),
});

export async function POST(request: Request) {
    try {
        const session = await getServerAuthSession();
        
        if (!session?.user?.id) {
            return NextResponse.json({ error: 'Authentication required.' }, { status: 401 });
        }

        const body = await request.json();
        const result = createReviewSchema.safeParse(body);

        if (!result.success) {
            return NextResponse.json({ error: 'Invalid input.' }, { status: 400 });
        }

        const {
            classId,
            instructorId,
            departmentId,
            title,
            rating,
            content,
            courseCode,
            isOnlineCourse,
            difficulty,
            wouldTakeAgain,
            attendanceMandatory,
            grade,
            tags,
        } = result.data;

        const review = await createReview({
            classId,
            instructorId,
            departmentId,
            userId: parseInt(session.user.id),
            title: title ?? null,
            rating,
            content,
            courseCode,
            isOnlineCourse,
            difficulty,
            wouldTakeAgain,
            attendanceMandatory,
            grade,
            tags,
        });

        return NextResponse.json(review, { status: 201 });
    } catch (error) {
        console.error('Failed to create review:', error);
        return NextResponse.json({ error: 'Something went wrong.' }, { status: 500 });
    }
}
