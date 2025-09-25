import { getServerAuthSession } from '@/lib/auth';
import { createReview } from '@/lib/data';
import { NextResponse } from 'next/server';
import { createReviewSchema } from '@/lib/validation/review-schema';

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
