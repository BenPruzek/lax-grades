import { getServerAuthSession } from '@/lib/auth';
import { createReview } from '@/lib/data';
import { NextResponse } from 'next/server';
import { createReviewSchema } from '@/lib/validation/review-schema';
import { moderateContent } from '@/lib/moderation';
import rateLimit from '@/lib/rate-limit';

const limiter = rateLimit({
    interval: 60 * 1000, // 60 seconds
    uniqueTokenPerInterval: 500, // Max 500 users per second
});

export async function POST(request: Request) {
    try {
        const session = await getServerAuthSession();
        
        if (!session?.user?.id) {
            return NextResponse.json({ error: 'Authentication required.' }, { status: 401 });
        }

        try {
            await limiter.check(5, session.user.id); // 5 requests per minute
        } catch {
            return NextResponse.json({ error: 'Rate limit exceeded' }, { status: 429 });
        }

        const body = await request.json();
        
        // NOTE: You must update createReviewSchema in 'src/lib/validation/review-schema.ts'
        // to include clarity, workload, and support, or this check will fail!
        const result = createReviewSchema.safeParse(body);

        if (!result.success) {
            // Log the error so you can see if it's the schema rejecting the new fields
            console.error("Validation failed:", result.error); 
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
            // --- NEW METRICS ---
            clarity,
            workload,
            support,
            // -------------------
            wouldTakeAgain,
            attendanceMandatory,
            grade,
            tags,
        } = result.data;

        // Moderate content
        const contentCheck = moderateContent(content);
        if (contentCheck.flagged) {
            return NextResponse.json({ error: 'Review contains inappropriate content.' }, { status: 400 });
        }

        if (title) {
            const titleCheck = moderateContent(title);
            if (titleCheck.flagged) {
                return NextResponse.json({ error: 'Review title contains inappropriate content.' }, { status: 400 });
            }
        }

        // NOTE: You must update the createReview function in 'src/lib/data.ts'
        // to accept these new arguments!
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
            // --- PASS NEW METRICS TO DATABASE ---
            clarity,
            workload,
            support,
            // ------------------------------------
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
