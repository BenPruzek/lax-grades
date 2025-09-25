import { fetchClassReviews } from '@/lib/data';
import { NextResponse } from 'next/server';

export async function GET(
    request: Request,
    { params }: { params: { classId: string } }
) {
    try {
        const classId = parseInt(params.classId);
        
        if (isNaN(classId)) {
            return NextResponse.json({ error: 'Invalid class ID.' }, { status: 400 });
        }

        const reviews = await fetchClassReviews(classId);
        return NextResponse.json(reviews);
    } catch (error) {
        console.error('Failed to fetch class reviews:', error);
        return NextResponse.json({ error: 'Something went wrong.' }, { status: 500 });
    }
}
