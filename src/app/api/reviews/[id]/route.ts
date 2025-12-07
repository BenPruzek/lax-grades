import { fetchClassReviews } from '@/lib/data';
import { NextResponse } from 'next/server';

export async function GET(
    request: Request,
    { params }: { params: { id: string } } // UPDATED: changed classId to id
) {
    try {
        // UPDATED: Use params.id
        const classId = parseInt(params.id);

        if (isNaN(classId)) {
            return NextResponse.json({ error: 'Invalid class ID' }, { status: 400 });
        }

        const reviews = await fetchClassReviews(classId);
        return NextResponse.json(reviews);
    } catch (error) {
        console.error('Failed to fetch reviews:', error);
        return NextResponse.json({ error: 'Internal Server Error' }, { status: 500 });
    }
}