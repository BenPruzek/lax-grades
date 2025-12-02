import { getServerAuthSession } from '@/lib/auth';
import { prisma } from '@/lib/prisma';
import { NextResponse } from 'next/server';
import { Prisma } from '@prisma/client';

export async function POST(
    request: Request,
    { params }: { params: { reviewId: string } }
) {
    const session = await getServerAuthSession();
    if (!session?.user?.id) {
        return NextResponse.json({ error: 'Authentication required' }, { status: 401 });
    }

    const userId = parseInt(session.user.id);
    const reviewId = parseInt(params.reviewId);

    if (isNaN(reviewId)) {
        return NextResponse.json({ error: 'Invalid review ID' }, { status: 400 });
    }

    const body = await request.json();
    const { type } = body; // "LIKE" or "DISLIKE"

    if (type !== 'LIKE' && type !== 'DISLIKE') {
        return NextResponse.json({ error: 'Invalid vote type' }, { status: 400 });
    }

    try {
        // Check if user already voted
        const existingVote = await prisma.reviewVote.findUnique({
            where: {
                userId_reviewId: {
                    userId,
                    reviewId,
                },
            },
        });

        if (existingVote) {
            if (existingVote.type === type) {
                // Toggle off (remove vote)
                await prisma.$transaction(async (tx: Prisma.TransactionClient) => {
                    await tx.reviewVote.delete({
                        where: { id: existingVote.id },
                    });
                    
                    // Update counts
                    const updateData = type === 'LIKE' 
                        ? { likes: { decrement: 1 }, score: { decrement: 1 } }
                        : { dislikes: { decrement: 1 }, score: { increment: 1 } };

                    await tx.review.update({
                        where: { id: reviewId },
                        data: updateData,
                    });
                });
                return NextResponse.json({ message: 'Vote removed' });
            } else {
                // Change vote (e.g. LIKE -> DISLIKE)
                await prisma.$transaction(async (tx: Prisma.TransactionClient) => {
                    await tx.reviewVote.update({
                        where: { id: existingVote.id },
                        data: { type },
                    });

                    // Update counts
                    // If was LIKE now DISLIKE: likes -1, dislikes +1, score -2
                    // If was DISLIKE now LIKE: dislikes -1, likes +1, score +2
                    const updateData = type === 'LIKE'
                        ? { likes: { increment: 1 }, dislikes: { decrement: 1 }, score: { increment: 2 } }
                        : { likes: { decrement: 1 }, dislikes: { increment: 1 }, score: { decrement: 2 } };

                    await tx.review.update({
                        where: { id: reviewId },
                        data: updateData,
                    });
                });
                return NextResponse.json({ message: 'Vote updated' });
            }
        } else {
            // New vote
            await prisma.$transaction(async (tx: Prisma.TransactionClient) => {
                await tx.reviewVote.create({
                    data: {
                        userId,
                        reviewId,
                        type,
                    },
                });

                // Update counts
                const updateData = type === 'LIKE'
                    ? { likes: { increment: 1 }, score: { increment: 1 } }
                    : { dislikes: { increment: 1 }, score: { decrement: 1 } };

                await tx.review.update({
                    where: { id: reviewId },
                    data: updateData,
                });
            });
            return NextResponse.json({ message: 'Vote added' });
        }
    } catch (error) {
        console.error('Vote failed:', error);
        return NextResponse.json({ error: 'Failed to process vote' }, { status: 500 });
    }
}
