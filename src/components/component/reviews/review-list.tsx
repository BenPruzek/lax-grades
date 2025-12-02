import type { Review } from '@/lib/types';
import { useSession } from 'next-auth/react';
import Link from 'next/link';
import { Button } from '@/components/ui/button';
import { Lock, ThumbsUp, ThumbsDown } from 'lucide-react';
import { useState } from 'react';
import { useRouter } from 'next/navigation';

export default function ReviewList({ reviews: initialReviews }: { reviews: Review[] }) {
    const { data: session } = useSession();
    const router = useRouter();
    const isAuthenticated = !!session;
    const [reviews, setReviews] = useState(initialReviews);

    const handleVote = async (reviewId: number, type: 'LIKE' | 'DISLIKE') => {
        if (!isAuthenticated) {
            router.push('/sign-in');
            return;
        }

        // Optimistic update
        setReviews(currentReviews => currentReviews.map(review => {
            if (review.id !== reviewId) return review;

            let newLikes = review.likes;
            let newDislikes = review.dislikes;
            let newUserVote = review.userVote;

            if (review.userVote === type) {
                // Toggle off
                newUserVote = null;
                if (type === 'LIKE') newLikes--;
                else newDislikes--;
            } else {
                // Change vote or new vote
                if (review.userVote === 'LIKE') newLikes--;
                else if (review.userVote === 'DISLIKE') newDislikes--;

                newUserVote = type;
                if (type === 'LIKE') newLikes++;
                else newDislikes++;
            }

            return {
                ...review,
                likes: newLikes,
                dislikes: newDislikes,
                score: newLikes - newDislikes,
                userVote: newUserVote
            };
        }));

        try {
            const response = await fetch(`/api/reviews/${reviewId}/vote`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ type }),
            });

            if (!response.ok) {
                throw new Error('Vote failed');
            }
        } catch (error) {
            console.error('Vote failed:', error);
            // Revert optimistic update (optional, but good practice)
            // For now we just log error
        }
    };

    if (!reviews.length) {
        return (
            <div className="p-6 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg text-center text-gray-600 dark:text-gray-300">
                No reviews yet. Be the first to share your experience.
            </div>
        );
    }

    const visibleReviews = isAuthenticated ? reviews : reviews.slice(0, 3); // Show a few more blurred out

    return (
        <div className="space-y-4 relative">
            {visibleReviews.map((review, index) => {
                const isBlurred = !isAuthenticated && index > 0;
                
                return (
                    <article 
                        key={review.id} 
                        className={`border border-gray-200 dark:border-gray-700 rounded-lg p-6 bg-white dark:bg-gray-900 shadow-sm space-y-4 ${isBlurred ? 'blur-sm select-none' : ''}`}
                        aria-hidden={isBlurred}
                    >
                        <header className="flex flex-wrap items-start justify-between gap-3">
                            <div>
                                <h3 className="text-lg font-semibold text-gray-900 dark:text-gray-100">
                                    {review.title ?? `Review by ${review.user.name ?? review.user.email}`}
                                </h3>
                                <p className="mt-1 text-sm text-gray-600 dark:text-gray-300">
                                    {review.courseCode} • {review.instructor.name} • {review.isOnlineCourse ? 'Online' : 'In-person'}
                                </p>
                            </div>
                            <time className="text-sm text-gray-500 dark:text-gray-400">
                                {new Date(review.createdAt).toLocaleDateString('en-US', {
                                    year: 'numeric',
                                    month: 'short',
                                    day: 'numeric',
                                })}
                            </time>
                        </header>

                        <dl className="grid gap-2 text-sm text-gray-700 dark:text-gray-300 sm:grid-cols-2">
                            <div className="flex gap-2">
                                <dt className="font-medium">Rating:</dt>
                                <dd>{review.rating}/5</dd>
                            </div>
                            <div className="flex gap-2">
                                <dt className="font-medium">Difficulty:</dt>
                                <dd>{review.difficulty ?? '—'} / 5</dd>
                            </div>
                            <div className="flex gap-2">
                                <dt className="font-medium">Would take again:</dt>
                                <dd>{review.wouldTakeAgain === null || review.wouldTakeAgain === undefined ? '—' : review.wouldTakeAgain ? 'Yes' : 'No'}</dd>
                            </div>
                            <div className="flex gap-2">
                                <dt className="font-medium">Attendance mandatory:</dt>
                                <dd>
                                    {review.attendanceMandatory === null || review.attendanceMandatory === undefined
                                        ? '—'
                                        : review.attendanceMandatory
                                        ? 'Yes'
                                        : 'No'}
                                </dd>
                            </div>
                            <div className="flex gap-2">
                                <dt className="font-medium">Grade received:</dt>
                                <dd>{review.grade ?? '—'}</dd>
                            </div>
                        </dl>

                        {review.tags.length ? (
                            <div className="flex flex-wrap gap-2">
                                {review.tags.map((tag) => (
                                    <span
                                        key={tag}
                                        className="inline-flex items-center rounded-full bg-red-50 px-3 py-1 text-xs font-medium text-red-700 dark:bg-red-900/30 dark:text-red-200"
                                    >
                                        {tag}
                                    </span>
                                ))}
                            </div>
                        ) : null}

                        <p className="text-gray-700 dark:text-gray-200 whitespace-pre-wrap">{review.content}</p>

                        <div className="flex items-center gap-4 pt-2 border-t border-gray-100 dark:border-gray-800">
                            <button 
                                onClick={() => handleVote(review.id, 'LIKE')}
                                className={`flex items-center gap-1.5 text-sm font-medium transition-colors ${review.userVote === 'LIKE' ? 'text-green-600 dark:text-green-400' : 'text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-gray-200'}`}
                                disabled={isBlurred}
                            >
                                <ThumbsUp className={`w-4 h-4 ${review.userVote === 'LIKE' ? 'fill-current' : ''}`} />
                                <span>{review.likes}</span>
                            </button>
                            <button 
                                onClick={() => handleVote(review.id, 'DISLIKE')}
                                className={`flex items-center gap-1.5 text-sm font-medium transition-colors ${review.userVote === 'DISLIKE' ? 'text-red-600 dark:text-red-400' : 'text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-gray-200'}`}
                                disabled={isBlurred}
                            >
                                <ThumbsDown className={`w-4 h-4 ${review.userVote === 'DISLIKE' ? 'fill-current' : ''}`} />
                                <span>{review.dislikes}</span>
                            </button>
                        </div>
                    </article>
                );
            })}
            
            {!isAuthenticated && reviews.length > 1 && (
                <div className="absolute inset-0 top-[200px] z-10 flex flex-col items-center justify-center bg-gradient-to-b from-transparent via-white/80 to-white dark:via-gray-950/80 dark:to-gray-950 backdrop-blur-[2px]">
                    <div className="p-6 text-center max-w-md mx-auto bg-white dark:bg-gray-900 rounded-xl shadow-lg border border-gray-200 dark:border-gray-800">
                        <div className="w-12 h-12 bg-red-100 dark:bg-red-900/20 rounded-full flex items-center justify-center mx-auto mb-4">
                            <Lock className="w-6 h-6 text-red-600 dark:text-red-400" />
                        </div>
                        <h3 className="text-xl font-bold text-gray-900 dark:text-white mb-2">
                            Join to see full reviews
                        </h3>
                        <p className="text-gray-600 dark:text-gray-300 mb-6">
                            Sign up for free to read all reviews, ratings, and grade data from students.
                        </p>
                        <div className="flex flex-col sm:flex-row gap-3 justify-center">
                            <Link href="/sign-up" className="w-full sm:w-auto">
                                <Button className="w-full bg-red-600 hover:bg-red-700 text-white font-semibold">
                                    Sign up for free
                                </Button>
                            </Link>
                            <Link href="/sign-in" className="w-full sm:w-auto">
                                <Button variant="outline" className="w-full">
                                    Log in
                                </Button>
                            </Link>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
