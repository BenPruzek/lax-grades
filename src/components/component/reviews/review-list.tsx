import type { Review } from '@/lib/types';

export default function ReviewList({ reviews }: { reviews: Review[] }) {
    if (!reviews.length) {
        return (
            <div className="p-6 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg text-center text-gray-600 dark:text-gray-300">
                No reviews yet. Be the first to share your experience.
            </div>
        );
    }

    return (
        <div className="space-y-4">
            {reviews.map((review) => (
                <article key={review.id} className="border border-gray-200 dark:border-gray-700 rounded-lg p-6 bg-white dark:bg-gray-900 shadow-sm">
                    <div className="flex items-center justify-between">
                        <h3 className="text-lg font-semibold text-gray-900 dark:text-gray-100">
                            {review.title ?? `Review by ${review.user.name ?? review.user.email}`}
                        </h3>
                        <span className="text-sm text-gray-500 dark:text-gray-400">
                            {new Date(review.createdAt).toLocaleDateString('en-US', { 
                                year: 'numeric', 
                                month: 'short', 
                                day: 'numeric' 
                            })}
                        </span>
                    </div>
                    <div className="mt-2 flex items-center gap-3 text-sm text-gray-600 dark:text-gray-300">
                        <span>Rating: {review.rating}/5</span>
                        <span>&bull;</span>
                        <span>{review.instructor.name}</span>
                    </div>
                    <p className="mt-4 text-gray-700 dark:text-gray-200 whitespace-pre-wrap">{review.content}</p>
                </article>
            ))}
        </div>
    );
}
