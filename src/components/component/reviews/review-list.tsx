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
                <article key={review.id} className="border border-gray-200 dark:border-gray-700 rounded-lg p-6 bg-white dark:bg-gray-900 shadow-sm space-y-4">
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
                </article>
            ))}
        </div>
    );
}
