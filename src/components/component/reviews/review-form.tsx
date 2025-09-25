"use client";

import { useState } from 'react';
import { useSession } from 'next-auth/react';
import { useRouter } from 'next/navigation';

interface ReviewFormProps {
    classId: number;
    instructorId: number;
    departmentId: number;
    onSuccess?: () => void;
}

export default function ReviewForm({ classId, instructorId, departmentId, onSuccess }: ReviewFormProps) {
    const { data: session } = useSession();
    const router = useRouter();
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [error, setError] = useState<string | null>(null);

    if (!session) {
        return (
            <div className="p-6 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg text-center">
                <p className="text-gray-600 dark:text-gray-300 mb-4">
                    You must be signed in to submit a review.
                </p>
                <button
                    onClick={() => router.push('/sign-in')}
                    className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-red-500"
                >
                    Sign In
                </button>
            </div>
        );
    }

    const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
        event.preventDefault();
        setError(null);
        setIsSubmitting(true);

        const formData = new FormData(event.currentTarget);
        const title = formData.get('title') as string;
        const rating = parseInt(formData.get('rating') as string);
        const content = formData.get('content') as string;

        if (!rating || rating < 1 || rating > 5) {
            setError('Please select a rating between 1 and 5.');
            setIsSubmitting(false);
            return;
        }

        if (!content.trim()) {
            setError('Please provide review content.');
            setIsSubmitting(false);
            return;
        }

        try {
            const response = await fetch('/api/reviews', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    classId,
                    instructorId,
                    departmentId,
                    title: title || null,
                    rating,
                    content: content.trim(),
                }),
            });

            if (!response.ok) {
                const data = await response.json();
                setError(data?.error ?? 'Failed to submit review.');
                return;
            }

            // Reset form
            (event.target as HTMLFormElement).reset();
            onSuccess?.();
        } catch (submitError) {
            console.error('Review submission failed:', submitError);
            setError('Something went wrong. Please try again.');
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <form onSubmit={handleSubmit} className="space-y-4 p-6 border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-900">
            <h3 className="text-lg font-semibold text-gray-900 dark:text-gray-100">Write a Review</h3>
            
            <div>
                <label htmlFor="title" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Title (optional)
                </label>
                <input
                    id="title"
                    name="title"
                    type="text"
                    maxLength={100}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    placeholder="Brief summary of your experience"
                    disabled={isSubmitting}
                />
            </div>

            <div>
                <label htmlFor="rating" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Rating *
                </label>
                <select
                    id="rating"
                    name="rating"
                    required
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    disabled={isSubmitting}
                >
                    <option value="">Select a rating</option>
                    <option value="5">5 - Excellent</option>
                    <option value="4">4 - Good</option>
                    <option value="3">3 - Average</option>
                    <option value="2">2 - Poor</option>
                    <option value="1">1 - Terrible</option>
                </select>
            </div>

            <div>
                <label htmlFor="content" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Review *
                </label>
                <textarea
                    id="content"
                    name="content"
                    rows={4}
                    required
                    maxLength={2000}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    placeholder="Share your experience with this class and instructor..."
                    disabled={isSubmitting}
                />
            </div>

            {error && (
                <p className="text-sm text-red-600 dark:text-red-400" role="alert">
                    {error}
                </p>
            )}

            <button
                type="submit"
                disabled={isSubmitting}
                className="w-full px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-red-500 disabled:opacity-50 disabled:cursor-not-allowed"
            >
                {isSubmitting ? 'Submitting...' : 'Submit Review'}
            </button>
        </form>
    );
}
