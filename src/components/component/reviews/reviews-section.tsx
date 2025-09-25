"use client";

import { useState, useEffect } from 'react';
import { useSession } from 'next-auth/react';
import ReviewForm from './review-form';
import ReviewList from './review-list';
import type { Review } from '@/lib/types';

interface ReviewsSectionProps {
    classId: number;
    instructorId: number;
    departmentId: number;
    initialReviews?: Review[];
}

export default function ReviewsSection({ 
    classId, 
    instructorId, 
    departmentId, 
    initialReviews = [] 
}: ReviewsSectionProps) {
    const { data: session } = useSession();
    const [reviews, setReviews] = useState<Review[]>(initialReviews);
    const [showForm, setShowForm] = useState(false);
    const [loading, setLoading] = useState(false);

    const fetchReviews = async () => {
        setLoading(true);
        try {
            const response = await fetch(`/api/reviews?classId=${classId}`);
            if (response.ok) {
                const data = await response.json();
                setReviews(data);
            }
        } catch (error) {
            console.error('Failed to fetch reviews:', error);
        } finally {
            setLoading(false);
        }
    };

    const handleReviewSubmitted = () => {
        setShowForm(false);
        fetchReviews(); // Refresh reviews after submission
    };

    return (
        <div className="space-y-6">
            <div className="flex items-center justify-between">
                <h2 className="text-2xl font-bold text-gray-900 dark:text-gray-100">
                    Student Reviews ({reviews.length})
                </h2>
                {session && (
                    <button
                        onClick={() => setShowForm(!showForm)}
                        className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-red-500"
                    >
                        {showForm ? 'Cancel' : 'Write Review'}
                    </button>
                )}
            </div>

            {showForm && (
                <ReviewForm
                    classId={classId}
                    instructorId={instructorId}
                    departmentId={departmentId}
                    onSuccess={handleReviewSubmitted}
                />
            )}

            {loading ? (
                <div className="text-center py-8">
                    <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-red-600"></div>
                    <p className="mt-2 text-gray-600 dark:text-gray-300">Loading reviews...</p>
                </div>
            ) : (
                <ReviewList reviews={reviews} />
            )}
        </div>
    );
}
