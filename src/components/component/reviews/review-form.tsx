"use client";

import { useState, useEffect, type ChangeEvent } from 'react';
import { useSession } from 'next-auth/react';
import { useRouter } from 'next/navigation';
import { REVIEW_TAG_OPTIONS, GRADE_OPTIONS } from '@/lib/review-constants';

interface ReviewFormProps {
    classId: number;
    instructorId: number;
    departmentId: number;
    classCode: string;
    onSuccess?: () => void;
}

const MAX_TAG_SELECTION = 3;
const MAX_CONTENT_LENGTH = 350;

export default function ReviewForm({ classId, instructorId, departmentId, classCode, onSuccess }: ReviewFormProps) {
    const { data: session } = useSession();
    const router = useRouter();
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [courseCode, setCourseCode] = useState(classCode);
    const [isOnlineCourse, setIsOnlineCourse] = useState(false);
    const [difficulty, setDifficulty] = useState('');
    const [wouldTakeAgain, setWouldTakeAgain] = useState('');
    const [attendanceMandatory, setAttendanceMandatory] = useState('');
    const [grade, setGrade] = useState('');
    const [selectedTags, setSelectedTags] = useState<string[]>([]);
    const [contentLength, setContentLength] = useState(0);

    useEffect(() => {
        setCourseCode(classCode);
    }, [classCode]);

    const handleTagToggle = (tag: string) => {
        setSelectedTags((prev) => {
            if (prev.includes(tag)) {
                const updated = prev.filter((item) => item !== tag);
                if (error && error.includes('3 tags')) {
                    setError(null);
                }
                return updated;
            }

            if (prev.length >= MAX_TAG_SELECTION) {
                setError('You can select up to 3 tags.');
                return prev;
            }

            if (error && error.includes('3 tags')) {
                setError(null);
            }

            return [...prev, tag];
        });
    };

    const handleContentChange = (event: ChangeEvent<HTMLTextAreaElement>) => {
        setContentLength(event.target.value.length);
    };

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
        const title = (formData.get('title') as string) ?? '';
        const rating = parseInt((formData.get('rating') as string) ?? '', 10);
        const content = (formData.get('content') as string) ?? '';
        const trimmedCourseCode = courseCode.trim();
        const trimmedContent = content.trim();
        const difficultyValue = parseInt(difficulty, 10);
        const wouldTakeAgainValue = wouldTakeAgain === 'yes';
        const attendanceValue = attendanceMandatory === '' ? null : attendanceMandatory === 'yes';
        const gradeValue = grade || null;

        if (!trimmedCourseCode) {
            setError('Please select a course code.');
            setIsSubmitting(false);
            return;
        }

        if (!rating || rating < 1 || rating > 5) {
            setError('Please select a rating between 1 and 5.');
            setIsSubmitting(false);
            return;
        }

        if (!difficulty || Number.isNaN(difficultyValue) || difficultyValue < 1 || difficultyValue > 5) {
            setError('Please rate the difficulty.');
            setIsSubmitting(false);
            return;
        }

        if (!wouldTakeAgain) {
            setError('Please indicate if you would take this professor again.');
            setIsSubmitting(false);
            return;
        }

        if (!trimmedContent) {
            setError('Please provide review content.');
            setIsSubmitting(false);
            return;
        }

        if (trimmedContent.length > MAX_CONTENT_LENGTH) {
            setError(`Review must be ${MAX_CONTENT_LENGTH} characters or less.`);
            setIsSubmitting(false);
            return;
        }

        if (selectedTags.length > MAX_TAG_SELECTION) {
            setError('You can select up to 3 tags.');
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
                    content: trimmedContent,
                    courseCode: trimmedCourseCode,
                    isOnlineCourse,
                    difficulty: difficultyValue,
                    wouldTakeAgain: wouldTakeAgainValue,
                    attendanceMandatory: attendanceValue,
                    grade: gradeValue,
                    tags: selectedTags,
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
        <form onSubmit={handleSubmit} className="space-y-6 p-6 border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-900">
            <h3 className="text-lg font-semibold text-gray-900 dark:text-gray-100">Write a Review</h3>

            <div>
                <label htmlFor="courseCode" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Select Course Code *
                </label>
                <input
                    id="courseCode"
                    name="courseCode"
                    type="text"
                    value={courseCode}
                    onChange={(event) => setCourseCode(event.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    disabled={isSubmitting}
                />
            </div>

            <div className="flex items-center gap-3">
                <input
                    id="isOnlineCourse"
                    name="isOnlineCourse"
                    type="checkbox"
                    checked={isOnlineCourse}
                    onChange={(event) => setIsOnlineCourse(event.target.checked)}
                    disabled={isSubmitting}
                    className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300 rounded"
                />
                <label htmlFor="isOnlineCourse" className="text-sm text-gray-700 dark:text-gray-300">
                    This is an online course
                </label>
            </div>

            <div>
                <label htmlFor="rating" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Rate your professor *
                </label>
                <select
                    id="rating"
                    name="rating"
                    required
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    disabled={isSubmitting}
                >
                    <option value="">Select a rating</option>
                    <option value="5">5 - Awesome</option>
                    <option value="4">4 - Good</option>
                    <option value="3">3 - Average</option>
                    <option value="2">2 - Poor</option>
                    <option value="1">1 - Awful</option>
                </select>
            </div>

            <div>
                <label htmlFor="difficulty" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    How difficult was this professor? *
                </label>
                <select
                    id="difficulty"
                    name="difficulty"
                    value={difficulty}
                    onChange={(event) => setDifficulty(event.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    disabled={isSubmitting}
                    required
                >
                    <option value="">Select difficulty</option>
                    <option value="1">1 - Very Easy</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="5">5 - Very Difficult</option>
                </select>
            </div>

            <fieldset className="space-y-2">
                <legend className="text-sm font-medium text-gray-700 dark:text-gray-300">Would you take this professor again? *</legend>
                <div className="flex gap-4">
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input
                            type="radio"
                            name="wouldTakeAgain"
                            value="yes"
                            checked={wouldTakeAgain === 'yes'}
                            onChange={(event) => setWouldTakeAgain(event.target.value)}
                            disabled={isSubmitting}
                            className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300"
                        />
                        Yes
                    </label>
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input
                            type="radio"
                            name="wouldTakeAgain"
                            value="no"
                            checked={wouldTakeAgain === 'no'}
                            onChange={(event) => setWouldTakeAgain(event.target.value)}
                            disabled={isSubmitting}
                            className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300"
                        />
                        No
                    </label>
                </div>
            </fieldset>

            <fieldset className="space-y-2">
                <legend className="text-sm font-medium text-gray-700 dark:text-gray-300">Was attendance mandatory?</legend>
                <div className="flex gap-4">
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input
                            type="radio"
                            name="attendanceMandatory"
                            value="yes"
                            checked={attendanceMandatory === 'yes'}
                            onChange={(event) => setAttendanceMandatory(event.target.value)}
                            disabled={isSubmitting}
                            className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300"
                        />
                        Yes
                    </label>
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input
                            type="radio"
                            name="attendanceMandatory"
                            value="no"
                            checked={attendanceMandatory === 'no'}
                            onChange={(event) => setAttendanceMandatory(event.target.value)}
                            disabled={isSubmitting}
                            className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300"
                        />
                        No
                    </label>
                </div>
            </fieldset>

            <div>
                <label htmlFor="grade" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Select grade received
                </label>
                <select
                    id="grade"
                    name="grade"
                    value={grade}
                    onChange={(event) => setGrade(event.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    disabled={isSubmitting}
                >
                    <option value="">Select grade</option>
                    {GRADE_OPTIONS.map((option) => (
                        <option key={option} value={option}>
                            {option}
                        </option>
                    ))}
                </select>
            </div>

            <div>
                <p className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Select up to 3 tags</p>
                <div className="flex flex-wrap gap-2">
                    {REVIEW_TAG_OPTIONS.map((tag) => {
                        const isSelected = selectedTags.includes(tag);
                        return (
                            <button
                                key={tag}
                                type="button"
                                onClick={() => handleTagToggle(tag)}
                                disabled={isSubmitting && !isSelected}
                                className={`px-3 py-1.5 text-xs font-medium rounded-full border transition ${
                                    isSelected
                                        ? 'bg-red-600 text-white border-red-600'
                                        : 'bg-white text-gray-700 border-gray-300 dark:bg-gray-900 dark:text-gray-200 dark:border-gray-700'
                                }`}
                            >
                                {tag}
                            </button>
                        );
                    })}
                </div>
            </div>

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
                <label htmlFor="content" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Write a Review *
                </label>
                <textarea
                    id="content"
                    name="content"
                    rows={6}
                    required
                    maxLength={MAX_CONTENT_LENGTH}
                    onChange={handleContentChange}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                    placeholder="Discuss the professor's professional abilities including teaching style and ability to convey the material clearly"
                    disabled={isSubmitting}
                />
                <div className="mt-1 flex items-center justify-between text-xs text-gray-500 dark:text-gray-400">
                    <span>{contentLength}/{MAX_CONTENT_LENGTH}</span>
                    <button
                        type="button"
                        className="text-red-600 dark:text-red-400 hover:underline"
                        onClick={() => window.open('https://www.ratemyprofessors.com/guidelines', '_blank')}
                    >
                        View all guidelines
                    </button>
                </div>
                <ul className="mt-2 text-xs text-gray-500 dark:text-gray-400 space-y-1 list-disc list-inside">
                    <li>Your rating could be removed if you use profanity or derogatory terms.</li>
                    <li>Don't claim that the professor shows bias or favoritism for or against students.</li>
                    <li>Don’t forget to proofread!</li>
                </ul>
            </div>

            {error && (
                <p className="text-sm text-red-600 dark:text-red-400" role="alert">
                    {error}
                </p>
            )}

            <p className="text-xs text-gray-500 dark:text-gray-400">
                By clicking the "Submit" button, I acknowledge that I have read and agreed to the Rate My Professors Site Guidelines, Terms of Use and Privacy Policy. Submitted data becomes the property of Rate My Professors.
            </p>

            <button
                type="submit"
                disabled={isSubmitting}
                className="w-full px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-red-500 disabled:opacity-50 disabled:cursor-not-allowed"
            >
                {isSubmitting ? 'Submitting...' : 'Submit Rating'}
            </button>
        </form>
    );
}
