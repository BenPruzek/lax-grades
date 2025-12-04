"use client";

import { useState, useEffect, type ChangeEvent } from 'react';
import { useSession } from 'next-auth/react';
import { useRouter } from 'next/navigation';
import { REVIEW_TAG_OPTIONS, GRADE_OPTIONS } from '@/lib/review-constants';

// 1. Define what an Instructor looks like
interface InstructorOption {
    id: number;
    name: string;
}

interface ReviewFormProps {
    classId: number;
    // We make this optional now, because the user might select it in the form
    instructorId?: number; 
    departmentId: number;
    classCode: string;
    // 2. Add this prop to receive the list of professors
    availableInstructors?: InstructorOption[]; 
    onSuccess?: () => void;
}

const MAX_TAG_SELECTION = 3;
const MAX_CONTENT_LENGTH = 350;

export default function ReviewForm({ 
    classId, 
    instructorId, 
    departmentId, 
    classCode, 
    availableInstructors = [], // Default to empty list if not provided
    onSuccess 
}: ReviewFormProps) {
    const { data: session } = useSession();
    const router = useRouter();
    
    // 3. State for the selected instructor
    // If an instructorId was passed in props (pre-selected), use it. Otherwise, empty.
    const [selectedInstructorId, setSelectedInstructorId] = useState<string>(
        instructorId ? instructorId.toString() : ''
    );

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

    const [clarity, setClarity] = useState('');
    const [workload, setWorkload] = useState('');
    const [support, setSupport] = useState('');

    useEffect(() => {
        setCourseCode(classCode);
    }, [classCode]);

    const handleTagToggle = (tag: string) => {
        setSelectedTags((prev) => {
            if (prev.includes(tag)) {
                return prev.filter((item) => item !== tag);
            }
            if (prev.length >= MAX_TAG_SELECTION) {
                setError('You can select up to 3 tags.');
                return prev;
            }
            setError(null);
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
        
        const clarityValue = parseInt(clarity, 10);
        const workloadValue = parseInt(workload, 10);
        const supportValue = parseInt(support, 10);

        const wouldTakeAgainValue = wouldTakeAgain === 'yes';
        const attendanceValue = attendanceMandatory === '' ? null : attendanceMandatory === 'yes';
        const gradeValue = grade || null;

        // 4. Validate that an instructor is selected
        const finalInstructorId = parseInt(selectedInstructorId, 10);
        if (!finalInstructorId || isNaN(finalInstructorId)) {
            setError('Please select an instructor.');
            setIsSubmitting(false);
            return;
        }

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

        if (!difficulty || Number.isNaN(difficultyValue)) {
            setError('Please rate the difficulty.');
            setIsSubmitting(false);
            return;
        }

        if (!clarity || !workload || !support) {
            setError('Please complete the Clarity, Workload, and Support ratings.');
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
                    instructorId: finalInstructorId, // Use the selected ID
                    departmentId,
                    title: title || null,
                    rating,
                    content: trimmedContent,
                    courseCode: trimmedCourseCode,
                    isOnlineCourse,
                    difficulty: difficultyValue,
                    clarity: clarityValue,
                    workload: workloadValue,
                    support: supportValue,
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

            (event.target as HTMLFormElement).reset();
            setClarity('');
            setWorkload('');
            setSupport('');
            setDifficulty('');
            setWouldTakeAgain('');
            setAttendanceMandatory('');
            setGrade('');
            setSelectedTags([]);
            // Don't reset selectedInstructorId intentionally, user might want to review same prof? 
            // Or reset it if you prefer: setSelectedInstructorId('');
            
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

            {/* 5. NEW: Instructor Dropdown */}
            {availableInstructors.length > 0 ? (
                <div>
                    <label htmlFor="instructor" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                        Select Instructor *
                    </label>
                    <select
                        id="instructor"
                        value={selectedInstructorId}
                        onChange={(e) => setSelectedInstructorId(e.target.value)}
                        className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100"
                        disabled={isSubmitting}
                        required
                    >
                        <option value="">-- Choose an Instructor --</option>
                        {availableInstructors.map((inst) => (
                            <option key={inst.id} value={inst.id}>
                                {inst.name}
                            </option>
                        ))}
                    </select>
                </div>
            ) : (
                // Fallback if no list provided (e.g. if parent component isn't updated yet)
                <div className="text-sm text-gray-500 italic">
                    Reviewing for: {courseCode} (Instructor selection unavailable)
                </div>
            )}

            <div>
                <label htmlFor="courseCode" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Confirm Course Code *
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

            {/* Quality Triad Section */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 p-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg border border-gray-200 dark:border-gray-700">
                <h4 className="col-span-1 md:col-span-2 text-sm font-semibold text-gray-900 dark:text-gray-100 mb-2">Detailed Ratings</h4>
                
                <div>
                    <label htmlFor="clarity" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">Clarity of Instruction *</label>
                    <select id="clarity" value={clarity} onChange={(e) => setClarity(e.target.value)} disabled={isSubmitting} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 dark:bg-gray-800 dark:text-gray-100">
                        <option value="">Select...</option>
                        <option value="5">5 - Crystal Clear</option>
                        <option value="4">4 - Clear</option>
                        <option value="3">3 - Average</option>
                        <option value="2">2 - Confusing</option>
                        <option value="1">1 - Unintelligible</option>
                    </select>
                </div>

                <div>
                    <label htmlFor="workload" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">Workload Volume *</label>
                    <select id="workload" value={workload} onChange={(e) => setWorkload(e.target.value)} disabled={isSubmitting} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 dark:bg-gray-800 dark:text-gray-100">
                        <option value="">Select...</option>
                        <option value="5">5 - Very Light</option>
                        <option value="4">4 - Light</option>
                        <option value="3">3 - Moderate</option>
                        <option value="2">2 - Heavy</option>
                        <option value="1">1 - Extreme</option>
                    </select>
                </div>

                <div>
                    <label htmlFor="support" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">Instructor Support *</label>
                    <select id="support" value={support} onChange={(e) => setSupport(e.target.value)} disabled={isSubmitting} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 dark:bg-gray-800 dark:text-gray-100">
                        <option value="">Select...</option>
                        <option value="5">5 - Very Accessible</option>
                        <option value="4">4 - Helpful</option>
                        <option value="3">3 - Average</option>
                        <option value="2">2 - Hard to Reach</option>
                        <option value="1">1 - Ghosted Me</option>
                    </select>
                </div>

                <div>
                    <label htmlFor="difficulty" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">Course Difficulty *</label>
                    <select id="difficulty" name="difficulty" value={difficulty} onChange={(event) => setDifficulty(event.target.value)} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 dark:bg-gray-800 dark:text-gray-100" disabled={isSubmitting} required>
                        <option value="">Select difficulty</option>
                        <option value="1">1 - Very Easy</option>
                        <option value="2">2</option>
                        <option value="3">3</option>
                        <option value="4">4</option>
                        <option value="5">5 - Very Difficult</option>
                    </select>
                </div>
            </div>

            <div>
                <label htmlFor="rating" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Overall rating *
                </label>
                <select id="rating" name="rating" required className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100" disabled={isSubmitting}>
                    <option value="">Select a rating</option>
                    <option value="5">5 - Awesome</option>
                    <option value="4">4 - Good</option>
                    <option value="3">3 - Average</option>
                    <option value="2">2 - Poor</option>
                    <option value="1">1 - Awful</option>
                </select>
            </div>

            <fieldset className="space-y-2">
                <legend className="text-sm font-medium text-gray-700 dark:text-gray-300">Would you choose this instructor again? *</legend>
                <div className="flex gap-4">
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input type="radio" name="wouldTakeAgain" value="yes" checked={wouldTakeAgain === 'yes'} onChange={(event) => setWouldTakeAgain(event.target.value)} disabled={isSubmitting} className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300" />
                        Yes
                    </label>
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input type="radio" name="wouldTakeAgain" value="no" checked={wouldTakeAgain === 'no'} onChange={(event) => setWouldTakeAgain(event.target.value)} disabled={isSubmitting} className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300" />
                        No
                    </label>
                </div>
            </fieldset>

            <fieldset className="space-y-2">
                <legend className="text-sm font-medium text-gray-700 dark:text-gray-300">Was attendance required?</legend>
                <div className="flex gap-4">
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input type="radio" name="attendanceMandatory" value="yes" checked={attendanceMandatory === 'yes'} onChange={(event) => setAttendanceMandatory(event.target.value)} disabled={isSubmitting} className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300" />
                        Yes
                    </label>
                    <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                        <input type="radio" name="attendanceMandatory" value="no" checked={attendanceMandatory === 'no'} onChange={(event) => setAttendanceMandatory(event.target.value)} disabled={isSubmitting} className="h-4 w-4 text-red-600 focus:ring-red-500 border-gray-300" />
                        No
                    </label>
                </div>
            </fieldset>

            <div>
                <label htmlFor="grade" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                    Grade received (optional)
                </label>
                <select id="grade" name="grade" value={grade} onChange={(event) => setGrade(event.target.value)} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100" disabled={isSubmitting}>
                    <option value="">Select grade</option>
                    {GRADE_OPTIONS.map((option) => (
                        <option key={option} value={option}>
                            {option}
                        </option>
                    ))}
                </select>
            </div>

            <div>
                <p className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Select up to 3 highlights that describe your experience</p>
                <div className="flex flex-wrap gap-2">
                    {REVIEW_TAG_OPTIONS.map((tag) => {
                        const isSelected = selectedTags.includes(tag);
                        return (
                            <button key={tag} type="button" onClick={() => handleTagToggle(tag)} disabled={isSubmitting && !isSelected} className={`px-3 py-1.5 text-xs font-medium rounded-full border transition ${isSelected ? 'bg-red-600 text-white border-red-600' : 'bg-white text-gray-700 border-gray-300 dark:bg-gray-900 dark:text-gray-200 dark:border-gray-700'}`}>
                                {tag}
                            </button>
                        );
                    })}
                </div>
            </div>

            <div>
                <label htmlFor="title" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">Title (optional)</label>
                <input id="title" name="title" type="text" maxLength={100} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100" placeholder="Brief summary of your experience" disabled={isSubmitting} />
            </div>

            <div>
                <label htmlFor="content" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">Share your experience *</label>
                <textarea id="content" name="content" rows={6} required maxLength={MAX_CONTENT_LENGTH} onChange={handleContentChange} className="w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm focus:outline-none focus:ring-red-500 focus:border-red-500 dark:bg-gray-800 dark:text-gray-100" placeholder="Focus on the instructor's teaching style, clarity, and the pace of the course" disabled={isSubmitting} />
                <div className="mt-1 flex items-center justify-between text-xs text-gray-500 dark:text-gray-400">
                    <span>{contentLength}/{MAX_CONTENT_LENGTH}</span>
                    <span className="text-right">Keep it constructive, specific, and respectful.</span>
                </div>
            </div>

            {error && <p className="text-sm text-red-600 dark:text-red-400" role="alert">{error}</p>}

            <p className="text-xs text-gray-500 dark:text-gray-400">
                By submitting, you confirm that your feedback aligns with the LAX Grades community guidelines and terms of use.
            </p>

            <button type="submit" disabled={isSubmitting} className="w-full px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-red-500 disabled:opacity-50 disabled:cursor-not-allowed">
                {isSubmitting ? 'Submitting...' : 'Submit Rating'}
            </button>
        </form>
    );
}
