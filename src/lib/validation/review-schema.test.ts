import { describe, expect, it } from 'vitest';
import { createReviewSchema } from './review-schema';

const BASE_REVIEW = {
    classId: 1,
    instructorId: 10,
    departmentId: 5,
    title: 'Engaging course',
    rating: 4,
    content: 'Clear lectures and fair assessments.',
    courseCode: 'CSCI1010',
    isOnlineCourse: false,
    difficulty: 3,
    wouldTakeAgain: true,
    attendanceMandatory: true,
    grade: 'A',
    tags: ['Clear Grading Criteria', 'Gives Good Feedback'],
} as const;

describe('createReviewSchema', () => {
    it('accepts a fully valid payload', () => {
        const result = createReviewSchema.safeParse(BASE_REVIEW);
        expect(result.success).toBe(true);
    });

    it('rejects payloads missing the rating', () => {
        const { rating, ...rest } = BASE_REVIEW;
        const result = createReviewSchema.safeParse(rest);
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.issues[0].path).toContain('rating');
        }
    });

    it('rejects payloads with content exceeding the character limit', () => {
        const result = createReviewSchema.safeParse({
            ...BASE_REVIEW,
            content: 'a'.repeat(351),
        });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.issues[0].path).toContain('content');
        }
    });

    it('rejects payloads with more than the maximum number of tags', () => {
        const result = createReviewSchema.safeParse({
            ...BASE_REVIEW,
            tags: [
                'Clear Grading Criteria',
                'Gives Good Feedback',
                'Inspirational',
                'Lots Of Homework',
            ],
        });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.issues[0].path).toContain('tags');
        }
    });

    it('allows nullable attendanceMandatory values', () => {
        const result = createReviewSchema.safeParse({
            ...BASE_REVIEW,
            attendanceMandatory: null,
        });
        expect(result.success).toBe(true);
    });

    it('rejects non-boolean values for wouldTakeAgain', () => {
        const result = createReviewSchema.safeParse({
            ...BASE_REVIEW,
            wouldTakeAgain: 'yes',
        });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.issues[0].path).toContain('wouldTakeAgain');
        }
    });
});
