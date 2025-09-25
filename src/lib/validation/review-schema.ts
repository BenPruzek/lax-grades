import { z } from 'zod';
import { REVIEW_TAG_OPTIONS, type ReviewTagOption } from '@/lib/review-constants';

const REVIEW_TAG_ENUM = z.enum([
    ...REVIEW_TAG_OPTIONS,
] as [ReviewTagOption, ...ReviewTagOption[]]);

export const createReviewSchema = z.object({
    classId: z.number().int().positive(),
    instructorId: z.number().int().positive(),
    departmentId: z.number().int().positive(),
    title: z.string().max(100).nullable().optional(),
    rating: z.number().int().min(1).max(5),
    content: z.string().min(1).max(350),
    courseCode: z.string().min(1).max(20),
    isOnlineCourse: z.boolean(),
    difficulty: z.number().int().min(1).max(5),
    wouldTakeAgain: z.boolean(),
    attendanceMandatory: z.boolean().nullable(),
    grade: z.string().max(20).nullable(),
    tags: z
        .array(REVIEW_TAG_ENUM)
        .max(3)
        .optional()
        .default([]),
});

export type CreateReviewInput = z.infer<typeof createReviewSchema>;
