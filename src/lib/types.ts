import { Prisma } from "@prisma/client";
import type { ReviewTagOption } from "@/lib/review-constants";
import { z } from "zod";

export interface ClassData {
    id: number;
    name: string;
    code: string;
    department: {
        id: number;
        name: string;
        code: string;
    };
}

export interface FormData {
    firstName: string;
    lastName: string;
    email: string;
    message: string;
}

export interface GradeDistribution {
    term: string;
    grades: Prisma.JsonValue;
    studentHeadcount: number;
    avgCourseGrade: number;
    class: {
        code: string;
        name: string;
        department: {
            code: string;
            name: string;
        };
    };
    instructor?: {
        id: number;
        name: string;
        department: string;
    } | null;
}

export interface Department {
    id: number;
    code: string;
    name: string;
}

export interface GradePercentages {
    [key: string]: number;
}

export interface UserSummary {
    id: number;
    name?: string | null;
    email: string;
}

export interface Review {
    id: number;
    title?: string | null;
    rating: number;
    content: string;
    courseCode: string;
    isOnlineCourse: boolean;
    
    // Existing Metric
    difficulty?: number | null;
    
    // --- NEW QUALITY METRICS ---
    clarity: number;
    workload: number;
    support: number;
    // ---------------------------

    wouldTakeAgain?: boolean | null;
    attendanceMandatory?: boolean | null;
    grade?: string | null;
    tags: ReviewTagOption[];
    likes: number;
    dislikes: number;
    score: number;
    userVote?: 'LIKE' | 'DISLIKE' | null;
    createdAt: string;
    updatedAt: string;
    user: UserSummary;
    class: {
        id: number;
        code: string;
        name: string;
    };
    instructor: {
        id: number;
        name: string;
    };
    department: {
        id: number;
        code: string;
        name: string;
    };
}
