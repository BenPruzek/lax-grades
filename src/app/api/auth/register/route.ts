import { prisma } from '@/lib/prisma';
import { hash } from 'bcryptjs';
import { NextResponse } from 'next/server';
import { z } from 'zod';

const registerSchema = z.object({
    email: z.string().email(),
    password: z.string().min(8),
    name: z.string().min(1).max(100),
});

export async function POST(request: Request) {
    try {
        const body = await request.json();
        const result = registerSchema.safeParse(body);

        if (!result.success) {
            return NextResponse.json({ error: 'Invalid input.' }, { status: 400 });
        }

        const { email, password, name } = result.data;

        const existingUser = await prisma.user.findUnique({
            where: { email },
        });

        if (existingUser) {
            return NextResponse.json({ error: 'An account with that email already exists.' }, { status: 409 });
        }

        const hashedPassword = await hash(password, 10);

        await prisma.user.create({
            data: {
                email,
                password: hashedPassword,
                name,
            },
        });

        return NextResponse.json({ success: true }, { status: 201 });
    } catch (error) {
        console.error('Failed to register user:', error);
        return NextResponse.json({ error: 'Something went wrong.' }, { status: 500 });
    }
}
