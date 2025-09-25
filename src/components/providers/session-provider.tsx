"use client";

import { SessionProvider } from 'next-auth/react';
import type { Session } from 'next-auth';
import type { ReactNode } from 'react';

export default function AuthSessionProvider({ children, session }: { children: ReactNode; session?: Session | null }) {
    return <SessionProvider session={session}>{children}</SessionProvider>;
}
