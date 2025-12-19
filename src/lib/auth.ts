import { prisma } from '@/lib/prisma';
import { compare, hash } from 'bcryptjs';
import { getServerSession } from 'next-auth';
import type { NextAuthOptions } from 'next-auth';
import type { JWT } from 'next-auth/jwt';
import CredentialsProvider from 'next-auth/providers/credentials';
import GoogleProvider from 'next-auth/providers/google';

type Credentials = {
    email: string;
    password: string;
};

const credentialsProvider = CredentialsProvider({
    name: 'Credentials',
    credentials: {
        email: { label: 'Email', type: 'email', placeholder: 'name@example.com' },
        password: { label: 'Password', type: 'password' },
    },
    async authorize(credentials: Credentials | undefined) {
        if (!credentials?.email || !credentials?.password) {
            throw new Error('Please provide both email and password.');
        }

        const user = await prisma.user.findUnique({
            where: { email: credentials.email },
        });

        if (!user || !user.password) {
            throw new Error('Invalid email or password.');
        }

        const isValid = await compare(credentials.password, user.password);

        if (!isValid) {
            throw new Error('Invalid email or password.');
        }

        return {
            id: user.id.toString(),
            email: user.email,
            name: user.name,
        };
    },
});

const providers: NextAuthOptions['providers'] = [credentialsProvider];

if (process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET) {
    providers.push(
        GoogleProvider({
            clientId: process.env.GOOGLE_CLIENT_ID,
            clientSecret: process.env.GOOGLE_CLIENT_SECRET,
        })
    );
}

export const authOptions: NextAuthOptions = {
    session: {
        strategy: 'jwt',
    },
    pages: {
        signIn: '/sign-in',
    },
    providers,
    callbacks: {
        async signIn({ user, account }) {
            // --- 1. GATEKEEPER LOGIC START ---
            if (!user.email) return false;

            const isUWL = user.email.endsWith('@uwlax.edu');
            
            // Admin Whitelist (You and Henry)
            const allowedAdmins = ['henryczup@gmail.com', 'benjamin.pruzek@gmail.com'];
            const isAdmin = allowedAdmins.includes(user.email);

            if (!isUWL && !isAdmin) {
                // If they are not a student and not an admin, block them.
                // This redirects them back to the sign-in page with an error.
                return '/sign-in?error=AccessDenied'; 
            }
            // --- GATEKEEPER LOGIC END ---

            if (account?.provider === 'google') {
                const existingUser = await prisma.user.findUnique({
                    where: { email: user.email },
                });

                if (existingUser) {
                    user.id = existingUser.id.toString();
                } else {
                    const createdUser = await prisma.user.create({
                        data: {
                            email: user.email,
                            name: user.name,
                            // Use a dummy password for OAuth users
                            password: await hash(process.env.OAUTH_DEFAULT_PASSWORD ?? 'oauth-user', 10),
                        },
                    });
                    user.id = createdUser.id.toString();
                }
            }

            return true;
        },
        async jwt({ token, user }): Promise<JWT> {
            if (user) {
                const typedUser = user as { id: string };
                token.id = typedUser.id;
            }

            if (!token.id && token.email) {
                const existingUser = await prisma.user.findUnique({
                    where: { email: token.email },
                    select: { id: true },
                });

                if (existingUser) {
                    token.id = existingUser.id.toString();
                }
            }
            return token as JWT;
        },
        async session({ session, token }) {
            if (session.user && token.id) {
                session.user.id = token.id as string;
            }
            return session;
        },
    },
    secret: process.env.NEXTAUTH_SECRET,
};

export function getServerAuthSession() {
    return getServerSession(authOptions);
}
