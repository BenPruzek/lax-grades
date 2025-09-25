"use client";

import { useState } from 'react';
import { signIn, type SignInResponse } from 'next-auth/react';
import { useRouter } from 'next/navigation';
import Link from 'next/link';

export default function SignInForm() {
    const router = useRouter();
    const [error, setError] = useState<string | null>(null);
    const [isSubmitting, setIsSubmitting] = useState(false);

    const onSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
        event.preventDefault();
        setError(null);
        setIsSubmitting(true);

        const formData = new FormData(event.currentTarget);
        const email = formData.get('email') as string;
        const password = formData.get('password') as string;

        try {
            const result: SignInResponse | undefined = await signIn('credentials', {
                email,
                password,
                redirect: false,
            });

            if (result?.error) {
                setError(result.error);
                return;
            }

            router.push('/');
            router.refresh();
        } catch (submitError) {
            console.error('Sign in failed:', submitError);
            setError('Something went wrong. Please try again.');
        } finally {
            setIsSubmitting(false);
        }
    };

    const handleGoogleSignIn = () => {
        setError(null);
        void signIn('google', { callbackUrl: '/' });
    };

    return (
        <div className="min-h-screen flex items-center justify-center bg-gray-50 dark:bg-gray-900 px-4">
            <div className="max-w-md w-full space-y-8 bg-white dark:bg-gray-800 p-10 rounded-xl shadow-lg">
                <div className="text-center">
                    <h2 className="text-3xl font-extrabold text-gray-900 dark:text-white">Sign in to your account</h2>
                    <p className="mt-2 text-sm text-gray-600 dark:text-gray-300">
                        <Link href="/sign-up" className="text-red-600 hover:text-red-700 dark:text-red-400">
                            create a new account
                        </Link>
                    </p>
                </div>
                <div className="mt-8 space-y-6">
                    <button
                        type="button"
                        onClick={handleGoogleSignIn}
                        disabled={isSubmitting}
                        className="w-full flex justify-center items-center gap-2 py-2 px-4 border border-gray-300 dark:border-gray-700 text-sm font-medium rounded-md text-gray-700 dark:text-gray-200 bg-white dark:bg-gray-900 hover:bg-gray-50 dark:hover:bg-gray-800 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-red-500 disabled:opacity-50"
                    >
                        <span>Continue with Google</span>
                    </button>

                    <div className="relative flex items-center">
                        <div className="flex-grow border-t border-gray-300 dark:border-gray-700" />
                        <span className="px-3 text-sm text-gray-500 dark:text-gray-400">or</span>
                        <div className="flex-grow border-t border-gray-300 dark:border-gray-700" />
                    </div>

                    <form className="space-y-6" onSubmit={onSubmit}>
                        <div className="rounded-md shadow-sm -space-y-px">
                            <div>
                                <label htmlFor="email-address" className="sr-only">
                                    Email address
                                </label>
                                <input
                                    id="email-address"
                                    name="email"
                                    type="email"
                                    autoComplete="email"
                                    required
                                    className="appearance-none rounded-none relative block w-full px-3 py-2 border border-gray-300 dark:border-gray-700 placeholder-gray-500 text-gray-900 dark:text-gray-100 rounded-t-md focus:outline-none focus:ring-red-500 focus:border-red-500 focus:z-10 sm:text-sm"
                                    placeholder="Email address"
                                    disabled={isSubmitting}
                                />
                            </div>
                            <div>
                                <label htmlFor="password" className="sr-only">
                                    Password
                                </label>
                                <input
                                    id="password"
                                    name="password"
                                    type="password"
                                    autoComplete="current-password"
                                    required
                                    className="appearance-none rounded-none relative block w-full px-3 py-2 border border-gray-300 dark:border-gray-700 placeholder-gray-500 text-gray-900 dark:text-gray-100 rounded-b-md focus:outline-none focus:ring-red-500 focus:border-red-500 focus:z-10 sm:text-sm"
                                    placeholder="Password"
                                    disabled={isSubmitting}
                                />
                            </div>
                        </div>

                        {error ? (
                            <p className="text-sm text-red-600 dark:text-red-400" role="alert">
                                {error}
                            </p>
                        ) : null}

                        <div>
                            <button
                                type="submit"
                                disabled={isSubmitting}
                                className="group relative w-full flex justify-center py-2 px-4 border border-transparent text-sm font-medium rounded-md text-white bg-red-600 hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-red-500"
                            >
                                {isSubmitting ? 'Signing in…' : 'Sign in'}
                            </button>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    );
}
