"use client";

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { signIn } from 'next-auth/react';

export default function SignUpForm() {
    const router = useRouter();
    const [error, setError] = useState<string | null>(null);
    const [isSubmitting, setIsSubmitting] = useState(false);

    const onSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
        event.preventDefault();
        setError(null);
        setIsSubmitting(true);

        const formData = new FormData(event.currentTarget);
        const name = formData.get('name') as string;
        const email = formData.get('email') as string;
        const password = formData.get('password') as string;
        const confirmPassword = formData.get('confirmPassword') as string;

        if (password !== confirmPassword) {
            setError('Passwords do not match.');
            setIsSubmitting(false);
            return;
        }

        try {
            const response = await fetch('/api/auth/register', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, email, password }),
            });

            if (!response.ok) {
                const data = await response.json();
                setError(data?.error ?? 'Something went wrong.');
                return;
            }

            const signInResult = await signIn('credentials', {
                email,
                password,
                redirect: false,
            });

            if (signInResult?.error) {
                setError(signInResult.error);
                return;
            }

            router.push('/');
            router.refresh();
        } catch (submitError) {
            console.error('Sign up failed:', submitError);
            setError('Something went wrong. Please try again.');
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <div className="min-h-screen flex items-center justify-center bg-gray-50 dark:bg-gray-900 px-4">
            <div className="max-w-md w-full space-y-8 bg-white dark:bg-gray-800 p-10 rounded-xl shadow-lg">
                <div className="text-center">
                    <h2 className="text-3xl font-extrabold text-gray-900 dark:text-white">Create your account</h2>
                    <p className="mt-2 text-sm text-gray-600 dark:text-gray-300">
                        Already have an account?{' '}
                        <a href="/sign-in" className="text-red-600 hover:text-red-700 dark:text-red-400">
                            Sign in
                        </a>
                    </p>
                </div>
                <form className="mt-8 space-y-6" onSubmit={onSubmit}>
                    <div className="space-y-4">
                        <div>
                            <label htmlFor="name" className="sr-only">
                                Full name
                            </label>
                            <input
                                id="name"
                                name="name"
                                type="text"
                                required
                                className="appearance-none block w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm placeholder-gray-500 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-red-500 focus:border-red-500 sm:text-sm"
                                placeholder="Full name"
                                disabled={isSubmitting}
                            />
                        </div>
                        <div>
                            <label htmlFor="email" className="sr-only">
                                Email address
                            </label>
                            <input
                                id="email"
                                name="email"
                                type="email"
                                autoComplete="email"
                                required
                                className="appearance-none block w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm placeholder-gray-500 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-red-500 focus:border-red-500 sm:text-sm"
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
                                autoComplete="new-password"
                                minLength={8}
                                required
                                className="appearance-none block w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm placeholder-gray-500 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-red-500 focus:border-red-500 sm:text-sm"
                                placeholder="Password"
                                disabled={isSubmitting}
                            />
                        </div>
                        <div>
                            <label htmlFor="confirmPassword" className="sr-only">
                                Confirm password
                            </label>
                            <input
                                id="confirmPassword"
                                name="confirmPassword"
                                type="password"
                                autoComplete="new-password"
                                minLength={8}
                                required
                                className="appearance-none block w-full px-3 py-2 border border-gray-300 dark:border-gray-700 rounded-md shadow-sm placeholder-gray-500 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-red-500 focus:border-red-500 sm:text-sm"
                                placeholder="Confirm password"
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
                            {isSubmitting ? 'Creating account…' : 'Create account'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}
