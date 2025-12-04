import SignUpForm from '@/components/component/auth/sign-up-form';
import { Info } from 'lucide-react';

export default function SignUpPage() {
    return (
        <div className="max-w-md mx-auto mt-10 px-4">
            {/* Banner Start */}
            <div className="mb-6 rounded-md bg-blue-50 dark:bg-blue-900/20 p-4 border border-blue-200 dark:border-blue-800">
                <div className="flex">
                    <div className="flex-shrink-0">
                        <Info className="h-5 w-5 text-blue-400" aria-hidden="true" />
                    </div>
                    <div className="ml-3">
                        <h3 className="text-sm font-medium text-blue-800 dark:text-blue-200">
                            UWL Students Only
                        </h3>
                        <div className="mt-2 text-sm text-blue-700 dark:text-blue-300">
                            <p>
                                You must use your <strong>@uwlax.edu</strong> university email to create an account.
                            </p>
                        </div>
                    </div>
                </div>
            </div>
            {/* Banner End */}

            <SignUpForm />
        </div>
    );
}
