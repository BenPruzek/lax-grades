"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { motion } from "framer-motion";
import { cn } from "@/lib/utils";
import Image from "next/image";
import { useSession } from "next-auth/react";

const routes = [
    {
        name: "Home",
        path: "/",
    },
    {
        name: "All Classes",
        path: "/search",
    },
];

export default function Header() {
    const activePathname = usePathname();
    const { data: session, status } = useSession();

    return (
        <header className="flex items-center justify-between border-b border-black/10 dark:border-white/10 h-14 px-9 bg-white dark:bg-black">
            <Link href="/" className="flex flex-row items-center space-x-1">
                <Image src="/LAXGradesDistributionLogo.svg" alt="LAX Grades Logo" width={50} height={50} />
                <h1 className="text-gray-800 dark:text-gray-100"><strong><span className="text-red-900 dark:text-red-400">LAX</span>GRADES</strong></h1>
            </Link>

            <div className="flex items-center gap-6">
                <nav className="h-full">
                    <ul className="flex gap-x-6 h-full text-sm">
                        {routes.map((route) => (
                            <li
                                key={route.path}
                                className={cn(
                                    "text-gray-900 dark:text-gray-100 flex items-center relative transition",
                                    {
                                        "text-red-800 dark:text-red-400": activePathname === route.path,
                                        "text-gray-700 dark:text-gray-400": activePathname !== route.path,
                                    }
                                )}
                            >
                                <Link href={route.path}>{route.name}</Link>

                                {activePathname === route.path && (
                                    <motion.div
                                        layoutId="header-active-link"
                                        className="bg-red-800 dark:bg-red-400 h-1 w-full absolute bottom-0"
                                    ></motion.div>
                                )}
                            </li>
                        ))}
                    </ul>
                </nav>

                {status === "loading" ? null : session ? null : (
                    <div className="flex items-center gap-3 text-sm">
                        <Link
                            href="/sign-in"
                            className="text-gray-700 dark:text-gray-300 hover:text-red-800 dark:hover:text-red-400 transition"
                        >
                            Log In
                        </Link>
                        <Link
                            href="/sign-up"
                            className="px-3 py-1.5 rounded-md bg-red-700 text-white hover:bg-red-800 focus:outline-none focus:ring-2 focus:ring-red-500"
                        >
                            Sign Up
                        </Link>
                    </div>
                )}
            </div>
        </header>
    );
}
