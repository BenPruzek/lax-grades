import Link from "next/link";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

export default function SearchCard({ data, type }: { data: any; type: "classes" | "instructors" | "departments" }) {
    let title = "";
    let href = "";

    if (type === "classes") {
        title = `${data.code} - ${data.name}`;
        href = `/class/${data.code}`;
    } else if (type === "instructors") {
        title = data.name;
        href = `/instructor/${data.id}`;
    } else if (type === "departments") {
        title = data.code;
        href = `/department/${data.code}`;
    }

    return (
        <Link href={href}>
            <Card className="hover:bg-gray-100 dark:hover:bg-gray-800 bg-white dark:bg-gray-900 transition-colors border-gray-200 dark:border-gray-800 duration-200 flex flex-row items-center w-full px-6 mb-4">
                <CardHeader>
                    <CardTitle className="text-gray-900 dark:text-gray-100">{title}</CardTitle>
                </CardHeader>
                <CardContent></CardContent>
            </Card>
        </Link>
    );
}