import { NextResponse } from "next/server";
import { PrismaClient } from "@prisma/client";

export const dynamic = "force-dynamic";

const prisma = new PrismaClient();

export async function GET(req: Request) {
  try {
    const { searchParams } = new URL(req.url);
    const q = (searchParams.get("q") || "").trim();
    const limitParam = Number(searchParams.get("limit") || 5);
    const limit = Math.max(1, Math.min(10, limitParam));

    if (!q) {
      return NextResponse.json({ status: "ok", data: { classes: [], instructors: [], departments: [] } });
    }

    // Helper to rank results
    const scoreText = (text: string, query: string) => {
      const t = text.toLowerCase();
      const qq = query.toLowerCase();
      if (t === qq) return 100;
      if (t.startsWith(qq)) return 60;
      if (t.includes(qq)) return 30;
      return 0;
    };

    const [classPrefix, classNameContains, classCodeContains] = await Promise.all([
      prisma.class.findMany({
        where: { code: { startsWith: q, mode: "insensitive" } },
        select: {
          id: true,
          code: true,
          name: true,
          department: { select: { code: true, name: true } },
        },
        take: limit,
      }),
      prisma.class.findMany({
        where: { name: { contains: q, mode: "insensitive" } },
        select: {
          id: true,
          code: true,
          name: true,
          department: { select: { code: true, name: true } },
        },
        take: limit,
      }),
      prisma.class.findMany({
        where: { code: { contains: q, mode: "insensitive" } },
        select: {
          id: true,
          code: true,
          name: true,
          department: { select: { code: true, name: true } },
        },
        take: limit,
      }),
    ]);

    // Deduplicate classes by id
    const classMap = new Map<number, any>();
    [...classPrefix, ...classNameContains, ...classCodeContains].forEach((c) => {
      if (!classMap.has(c.id)) classMap.set(c.id, c);
    });
    const classesRanked = Array.from(classMap.values())
      .map((c) => ({
        id: c.id,
        type: "class" as const,
        title: c.code,
        subtitle: `${c.name} — ${c.department?.name ?? ""}`.trim(),
        href: `/class/${c.code}`,
        score: Math.max(scoreText(c.code, q) + 10, scoreText(c.name, q)), // bias towards code matches
      }))
      .sort((a, b) => b.score - a.score)
      .slice(0, limit);

    const [instructorPrefix, instructorContains] = await Promise.all([
      prisma.instructor.findMany({
        where: { name: { startsWith: q, mode: "insensitive" } },
        select: { id: true, name: true, department: true },
        take: limit,
      }),
      prisma.instructor.findMany({
        where: { name: { contains: q, mode: "insensitive" } },
        select: { id: true, name: true, department: true },
        take: limit,
      }),
    ]);
    const instructorMap = new Map<number, any>();
    [...instructorPrefix, ...instructorContains].forEach((i) => {
      if (!instructorMap.has(i.id)) instructorMap.set(i.id, i);
    });
    const instructorsRanked = Array.from(instructorMap.values())
      .map((i) => ({
        id: i.id,
        type: "instructor" as const,
        title: i.name,
        subtitle: i.department || "Instructor",
        href: `/instructor/${i.id}`,
        score: scoreText(i.name, q),
      }))
      .sort((a, b) => b.score - a.score)
      .slice(0, limit);

    const [deptPrefix, deptNameContains, deptCodeContains] = await Promise.all([
      prisma.department.findMany({
        where: { code: { startsWith: q, mode: "insensitive" } },
        select: { id: true, code: true, name: true },
        take: limit,
      }),
      prisma.department.findMany({
        where: { name: { contains: q, mode: "insensitive" } },
        select: { id: true, code: true, name: true },
        take: limit,
      }),
      prisma.department.findMany({
        where: { code: { contains: q, mode: "insensitive" } },
        select: { id: true, code: true, name: true },
        take: limit,
      }),
    ]);
    const deptMap = new Map<number, any>();
    [...deptPrefix, ...deptNameContains, ...deptCodeContains].forEach((d) => {
      if (!deptMap.has(d.id)) deptMap.set(d.id, d);
    });
    const departmentsRanked = Array.from(deptMap.values())
      .map((d) => ({
        id: d.id,
        type: "department" as const,
        title: d.code,
        subtitle: d.name,
        href: `/department/${d.code}`,
        score: Math.max(scoreText(d.code, q) + 10, scoreText(d.name, q)),
      }))
      .sort((a, b) => b.score - a.score)
      .slice(0, limit);

    return NextResponse.json({
      status: "ok",
      data: {
        classes: classesRanked.map(({ score, ...rest }) => rest),
        instructors: instructorsRanked.map(({ score, ...rest }) => rest),
        departments: departmentsRanked.map(({ score, ...rest }) => rest),
      },
    });
  } catch (err) {
    console.error("Suggest API error", err);
    return NextResponse.json({ status: "error" }, { status: 500 });
  }
}
