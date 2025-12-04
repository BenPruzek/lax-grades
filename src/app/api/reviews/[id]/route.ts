import { getServerAuthSession } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { NextResponse } from "next/server";

export async function DELETE(
  request: Request,
  { params }: { params: { id: string } }
) {
  try {
    const session = await getServerAuthSession();
    
    // 1. Security: Are you logged in?
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const reviewId = parseInt(params.id);

    // 2. Database: Find the review
    const review = await prisma.review.findUnique({
      where: { id: reviewId },
      select: { userId: true },
    });

    if (!review) {
      return NextResponse.json({ error: "Review not found" }, { status: 404 });
    }

    // 3. Ownership: Is this YOUR review?
    if (review.userId !== parseInt(session.user.id)) {
      return NextResponse.json({ error: "You can only delete your own reviews" }, { status: 403 });
    }

    // 4. Destroy it
    await prisma.review.delete({
      where: { id: reviewId },
    });

    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("Delete failed:", error);
    return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
  }
}
