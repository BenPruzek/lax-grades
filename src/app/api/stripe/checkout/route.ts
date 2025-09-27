import { NextRequest, NextResponse } from "next/server";
import Stripe from "stripe";
import { z } from "zod";

const stripeSecretKey = process.env.STRIPE_SECRET_KEY;
const stripe = stripeSecretKey
  ? new Stripe(stripeSecretKey, {
      apiVersion: "2025-08-27.basil",
    })
  : null;

const requestSchema = z.object({
  priceId: z.string().min(1, "priceId is required"),
  quantity: z.number().int().positive().default(1),
  mode: z.enum(["payment", "subscription"]).default("subscription"),
  successUrl: z.string().url().optional(),
  cancelUrl: z.string().url().optional(),
  promotekitReferral: z.string().optional(),
  customerEmail: z.string().email().optional(),
  metadata: z.record(z.string()).optional(),
});

const isStripeCardError = (
  error: unknown,
): error is Stripe.errors.StripeCardError =>
  error instanceof Stripe.errors.StripeCardError;

export async function POST(request: NextRequest) {
  if (!stripe) {
    console.error("Stripe secret key is not configured.");
    return NextResponse.json(
      { error: "Stripe is not configured. Please set STRIPE_SECRET_KEY." },
      { status: 500 },
    );
  }

  try {
    const json = await request.json();
    const parsed = requestSchema.parse(json);

    const origin = request.headers.get("origin") ?? process.env.NEXT_PUBLIC_APP_URL;
    if (!origin) {
      return NextResponse.json(
        {
          error: "Unable to determine origin. Provide successUrl/cancelUrl or configure NEXT_PUBLIC_APP_URL.",
        },
        { status: 400 },
      );
    }

    const metadata: Stripe.MetadataParam = {
      ...parsed.metadata,
    };

    if (parsed.promotekitReferral) {
      metadata.promotekit_referral = parsed.promotekitReferral;
    }

    const session = await stripe.checkout.sessions.create({
      mode: parsed.mode,
      line_items: [
        {
          price: parsed.priceId,
          quantity: parsed.quantity,
        },
      ],
      success_url:
        parsed.successUrl ?? `${origin}/billing/success?session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: parsed.cancelUrl ?? `${origin}/billing/cancel`,
      customer_email: parsed.customerEmail,
      metadata: Object.keys(metadata).length > 0 ? metadata : undefined,
    });

    return NextResponse.json({ sessionId: session.id, url: session.url }, { status: 201 });
  } catch (error: unknown) {
    console.error(
      "Failed to create Stripe checkout session:",
      error instanceof Error ? error : String(error),
    );

    if (error instanceof z.ZodError) {
      return NextResponse.json({ error: error.flatten() }, { status: 400 });
    }

    if (isStripeCardError(error)) {
      return NextResponse.json({ error: error.message }, { status: error.statusCode ?? 402 });
    }

    return NextResponse.json({ error: "Unable to create Stripe checkout session." }, { status: 500 });
  }
}
