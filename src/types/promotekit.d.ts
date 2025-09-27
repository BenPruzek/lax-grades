declare global {
  interface Window {
    promotekit_referral?: string;
    promotekit?: {
      refer: (email: string, stripe_customer_id?: string) => void;
    };
  }
}

export {};
