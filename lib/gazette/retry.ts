// Supabase's API gateway occasionally answers a burst of 5xx/401s while the
// database itself is fine (see the "Failed to get API key info", "Bad
// Gateway" and "Gateway Timeout" entries in the Vercel runtime log). A read
// that fails that way usually succeeds a moment later, so the gazette's
// server reads go through this helper: it re-runs the query a couple of
// times with a short pause before giving up. Reads are idempotent, so
// retrying on any error is safe; the last result (or thrown error) is what
// the caller sees.

type ReadResult = { error: { message: string } | null };

export type RetryOptions = {
  /** Total attempts, including the first (default 3). */
  attempts?: number;
  /** Pause before the first retry; each later retry waits one step longer (default 250 ms). */
  delayMs?: number;
};

export async function readWithRetry<R extends ReadResult>(
  run: () => PromiseLike<R>,
  { attempts = 3, delayMs = 250 }: RetryOptions = {},
): Promise<R> {
  const total = Math.max(1, Math.floor(attempts));
  let last: R | undefined;
  let thrown: unknown;
  for (let attempt = 1; attempt <= total; attempt++) {
    if (attempt > 1 && delayMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs * (attempt - 1)));
    }
    try {
      last = await run();
      thrown = undefined;
      if (!last.error) return last;
    } catch (error) {
      thrown = error;
      last = undefined;
    }
  }
  if (last) return last;
  throw thrown;
}
