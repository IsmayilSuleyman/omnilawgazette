import { describe, expect, it, vi } from "vitest";
import { readWithRetry } from "@/lib/gazette/retry";

const ok = { data: { issue_number: 12 }, error: null };
const gateway = { data: null, error: { message: "Bad Gateway" } };

describe("readWithRetry", () => {
  it("returns the first successful result without retrying", async () => {
    const run = vi.fn().mockResolvedValue(ok);
    await expect(readWithRetry(run, { delayMs: 0 })).resolves.toBe(ok);
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("re-runs a read that answered with an error until it succeeds", async () => {
    const run = vi.fn().mockResolvedValueOnce(gateway).mockResolvedValueOnce(gateway).mockResolvedValueOnce(ok);
    await expect(readWithRetry(run, { delayMs: 0 })).resolves.toBe(ok);
    expect(run).toHaveBeenCalledTimes(3);
  });

  it("gives up after the configured attempts and hands back the last error", async () => {
    const run = vi.fn().mockResolvedValue(gateway);
    await expect(readWithRetry(run, { attempts: 2, delayMs: 0 })).resolves.toBe(gateway);
    expect(run).toHaveBeenCalledTimes(2);
  });

  it("retries a read that threw, and rethrows only when every attempt threw", async () => {
    const boom = new Error("fetch failed");
    const recovers = vi.fn().mockRejectedValueOnce(boom).mockResolvedValueOnce(ok);
    await expect(readWithRetry(recovers, { delayMs: 0 })).resolves.toBe(ok);

    const never = vi.fn().mockRejectedValue(boom);
    await expect(readWithRetry(never, { attempts: 2, delayMs: 0 })).rejects.toBe(boom);
    expect(never).toHaveBeenCalledTimes(2);
  });

  it("waits between attempts, a step longer each time", async () => {
    vi.useFakeTimers();
    try {
      const run = vi.fn().mockResolvedValueOnce(gateway).mockResolvedValueOnce(gateway).mockResolvedValueOnce(ok);
      const pending = readWithRetry(run, { delayMs: 100 });
      await vi.advanceTimersByTimeAsync(0);
      expect(run).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(100);
      expect(run).toHaveBeenCalledTimes(2);
      await vi.advanceTimersByTimeAsync(199);
      expect(run).toHaveBeenCalledTimes(2);
      await vi.advanceTimersByTimeAsync(1);
      expect(run).toHaveBeenCalledTimes(3);
      await expect(pending).resolves.toBe(ok);
    } finally {
      vi.useRealTimers();
    }
  });
});
