import Link from "next/link";

/**
 * Omni Law Firm logo. The official asset lives at /public/omni-logo.png.
 */
export default function Logo({
  size = "md",
  href = "/",
}: {
  size?: "sm" | "md" | "lg";
  href?: string | null;
}) {
  const heights = { sm: 44, md: 58, lg: 90 };

  const mark = (
    // eslint-disable-next-line @next/next/no-img-element
    <img
      src="/omni-logo.png"
      alt="Omni Law Firm"
      style={{ height: heights[size], width: "auto" }}
      className="select-none object-contain"
    />
  );

  if (!href) return mark;
  return (
    <Link
      href={href}
      aria-label="Omni Law Firm — home"
      className="shrink-0 transition-opacity hover:opacity-85"
    >
      {mark}
    </Link>
  );
}
