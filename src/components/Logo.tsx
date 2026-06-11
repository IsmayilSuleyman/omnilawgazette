import Link from "next/link";

/**
 * Faithful CSS recreation of the Omni Law Firm logotype:
 * thin geometric lowercase "omni" on a royal-blue block,
 * letterspaced silver "LAW FIRM" beneath.
 */
export default function Logo({
  size = "md",
  href = "/",
}: {
  size?: "sm" | "md" | "lg";
  href?: string | null;
}) {
  const scale = { sm: "text-[15px]", md: "text-[19px]", lg: "text-[30px]" }[size];

  const mark = (
    <span className={`inline-flex flex-col items-stretch leading-none select-none ${scale}`}>
      <span className="bg-brand text-white rounded-[3px] px-[0.55em] pt-[0.06em] pb-[0.3em] font-jost font-extralight lowercase text-[1.85em] leading-[1.05] tracking-[0.03em] shadow-[0_4px_18px_-6px_rgba(29,66,135,0.9),inset_0_1px_0_rgba(255,255,255,0.18)]">
        omni<span className="sr-only"> law firm</span>
      </span>
      <span
        aria-hidden
        className="text-silver font-jost font-light uppercase text-[0.62em] mt-[0.5em] tracking-[0.62em] [text-indent:0.13em] whitespace-nowrap text-center"
      >
        Law Firm
      </span>
    </span>
  );

  if (!href) return mark;
  return (
    <Link href={href} className="shrink-0 transition-opacity hover:opacity-85">
      {mark}
    </Link>
  );
}
