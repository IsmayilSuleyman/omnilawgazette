# Omni Law Gazette

Internal weekly legislature digest of **Omni Law Firm**, presented as a library of weekly
issues. Anyone with the link can read; only editors can publish.

## What's inside

- **Library home** — masthead, featured latest issue, searchable/filterable archive grid with
  generated cover images.
- **Reading room** — an advanced in-browser PDF viewer (continuous scroll, thumbnails rail,
  zoom / fit-width / fit-page, in-document text search with highlights, rotate, fullscreen,
  keyboard shortcuts) plus a download button with a download counter.
- **Discussion** — an open comment section per issue (name + comment, no account needed);
  editors can moderate.
- **Admin desk** (`/admin`) — editor sign-in, drag-and-drop weekly publishing (cover image and
  page count are generated automatically from the PDF), issue management (edit/delete), and
  settings (password, editor list).

## Stack

- [Next.js 16](https://nextjs.org) (App Router) + Tailwind CSS v4, dark glassmorphism design
- [react-pdf / pdf.js](https://github.com/wojtekmaj/react-pdf) for the viewer
- [Supabase](https://supabase.com) — Postgres (issues, comments, editor allow-list with RLS),
  Storage (PDFs + covers in the public `gazette` bucket), Auth (editor accounts only)

The Supabase **publishable key** in `src/lib/supabase.ts` is safe to expose; every write is
guarded by row-level-security policies (`is_admin()` checks the signed-in editor's email
against the `admin_emails` table).

## Local development

```bash
npm install        # also copies the pdf.js worker into public/
npm run dev
```

Optional `.env.local` (falls back to the built-in project values):

```bash
NEXT_PUBLIC_SUPABASE_URL=https://<project-ref>.supabase.co
NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY=sb_publishable_...
```

## Weekly publishing routine

1. Open `/admin` and sign in.
2. Drop the week's PDF, check the auto-generated cover/page count, fill title + summary +
   tags, press **Publish issue**.
3. Done — readers see it on the home page within a minute (pages revalidate every 60 s).

## Operations notes

- Editor accounts live in Supabase Auth; the editor allow-list is the `admin_emails` table.
- Public sign-up is not exposed anywhere in the UI; for extra hardening you can disable
  sign-ups entirely in the Supabase dashboard (Auth → Sign In / Up).
- Storage bucket `gazette` is public-read (required for the "anyone with the link" model);
  uploads/deletes require an editor session.
- The site sets `robots: noindex` — it is meant for internal circulation, not search engines.
