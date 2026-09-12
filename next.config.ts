import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // The gazette is mounted under /gazette of İsmayıl Hüquq Bələdçisi, which
  // proxies that path here (Next.js multi-zones). Every route, asset and
  // link of this app therefore lives under /gazette.
  basePath: "/gazette",
  async redirects() {
    // Visiting this deployment's root directly still lands on the library.
    return [{ source: "/", destination: "/gazette", basePath: false, permanent: false }];
  },
  images: {
    remotePatterns: [
      {
        protocol: "https",
        hostname: "*.supabase.co",
        pathname: "/storage/v1/object/public/**",
      },
    ],
  },
};

export default nextConfig;
