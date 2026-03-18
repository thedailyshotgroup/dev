/** @type {import('next').NextConfig} */
const nextConfig = {
  async headers() {
    return [{
      source: '/api/:path*',
      headers: [
        { key: 'Access-Control-Allow-Methods', value: 'POST, OPTIONS' },
        { key: 'Access-Control-Allow-Headers', value: 'Content-Type' },
      ],
    }];
  },
  async redirects() {
    return [{
      source: '/',
      destination: '/dashboard',
      permanent: false,
    }];
  },
};

module.exports = nextConfig;
