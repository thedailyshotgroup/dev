'use client';

import dynamic from 'next/dynamic';

// Dynamic import para evitar SSR issues com localStorage/window
const PhotonDoq = dynamic(() => import('@/components/PhotonDoq'), {
  ssr: false,
  loading: () => (
    <div style={{
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      height: '100vh', background: '#0c0818', color: '#8b82a8',
      fontFamily: "'IBM Plex Sans', system-ui, sans-serif",
      gap: 8
    }}>
      <svg width={20} height={20} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ animation: 'spin 1s linear infinite' }}>
        <path d="M21 12a9 9 0 1 1-6.219-8.56" />
      </svg>
      A carregar PhotonDoq...
      <style>{`@keyframes spin { to { transform: rotate(360deg) } }`}</style>
    </div>
  ),
});

export default function DashboardPage() {
  return <PhotonDoq />;
}