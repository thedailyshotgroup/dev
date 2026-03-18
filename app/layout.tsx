import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'PhotonDoq — Documentos à velocidade da luz',
  description: 'OCR inteligente com IA. Extrai dados de faturas, guias e certificados. Exporta para TOConline, Odoo, PHC, Primavera.',
  icons: { icon: '/favicon.ico' },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pt">
      <body style={{ margin: 0, padding: 0 }}>{children}</body>
    </html>
  );
}
