import { NextRequest, NextResponse } from 'next/server';

// ═══════════════════════════════════════
// PhotonDoq — Proxy para Anthropic API
// ═══════════════════════════════════════
// Este ficheiro protege a chave API.
// O frontend chama /api/ocr em vez de api.anthropic.com
// A chave nunca sai do servidor.

const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const DEFAULT_MODEL = 'claude-sonnet-4-20250514';
const DEFAULT_MAX_TOKENS = 4096;

export async function POST(req: NextRequest) {
  try {
    // Verificar que temos a chave
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      console.error('ANTHROPIC_API_KEY não configurada');
      return NextResponse.json(
        { error: 'Serviço OCR não configurado' },
        { status: 503 }
      );
    }

    const body = await req.json();

    // Validação do request
    if (!body.messages || !Array.isArray(body.messages) || body.messages.length === 0) {
      return NextResponse.json(
        { error: 'Request inválido: messages em falta' },
        { status: 400 }
      );
    }

    // Limitar tamanho do request (20MB base64 = ~15MB ficheiro original)
    const bodyStr = JSON.stringify(body);
    if (bodyStr.length > 20 * 1024 * 1024) {
      return NextResponse.json(
        { error: 'Documento demasiado grande (máx. 15MB)' },
        { status: 413 }
      );
    }

    // Fazer o request à Anthropic
    const response = await fetch(ANTHROPIC_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: body.model || DEFAULT_MODEL,
        max_tokens: Math.min(body.max_tokens || DEFAULT_MAX_TOKENS, 8192),
        messages: body.messages,
      }),
    });

    // Tratar erros da Anthropic
    if (!response.ok) {
      const errorText = await response.text();
      console.error(`Anthropic API error [${response.status}]:`, errorText);

      if (response.status === 429) {
        return NextResponse.json(
          { error: 'Demasiados pedidos. Aguarda um momento.' },
          { status: 429 }
        );
      }

      if (response.status === 401) {
        return NextResponse.json(
          { error: 'Erro de configuração do serviço OCR' },
          { status: 503 }
        );
      }

      return NextResponse.json(
        { error: 'Erro no processamento OCR' },
        { status: 502 }
      );
    }

    const data = await response.json();
    return NextResponse.json(data);

  } catch (error) {
    console.error('OCR proxy error:', error);
    return NextResponse.json(
      { error: 'Erro interno do servidor' },
      { status: 500 }
    );
  }
}

// CORS preflight
export async function OPTIONS() {
  return new NextResponse(null, {
    status: 200,
    headers: {
      'Access-Control-Allow-Origin': process.env.NEXT_PUBLIC_APP_URL || '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    },
  });
}
