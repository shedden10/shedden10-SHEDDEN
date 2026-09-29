import { defineMiddleware } from 'astro:middleware';
import { env } from 'cloudflare:workers';
import { bearerToken, safeEqual, verifySession, SESSION_COOKIE } from './lib/auth';

const PUBLIC_PATHS = new Set(['/login', '/api/login']);

export const onRequest = defineMiddleware(async (context, next) => {
  const { request, locals, url, redirect, cookies } = context;
  const path = url.pathname;

  // Collector endpoints authenticate with a bearer token, bypassing the UI gate.
  if (path.startsWith('/api/collector/')) {
    // Trim so a stray newline/space pasted into the secret (common when the value
    // is piped in from a Windows shell) doesn't break the comparison.
    const expected = (env.INGEST_TOKEN ?? '').trim();
    if (!expected) {
      return new Response('INGEST_TOKEN is not configured', { status: 503 });
    }
    if (!safeEqual(bearerToken(request), expected)) {
      return new Response(
        'Unauthorized: the bearer token does not match the Worker\'s INGEST_TOKEN secret',
        { status: 401 }
      );
    }
    return next();
  }

  // Optional dashboard password gate.
  const appPassword = env.APP_PASSWORD;
  locals.isAuthed = true;
  if (appPassword) {
    const ok = await verifySession(cookies.get(SESSION_COOKIE)?.value, appPassword);
    locals.isAuthed = ok;
    if (!ok && !PUBLIC_PATHS.has(path) && !path.startsWith('/favicon')) {
      if (path.startsWith('/api/')) return new Response('Unauthorized', { status: 401 });
      return redirect('/login');
    }
  }

  return next();
});
