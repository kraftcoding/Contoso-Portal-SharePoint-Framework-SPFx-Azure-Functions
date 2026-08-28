export function serializeError(e: unknown): string {
  if (e instanceof Error) {
    const header = `[${e.name}] ${e.message}`;
    return e.stack ? header + '\n\n' + e.stack : header;
  }
  if (typeof e === 'string') return e;
  try {
    return JSON.stringify(e, null, 2);
  } catch {
    return String(e);
  }
}
