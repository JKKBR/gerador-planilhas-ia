export type ApiErrorPayload = {
  error?: unknown;
  details?: unknown;
};

function normalizeDetails(
  details: unknown
): string[] {
  if (!Array.isArray(details)) {
    return [];
  }

  return details
    .filter(
      (detail): detail is string =>
        typeof detail === "string"
    )
    .map((detail) => detail.trim())
    .filter(Boolean)
    .slice(0, 10);
}

export function readApiError(
  payload: unknown,
  status: number
) {
  const data =
    typeof payload === "object" &&
    payload !== null
      ? (payload as ApiErrorPayload)
      : {};

  const message =
    typeof data.error === "string" &&
    data.error.trim()
      ? data.error.trim()
      : `Não foi possível gerar a planilha. HTTP ${status}.`;

  return {
    message,
    details: normalizeDetails(data.details),
  };
}

export function formatApiErrorMessage(
  payload: unknown,
  status: number
) {
  const error = readApiError(
    payload,
    status
  );

  if (error.details.length === 0) {
    return error.message;
  }

  return [
    error.message,
    ...error.details.map(
      (detail) => `• ${detail}`
    ),
  ].join("\n");
}
