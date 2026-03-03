export function isGoogleProvider(provider: string | null | undefined) {
  return provider === "google";
}

export function isMicrosoftProvider(provider: string | null | undefined) {
  return provider === "microsoft";
}

export function isFastmailProvider(provider: string | null | undefined) {
  return provider === "fastmail";
}

export function supportsServerFilters(provider: string | null | undefined) {
  return isGoogleProvider(provider) || isMicrosoftProvider(provider);
}
