export const JMAP_SESSION_URL = "https://api.fastmail.com/jmap/session";

export const WANTED_CAPABILITIES = [
  "urn:ietf:params:jmap:core",
  "urn:ietf:params:jmap:mail",
  "urn:ietf:params:jmap:submission",
] as const;

export const EMAIL_PROPERTIES_MINIMAL = [
  "id",
  "threadId",
  "mailboxIds",
  "keywords",
  "subject",
  "from",
  "to",
  "cc",
  "bcc",
  "replyTo",
  "receivedAt",
  "sentAt",
  "hasAttachment",
  "preview",
  "size",
  "messageId",
  "inReplyTo",
  "references",
] as const;

export const EMAIL_PROPERTIES_FULL = [
  ...EMAIL_PROPERTIES_MINIMAL,
  "textBody",
  "htmlBody",
  "bodyValues",
  "attachments",
] as const;

export const FETCH_BODY_OPTIONS = {
  fetchTextBodyValues: true,
  fetchHTMLBodyValues: true,
  maxBodyValueBytes: 256_000,
} as const;

export const DEFAULT_QUERY_LIMIT = 50;

export const MAX_BULK_EMAILS = 500;
