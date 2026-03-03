export type JmapSession = {
  apiUrl: string;
  downloadUrl: string;
  uploadUrl: string;
  eventSourceUrl: string;
  username: string;
  primaryAccounts: Record<string, string>;
  capabilities: Record<string, unknown>;
  state: string;
};

export type JmapEmailAddress = {
  name: string | null;
  email: string;
};

export type JmapEmailBodyPart = {
  partId: string;
  type: string;
  size: number;
};

export type JmapEmailAttachment = {
  partId: string;
  blobId: string;
  name: string | null;
  type: string;
  size: number;
};

export type JmapEmailBodyValue = {
  value: string;
  isTruncated: boolean;
  isEncodingProblem: boolean;
};

export type JmapEmail = {
  id: string;
  blobId: string;
  threadId: string;
  mailboxIds: Record<string, boolean>;
  keywords: Record<string, boolean>;
  receivedAt: string;
  sentAt: string | null;
  size: number;
  subject: string;
  from: JmapEmailAddress[] | null;
  to: JmapEmailAddress[] | null;
  cc: JmapEmailAddress[] | null;
  bcc: JmapEmailAddress[] | null;
  replyTo: JmapEmailAddress[] | null;
  messageId: string[] | null;
  inReplyTo: string[] | null;
  references: string[] | null;
  hasAttachment: boolean;
  preview: string;
  textBody: JmapEmailBodyPart[];
  htmlBody: JmapEmailBodyPart[];
  attachments: JmapEmailAttachment[];
  bodyValues: Record<string, JmapEmailBodyValue>;
};

export type JmapMailbox = {
  id: string;
  name: string;
  parentId: string | null;
  role: string | null;
  sortOrder: number;
  totalEmails: number;
  unreadEmails: number;
  totalThreads: number;
  unreadThreads: number;
};

export type JmapThread = {
  id: string;
  emailIds: string[];
};

export type JmapIdentity = {
  id: string;
  name: string;
  email: string;
  replyTo: JmapEmailAddress[] | null;
  bcc: JmapEmailAddress[] | null;
  htmlSignature: string;
  textSignature: string;
};

export type JmapEmailSubmission = {
  id: string;
  emailId: string;
  identityId: string;
  envelope: {
    mailFrom: { email: string };
    rcptTo: Array<{ email: string }>;
  };
};

export type JmapMethodCall = [string, Record<string, unknown>, string];

export type JmapMethodResponse = [string, Record<string, unknown>, string];

export type JmapQueryResponse = {
  ids: string[];
  total: number;
  position: number;
};

export type JmapGetResponse<T> = {
  list: T[];
  state: string;
  notFound: string[];
};

export type JmapSetResponse = {
  created: Record<string, Record<string, unknown>> | null;
  updated: Record<string, unknown> | null;
  destroyed: string[] | null;
  notCreated: Record<string, Record<string, unknown>> | null;
  notUpdated: Record<string, Record<string, unknown>> | null;
  notDestroyed: Record<string, Record<string, unknown>> | null;
};
