import type {
  JmapEmail,
  JmapEmailAddress,
  JmapMethodCall,
  JmapGetResponse,
  JmapQueryResponse,
} from "@/utils/fastmail/types";
import {
  EMAIL_PROPERTIES_MINIMAL,
  EMAIL_PROPERTIES_FULL,
  FETCH_BODY_OPTIONS,
  DEFAULT_QUERY_LIMIT,
} from "@/utils/fastmail/constants";
import type { FastmailClient } from "@/utils/fastmail/client";
import type { ParsedMessage, Attachment } from "@/utils/types";

type QueryEmailsOptions = {
  accountId: string;
  filter?: Record<string, unknown>;
  sort?: Array<{ property: string; isAscending: boolean }>;
  limit?: number;
  position?: number;
  collapseThreads?: boolean;
  properties?: readonly string[];
  fetchBody?: boolean;
};

type QueryEmailsResult = {
  emails: JmapEmail[];
  total: number;
  nextPosition?: number;
};

type GetEmailsOptions = {
  accountId: string;
  ids: string[];
  fetchBody?: boolean;
};

type GetEmailOptions = {
  accountId: string;
  id: string;
  fetchBody?: boolean;
};

export async function queryEmails(
  client: FastmailClient,
  options: QueryEmailsOptions,
): Promise<QueryEmailsResult> {
  const {
    accountId,
    filter,
    sort = [{ property: "receivedAt", isAscending: false }],
    limit = DEFAULT_QUERY_LIMIT,
    position,
    collapseThreads,
    properties,
    fetchBody,
  } = options;

  const effectiveProperties =
    properties ??
    (fetchBody ? EMAIL_PROPERTIES_FULL : EMAIL_PROPERTIES_MINIMAL);
  const getArgs: Record<string, unknown> = {
    accountId,
    "#ids": {
      resultOf: "q",
      name: "Email/query",
      path: "/ids",
    },
    properties: effectiveProperties,
  };

  if (fetchBody) {
    Object.assign(getArgs, FETCH_BODY_OPTIONS);
  }

  const methodCalls: JmapMethodCall[] = [
    [
      "Email/query",
      {
        accountId,
        filter,
        sort,
        limit,
        ...(position != null && { position }),
        ...(collapseThreads != null && { collapseThreads }),
      },
      "q",
    ],
    ["Email/get", getArgs, "g"],
  ];

  const responses = await client.request(methodCalls);
  const queryData = responses[0][1] as JmapQueryResponse;
  const getData = responses[1][1] as JmapGetResponse<JmapEmail>;

  const nextPosition = calculateNextPageToken(
    queryData.position,
    queryData.ids.length,
    queryData.total,
  );

  return {
    emails: getData.list,
    total: queryData.total,
    nextPosition: nextPosition != null ? nextPosition : undefined,
  };
}

export async function getEmails(
  client: FastmailClient,
  options: GetEmailsOptions,
): Promise<JmapEmail[]> {
  const { accountId, ids, fetchBody } = options;

  const effectiveProperties = fetchBody
    ? EMAIL_PROPERTIES_FULL
    : EMAIL_PROPERTIES_MINIMAL;
  const getArgs: Record<string, unknown> = {
    accountId,
    ids,
    properties: effectiveProperties,
  };

  if (fetchBody) {
    Object.assign(getArgs, FETCH_BODY_OPTIONS);
  }

  const methodCalls: JmapMethodCall[] = [["Email/get", getArgs, "g"]];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapGetResponse<JmapEmail>;

  return data.list;
}

export async function getEmail(
  client: FastmailClient,
  options: GetEmailOptions,
): Promise<JmapEmail> {
  const emails = await getEmails(client, {
    accountId: options.accountId,
    ids: [options.id],
    fetchBody: options.fetchBody,
  });

  if (emails.length === 0) {
    throw new Error(`Email not found: ${options.id}`);
  }

  return emails[0];
}

export function parseJmapEmail(email: JmapEmail): ParsedMessage {
  const labelIds = Object.keys(email.mailboxIds);
  if (!email.keywords.$seen) labelIds.push("UNREAD");
  if (email.keywords.$flagged) labelIds.push("STARRED");

  const textPlain = extractBodyContent(email, "text/plain");
  const textHtml = extractBodyContent(email, "text/html");

  return {
    id: email.id,
    threadId: email.threadId,
    labelIds,
    snippet: email.preview,
    historyId: "",
    attachments:
      email.attachments && email.attachments.length > 0
        ? email.attachments.map(convertAttachment)
        : undefined,
    inline: [],
    headers: {
      subject: email.subject,
      from: formatAddressList(email.from),
      to: formatAddressList(email.to),
      cc: formatAddressList(email.cc) || undefined,
      bcc: formatAddressList(email.bcc) || undefined,
      date: email.sentAt ?? email.receivedAt,
      "message-id": email.messageId?.[0] ?? "",
      "reply-to": formatAddressList(email.replyTo) || undefined,
      "in-reply-to": email.inReplyTo?.[0],
      references: email.references?.join(" "),
    },
    subject: email.subject,
    date: email.sentAt ?? email.receivedAt,
    internalDate: String(new Date(email.receivedAt).getTime()),
    textPlain: textPlain || undefined,
    textHtml: textHtml || undefined,
  };
}

export function calculateNextPageToken(
  position: number,
  returnedCount: number,
  total: number,
): number | undefined {
  const next = position + returnedCount;
  if (next >= total) return undefined;
  return next;
}

// --- helpers ---

function extractBodyContent(
  email: JmapEmail,
  mimeType: string,
): string | undefined {
  const parts = mimeType === "text/plain" ? email.textBody : email.htmlBody;
  if (!parts || parts.length === 0) return undefined;

  const content = parts
    .map((part) => email.bodyValues?.[part.partId]?.value)
    .filter(Boolean)
    .join("");

  return content || undefined;
}

function convertAttachment(attachment: {
  partId: string;
  blobId: string;
  name: string | null;
  type: string;
  size: number;
}): Attachment {
  return {
    filename: attachment.name ?? "",
    mimeType: attachment.type,
    size: attachment.size,
    attachmentId: attachment.blobId,
    headers: {
      "content-type": attachment.type,
      "content-description": "",
      "content-transfer-encoding": "",
      "content-id": "",
    },
  };
}

function formatAddress(addr: JmapEmailAddress): string {
  if (!addr.name || addr.name === addr.email) return addr.email;
  return `${addr.name} <${addr.email}>`;
}

function formatAddressList(addresses: JmapEmailAddress[] | null): string {
  if (!addresses || addresses.length === 0) return "";
  return addresses.map(formatAddress).join(", ");
}
