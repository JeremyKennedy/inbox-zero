import type { FastmailClient } from "@/utils/fastmail/client";
import type {
  JmapIdentity,
  JmapEmailAddress,
  JmapGetResponse,
  JmapSetResponse,
  JmapMethodCall,
} from "@/utils/fastmail/types";

const SUBMISSION_CAPABILITY = "urn:ietf:params:jmap:submission";

type SendAttachment = {
  name: string;
  contentType: string;
  content: string; // base64-encoded
};

type UploadedBlob = {
  blobId: string;
  name: string;
  contentType: string;
};

type SendEmailOptions = {
  accountId: string;
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  textBody?: string;
  htmlBody?: string;
  from: string;
  replyTo?: string;
  inReplyTo?: string;
  references?: string;
  threadId?: string;
  attachments?: SendAttachment[];
};

type SendEmailResult = {
  messageId: string;
  threadId: string;
};

type CreateDraftOptions = {
  accountId: string;
  draftsMailboxId: string;
  to: string[];
  subject: string;
  htmlBody: string;
  inReplyTo?: string;
  references?: string;
  threadId?: string;
};

type UpdateDraftOptions = {
  accountId: string;
  draftId: string;
  htmlBody?: string;
  subject?: string;
};

type DeleteDraftOptions = {
  accountId: string;
  draftId: string;
};

type SendDraftOptions = {
  accountId: string;
  draftId: string;
  draftsMailboxId: string;
  sentMailboxId: string;
};

export async function getIdentities(
  client: FastmailClient,
  { accountId }: { accountId: string },
): Promise<JmapIdentity[]> {
  const methodCalls: JmapMethodCall[] = [["Identity/get", { accountId }, "id"]];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapGetResponse<JmapIdentity>;

  return data.list;
}

export async function sendEmail(
  client: FastmailClient,
  options: SendEmailOptions,
): Promise<SendEmailResult> {
  requireSubmissionCapability(client);

  const {
    accountId,
    to,
    cc,
    bcc,
    subject,
    textBody,
    htmlBody,
    from,
    replyTo,
    inReplyTo,
    references,
    threadId,
    attachments,
  } = options;

  const identities = await getIdentities(client, { accountId });
  const identity = identities.find((i) => i.email === from);
  if (!identity) {
    throw new Error(
      `No identity found for sender address "${from}". Available: ${identities.map((i) => i.email).join(", ")}`,
    );
  }

  let attachmentBlobs: UploadedBlob[] | undefined;
  if (attachments?.length) {
    attachmentBlobs = await uploadSendAttachments(client, attachments);
  }

  const sentMailboxId = await getSentMailboxId(client, accountId);

  const emailObject = buildEmailObject({
    mailboxId: sentMailboxId,
    from,
    to,
    cc,
    bcc,
    replyTo,
    subject,
    textBody,
    htmlBody,
    inReplyTo,
    references,
    threadId,
    attachmentBlobs,
    keywords: { $seen: true },
  });

  const methodCalls: JmapMethodCall[] = [
    [
      "Email/set",
      {
        accountId,
        create: { emailRef: emailObject },
      },
      "emailSet",
    ],
    [
      "EmailSubmission/set",
      {
        accountId,
        create: {
          sendRef: {
            emailId: "#emailRef",
            identityId: identity.id,
          },
        },
        onSuccessUpdateEmail: {
          "#sendRef": {
            "keywords/$draft": null,
          },
        },
      },
      "submit",
    ],
  ];

  const responses = await client.request(methodCalls);
  const emailSetData = responses[0][1] as JmapSetResponse;

  if (emailSetData.notCreated?.emailRef) {
    throw new Error(
      `Failed to create email: ${JSON.stringify(emailSetData.notCreated.emailRef)}`,
    );
  }

  const createdEmail = emailSetData.created?.emailRef;
  if (!createdEmail) {
    throw new Error("Email/set returned no created email");
  }

  const submissionData = responses[1][1] as JmapSetResponse;
  if (submissionData.notCreated?.sendRef) {
    throw new Error(
      `Failed to submit email: ${JSON.stringify(submissionData.notCreated.sendRef)}`,
    );
  }

  return {
    messageId: createdEmail.id as string,
    threadId: createdEmail.threadId as string,
  };
}

export async function createDraft(
  client: FastmailClient,
  options: CreateDraftOptions,
): Promise<{ id: string }> {
  const {
    accountId,
    draftsMailboxId,
    to,
    subject,
    htmlBody,
    inReplyTo,
    references,
    threadId,
  } = options;

  const emailObject = buildEmailObject({
    mailboxId: draftsMailboxId,
    from: undefined,
    to,
    subject,
    htmlBody,
    inReplyTo,
    references,
    threadId,
    keywords: { $seen: true, $draft: true },
  });

  const methodCalls: JmapMethodCall[] = [
    [
      "Email/set",
      {
        accountId,
        create: { draftRef: emailObject },
      },
      "draftSet",
    ],
  ];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapSetResponse;

  if (data.notCreated?.draftRef) {
    throw new Error(
      `Failed to create draft: ${JSON.stringify(data.notCreated.draftRef)}`,
    );
  }

  const created = data.created?.draftRef;
  if (!created) {
    throw new Error("Email/set returned no created draft");
  }

  return { id: created.id as string };
}

export async function updateDraft(
  client: FastmailClient,
  options: UpdateDraftOptions,
): Promise<void> {
  const { accountId, draftId, htmlBody, subject } = options;

  const updates: Record<string, unknown> = {};
  if (subject != null) {
    updates.subject = subject;
  }
  if (htmlBody != null) {
    updates.htmlBody = [{ partId: "html", type: "text/html" }];
    updates["bodyValues/html"] = {
      value: htmlBody,
      isEncodingProblem: false,
      isTruncated: false,
    };
  }

  if (Object.keys(updates).length === 0) return;

  const methodCalls: JmapMethodCall[] = [
    [
      "Email/set",
      {
        accountId,
        update: { [draftId]: updates },
      },
      "updateDraft",
    ],
  ];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapSetResponse;

  if (data.notUpdated?.[draftId]) {
    throw new Error(
      `Failed to update draft: ${JSON.stringify(data.notUpdated[draftId])}`,
    );
  }
}

export async function deleteDraft(
  client: FastmailClient,
  options: DeleteDraftOptions,
): Promise<void> {
  const { accountId, draftId } = options;

  const methodCalls: JmapMethodCall[] = [
    [
      "Email/set",
      {
        accountId,
        destroy: [draftId],
      },
      "deleteDraft",
    ],
  ];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapSetResponse;

  if (data.notDestroyed?.[draftId]) {
    throw new Error(
      `Failed to delete draft: ${JSON.stringify(data.notDestroyed[draftId])}`,
    );
  }
}

export async function sendDraft(
  client: FastmailClient,
  options: SendDraftOptions,
): Promise<SendEmailResult> {
  requireSubmissionCapability(client);

  const { accountId, draftId, draftsMailboxId, sentMailboxId } = options;

  const identities = await getIdentities(client, { accountId });
  if (identities.length === 0) {
    throw new Error("No identities available for sending");
  }

  const methodCalls: JmapMethodCall[] = [
    [
      "Email/set",
      {
        accountId,
        update: {
          [draftId]: {
            "keywords/$draft": null,
            [`mailboxIds/${draftsMailboxId}`]: null,
            [`mailboxIds/${sentMailboxId}`]: true,
          },
        },
      },
      "moveEmail",
    ],
    [
      "EmailSubmission/set",
      {
        accountId,
        create: {
          sendRef: {
            emailId: draftId,
            identityId: identities[0].id,
          },
        },
      },
      "submit",
    ],
    // EmailSubmission doesn't return threadId — fetch it from the email
    [
      "Email/get",
      { accountId, ids: [draftId], properties: ["threadId"] },
      "getThread",
    ],
  ];

  const responses = await client.request(methodCalls);

  const moveData = responses[0][1] as JmapSetResponse;
  if (moveData.notUpdated?.[draftId]) {
    throw new Error(
      `Failed to prepare draft for sending: ${JSON.stringify(moveData.notUpdated[draftId])}`,
    );
  }

  const submitData = responses[1][1] as JmapSetResponse;
  if (submitData.notCreated?.sendRef) {
    throw new Error(
      `Failed to submit draft: ${JSON.stringify(submitData.notCreated.sendRef)}`,
    );
  }

  const emailData = responses[2][1] as JmapGetResponse<{
    id: string;
    threadId: string;
  }>;
  const email = emailData.list[0];

  return {
    messageId: draftId,
    threadId: email?.threadId ?? draftId,
  };
}

// --- helpers ---

function requireSubmissionCapability(client: FastmailClient): void {
  if (!client.hasCapability(SUBMISSION_CAPABILITY)) {
    throw new Error(
      "EmailSubmission is not available. The Fastmail API token may be read-only " +
        "or missing the urn:ietf:params:jmap:submission capability. " +
        "Ensure the token has send permissions.",
    );
  }
}

function toJmapAddresses(emails: string[]): JmapEmailAddress[] {
  return emails.map((email) => ({ name: null, email }));
}

type BuildEmailObjectOptions = {
  mailboxId: string;
  from?: string;
  to: string[];
  cc?: string[];
  bcc?: string[];
  replyTo?: string;
  subject: string;
  textBody?: string;
  htmlBody?: string;
  inReplyTo?: string;
  references?: string;
  threadId?: string;
  attachmentBlobs?: UploadedBlob[];
  keywords: Record<string, boolean>;
};

function buildEmailObject(
  options: BuildEmailObjectOptions,
): Record<string, unknown> {
  const {
    mailboxId,
    from,
    to,
    cc,
    bcc,
    replyTo,
    subject,
    textBody,
    htmlBody,
    inReplyTo,
    references,
    threadId,
    attachmentBlobs,
    keywords,
  } = options;

  const bodyValues: Record<
    string,
    { value: string; isEncodingProblem: boolean; isTruncated: boolean }
  > = {};

  if (textBody) {
    bodyValues.text = {
      value: textBody,
      isEncodingProblem: false,
      isTruncated: false,
    };
  }
  if (htmlBody) {
    bodyValues.html = {
      value: htmlBody,
      isEncodingProblem: false,
      isTruncated: false,
    };
  }

  return {
    mailboxIds: { [mailboxId]: true },
    ...(from ? { from: [{ name: null, email: from }] } : {}),
    to: toJmapAddresses(to),
    ...(cc?.length ? { cc: toJmapAddresses(cc) } : {}),
    ...(bcc?.length ? { bcc: toJmapAddresses(bcc) } : {}),
    ...(replyTo ? { replyTo: [{ name: null, email: replyTo }] } : {}),
    subject,
    ...(textBody ? { textBody: [{ partId: "text", type: "text/plain" }] } : {}),
    ...(htmlBody ? { htmlBody: [{ partId: "html", type: "text/html" }] } : {}),
    bodyValues,
    keywords,
    ...(inReplyTo ? { inReplyTo: [inReplyTo] } : {}),
    ...(references ? { references: references.split(" ") } : {}),
    ...(threadId ? { threadId } : {}),
    ...(attachmentBlobs?.length
      ? {
          attachments: attachmentBlobs.map((b) => ({
            blobId: b.blobId,
            name: b.name,
            type: b.contentType,
          })),
        }
      : {}),
  };
}

async function uploadSendAttachments(
  client: FastmailClient,
  attachments: SendAttachment[],
): Promise<UploadedBlob[]> {
  const results: UploadedBlob[] = [];

  for (const attachment of attachments) {
    const buffer = base64ToArrayBuffer(attachment.content);
    const blobId = await client.uploadBlob(buffer, attachment.contentType);
    results.push({
      blobId,
      name: attachment.name,
      contentType: attachment.contentType,
    });
  }

  return results;
}

async function getSentMailboxId(
  client: FastmailClient,
  accountId: string,
): Promise<string> {
  const responses = await client.request([
    ["Mailbox/get", { accountId }, "mb"],
  ]);
  const data = responses[0][1] as JmapGetResponse<{
    id: string;
    role: string | null;
  }>;
  const sent = data.list.find((m) => m.role === "sent");
  if (!sent) {
    throw new Error('No mailbox with role "sent" found');
  }
  return sent.id;
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binaryString = atob(base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
}
