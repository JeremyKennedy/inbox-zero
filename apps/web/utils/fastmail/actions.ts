import type {
  JmapMethodCall,
  JmapSetResponse,
  JmapQueryResponse,
} from "@/utils/fastmail/types";
import type { FastmailClient } from "@/utils/fastmail/client";
import { MAX_BULK_EMAILS } from "@/utils/fastmail/constants";
import { getThreadEmailIds } from "@/utils/fastmail/thread";

type ArchiveThreadOptions = {
  accountId: string;
  threadId: string;
  inboxId: string;
  archiveId: string;
};

type ArchiveMessageOptions = {
  accountId: string;
  messageId: string;
  inboxId: string;
  archiveId: string;
};

type TrashThreadOptions = {
  accountId: string;
  threadId: string;
  inboxId: string;
  trashId: string;
};

type MarkThreadReadOptions = {
  accountId: string;
  threadId: string;
  read: boolean;
};

type MarkSpamOptions = {
  accountId: string;
  threadId: string;
  inboxId: string;
  junkId: string;
};

type LabelMessageOptions = {
  accountId: string;
  messageId: string;
  mailboxId: string;
};

type RemoveLabelOptions = {
  accountId: string;
  threadId: string;
  mailboxId: string;
};

type RemoveLabelsOptions = {
  accountId: string;
  threadId: string;
  mailboxIds: string[];
};

type MoveThreadToFolderOptions = {
  accountId: string;
  threadId: string;
  targetMailboxId: string;
};

type BulkArchiveFromSendersOptions = {
  accountId: string;
  senders: string[];
  inboxId: string;
  archiveId: string;
};

type BulkTrashFromSendersOptions = {
  accountId: string;
  senders: string[];
  inboxId: string;
  trashId: string;
};

export async function archiveThread(
  client: FastmailClient,
  options: ArchiveThreadOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, inboxId, archiveId } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  const update = buildUpdateForAll(emailIds, {
    [`mailboxIds/${inboxId}`]: null,
    [`mailboxIds/${archiveId}`]: true,
  });

  return emailSet(client, accountId, update);
}

export async function archiveMessage(
  client: FastmailClient,
  options: ArchiveMessageOptions,
): Promise<JmapSetResponse> {
  const { accountId, messageId, inboxId, archiveId } = options;

  const update = {
    [messageId]: {
      [`mailboxIds/${inboxId}`]: null,
      [`mailboxIds/${archiveId}`]: true,
    },
  };

  return emailSet(client, accountId, update);
}

export async function trashThread(
  client: FastmailClient,
  options: TrashThreadOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, inboxId, trashId } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  const update = buildUpdateForAll(emailIds, {
    [`mailboxIds/${inboxId}`]: null,
    [`mailboxIds/${trashId}`]: true,
  });

  return emailSet(client, accountId, update);
}

export async function markThreadRead(
  client: FastmailClient,
  options: MarkThreadReadOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, read } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  const update = buildUpdateForAll(emailIds, {
    "keywords/$seen": read ? true : null,
  });

  return emailSet(client, accountId, update);
}

export async function markSpam(
  client: FastmailClient,
  options: MarkSpamOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, inboxId, junkId } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  const update = buildUpdateForAll(emailIds, {
    [`mailboxIds/${inboxId}`]: null,
    [`mailboxIds/${junkId}`]: true,
  });

  return emailSet(client, accountId, update);
}

export async function labelMessage(
  client: FastmailClient,
  options: LabelMessageOptions,
): Promise<JmapSetResponse> {
  const { accountId, messageId, mailboxId } = options;

  const update = {
    [messageId]: {
      [`mailboxIds/${mailboxId}`]: true,
    },
  };

  return emailSet(client, accountId, update);
}

export async function removeLabel(
  client: FastmailClient,
  options: RemoveLabelOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, mailboxId } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  const update = buildUpdateForAll(emailIds, {
    [`mailboxIds/${mailboxId}`]: null,
  });

  return emailSet(client, accountId, update);
}

export async function removeLabels(
  client: FastmailClient,
  options: RemoveLabelsOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, mailboxIds } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  const patch: Record<string, null> = {};
  for (const mailboxId of mailboxIds) {
    patch[`mailboxIds/${mailboxId}`] = null;
  }

  const update = buildUpdateForAll(emailIds, patch);

  return emailSet(client, accountId, update);
}

export async function moveThreadToFolder(
  client: FastmailClient,
  options: MoveThreadToFolderOptions,
): Promise<JmapSetResponse> {
  const { accountId, threadId, targetMailboxId } = options;

  const emailIds = await getThreadEmailIds(client, { accountId, threadId });
  // moveToFolder is intentionally a full replacement — the target folder becomes the only mailbox
  const update = buildUpdateForAll(emailIds, {
    mailboxIds: { [targetMailboxId]: true },
  });

  return emailSet(client, accountId, update);
}

export async function bulkArchiveFromSenders(
  client: FastmailClient,
  options: BulkArchiveFromSendersOptions,
): Promise<void> {
  const { accountId, senders, inboxId, archiveId } = options;

  for (const sender of senders) {
    const emailIds = await queryEmailIdsBySender(
      client,
      accountId,
      sender,
      inboxId,
    );

    if (emailIds.length === 0) continue;

    const update = buildUpdateForAll(emailIds, {
      [`mailboxIds/${inboxId}`]: null,
      [`mailboxIds/${archiveId}`]: true,
    });

    await emailSet(client, accountId, update);
  }
}

export async function bulkTrashFromSenders(
  client: FastmailClient,
  options: BulkTrashFromSendersOptions,
): Promise<void> {
  const { accountId, senders, inboxId, trashId } = options;

  for (const sender of senders) {
    const emailIds = await queryEmailIdsBySender(
      client,
      accountId,
      sender,
      inboxId,
    );

    if (emailIds.length === 0) continue;

    const update = buildUpdateForAll(emailIds, {
      [`mailboxIds/${inboxId}`]: null,
      [`mailboxIds/${trashId}`]: true,
    });

    await emailSet(client, accountId, update);
  }
}

// --- helpers ---

export function buildUpdateForAll(
  emailIds: string[],
  patch: Record<string, unknown>,
): Record<string, Record<string, unknown>> {
  const update: Record<string, Record<string, unknown>> = {};
  for (const id of emailIds) {
    update[id] = { ...patch };
  }
  return update;
}

async function emailSet(
  client: FastmailClient,
  accountId: string,
  update: Record<string, Record<string, unknown>>,
): Promise<JmapSetResponse> {
  const methodCalls: JmapMethodCall[] = [
    ["Email/set", { accountId, update }, "s"],
  ];

  const responses = await client.request(methodCalls);
  return responses[0][1] as JmapSetResponse;
}

async function queryEmailIdsBySender(
  client: FastmailClient,
  accountId: string,
  sender: string,
  inboxId: string,
): Promise<string[]> {
  const methodCalls: JmapMethodCall[] = [
    [
      "Email/query",
      {
        accountId,
        filter: {
          inMailbox: inboxId,
          from: sender,
        },
        limit: MAX_BULK_EMAILS,
      },
      "q",
    ],
  ];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapQueryResponse;

  return data.ids;
}
