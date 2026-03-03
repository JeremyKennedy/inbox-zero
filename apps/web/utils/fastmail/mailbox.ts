import type {
  JmapMailbox,
  JmapMethodCall,
  JmapGetResponse,
} from "@/utils/fastmail/types";
import type { FastmailClient } from "@/utils/fastmail/client";
import type { EmailLabel } from "@/utils/email/types";
import type { OutlookFolder } from "@/utils/outlook/folders";

export class MailboxCache {
  private readonly byId: Map<string, JmapMailbox>;
  private readonly byRole: Map<string, JmapMailbox>;
  private readonly byName: Map<string, JmapMailbox>;

  constructor(mailboxes: JmapMailbox[]) {
    this.byId = new Map();
    this.byRole = new Map();
    this.byName = new Map();

    for (const mailbox of mailboxes) {
      this.byId.set(mailbox.id, mailbox);
      if (mailbox.role) {
        this.byRole.set(mailbox.role, mailbox);
      }
      this.byName.set(mailbox.name.toLowerCase(), mailbox);
    }
  }

  getById(id: string): JmapMailbox | undefined {
    return this.byId.get(id);
  }

  getByRole(role: string): JmapMailbox | undefined {
    return this.byRole.get(role);
  }

  getByName(name: string): JmapMailbox | undefined {
    return this.byName.get(name.toLowerCase());
  }

  getAll(): JmapMailbox[] {
    return Array.from(this.byId.values());
  }
}

export async function getMailboxes(
  client: FastmailClient,
): Promise<MailboxCache> {
  const accountId = await client.getAccountId();

  const methodCalls: JmapMethodCall[] = [["Mailbox/get", { accountId }, "mb"]];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapGetResponse<JmapMailbox>;

  return new MailboxCache(data.list);
}

export function getMailboxByRole(
  cache: MailboxCache,
  role: string,
): JmapMailbox {
  const mailbox = cache.getByRole(role);
  if (!mailbox) {
    throw new Error(`No mailbox found with role "${role}"`);
  }
  return mailbox;
}

export function getInboxId(cache: MailboxCache): string {
  return getMailboxByRole(cache, "inbox").id;
}

export function getArchiveId(cache: MailboxCache): string {
  return getMailboxByRole(cache, "archive").id;
}

export function getTrashId(cache: MailboxCache): string {
  return getMailboxByRole(cache, "trash").id;
}

export function getSentId(cache: MailboxCache): string {
  return getMailboxByRole(cache, "sent").id;
}

export function getDraftsId(cache: MailboxCache): string {
  return getMailboxByRole(cache, "drafts").id;
}

export function getJunkId(cache: MailboxCache): string {
  return getMailboxByRole(cache, "junk").id;
}

export function parseMailboxToLabel(mailbox: JmapMailbox): EmailLabel {
  return {
    id: mailbox.id,
    name: mailbox.name,
    type: mailbox.role ? "system" : "user",
    threadsTotal: mailbox.totalThreads,
  };
}

export function getLabels(cache: MailboxCache): EmailLabel[] {
  return cache.getAll().map(parseMailboxToLabel);
}

export function getFolders(cache: MailboxCache): OutlookFolder[] {
  return cache.getAll().map(convertMailboxToFolder);
}

// --- helpers ---

function convertMailboxToFolder(mailbox: JmapMailbox): OutlookFolder {
  return {
    id: mailbox.id,
    displayName: mailbox.name,
    childFolders: [],
    childFolderCount: 0,
  };
}
