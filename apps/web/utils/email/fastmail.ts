import type { FastmailClient } from "@/utils/fastmail/client";
import type { ParsedMessage } from "@/utils/types";
import type {
  EmailFilter,
  EmailLabel,
  EmailProvider,
  EmailSignature,
  EmailThread,
} from "@/utils/email/types";
import type { OutlookFolder } from "@/utils/outlook/folders";
import type { InboxZeroLabel } from "@/utils/label";
import { inboxZeroLabels, PARENT_LABEL } from "@/utils/label";
import type { ThreadsQuery } from "@/app/api/threads/validation";
import type { Logger } from "@/utils/logger";
import {
  type MailboxCache,
  getMailboxes,
  getLabels,
  getFolders,
  getInboxId,
  getArchiveId,
  getTrashId,
  getSentId,
  getDraftsId,
  getJunkId,
} from "@/utils/fastmail/mailbox";
import {
  queryEmails,
  getEmails,
  getEmail,
  parseJmapEmail,
} from "@/utils/fastmail/message";
import { getThread, getThreadEmailIds } from "@/utils/fastmail/thread";
import {
  archiveThread as jmapArchiveThread,
  archiveMessage as jmapArchiveMessage,
  trashThread as jmapTrashThread,
  markThreadRead,
  markSpam as jmapMarkSpam,
  labelMessage as jmapLabelMessage,
  removeLabel,
  removeLabels,
  moveThreadToFolder as jmapMoveThreadToFolder,
  bulkArchiveFromSenders as jmapBulkArchiveFromSenders,
  bulkTrashFromSenders as jmapBulkTrashFromSenders,
} from "@/utils/fastmail/actions";
import {
  sendEmail as jmapSendEmail,
  createDraft as jmapCreateDraft,
  updateDraft as jmapUpdateDraft,
  deleteDraft as jmapDeleteDraft,
  sendDraft as jmapSendDraft,
  getIdentities,
} from "@/utils/fastmail/mail";

export class FastmailProvider implements EmailProvider {
  readonly name = "fastmail" as const;
  private readonly client: FastmailClient;
  private readonly logger: Logger;
  private mailboxCache: MailboxCache | null = null;

  constructor(client: FastmailClient, logger: Logger) {
    this.client = client;
    this.logger = logger;
  }

  toJSON() {
    return { name: "fastmail", type: "fastmail" };
  }

  getAccessToken(): string {
    return this.client.getAccessToken();
  }

  // --- mailbox cache ---

  private async cache(): Promise<MailboxCache> {
    if (!this.mailboxCache) {
      this.mailboxCache = await getMailboxes(this.client);
    }
    return this.mailboxCache;
  }

  // --- labels & folders ---

  async getLabels(): Promise<EmailLabel[]> {
    return getLabels(await this.cache());
  }

  async getLabelById(labelId: string): Promise<EmailLabel | null> {
    const c = await this.cache();
    const mb = c.getById(labelId);
    if (!mb) return null;
    return { id: mb.id, name: mb.name, type: mb.role ? "system" : "user" };
  }

  async getLabelByName(name: string): Promise<EmailLabel | null> {
    const c = await this.cache();
    const mb = c.getByName(name);
    if (!mb) return null;
    return { id: mb.id, name: mb.name, type: mb.role ? "system" : "user" };
  }

  async getFolders(): Promise<OutlookFolder[]> {
    return getFolders(await this.cache());
  }

  async createLabel(name: string): Promise<EmailLabel> {
    const accountId = await this.client.getAccountId();
    const responses = await this.client.request([
      [
        "Mailbox/set",
        { accountId, create: { newLabel: { name, parentId: null } } },
        "cl",
      ],
    ]);
    const data = responses[0][1] as {
      created?: Record<string, { id: string }>;
    };
    const created = data.created?.newLabel;
    if (!created) throw new Error(`Failed to create mailbox "${name}"`);
    this.mailboxCache = null;
    return { id: created.id, name, type: "user" };
  }

  async deleteLabel(labelId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    await this.client.request([
      ["Mailbox/set", { accountId, destroy: [labelId] }, "dl"],
    ]);
    this.mailboxCache = null;
  }

  async getOrCreateInboxZeroLabel(key: InboxZeroLabel): Promise<EmailLabel> {
    const labelDef = inboxZeroLabels[key];

    let parent = await this.getLabelByName(PARENT_LABEL);
    if (!parent) {
      parent = await this.createLabel(PARENT_LABEL);
    }

    const existing = await this.getLabelByName(labelDef.name);
    if (existing) return existing;

    const accountId = await this.client.getAccountId();
    const responses = await this.client.request([
      [
        "Mailbox/set",
        {
          accountId,
          create: {
            izLabel: {
              name: labelDef.name.replace(`${PARENT_LABEL}/`, ""),
              parentId: parent.id,
            },
          },
        },
        "izl",
      ],
    ]);
    const data = responses[0][1] as {
      created?: Record<string, { id: string }>;
    };
    const created = data.created?.izLabel;
    if (!created)
      throw new Error(`Failed to create InboxZero label "${labelDef.name}"`);
    this.mailboxCache = null;
    return { id: created.id, name: labelDef.name, type: "user" };
  }

  // --- threads ---

  async getThreads(folderId?: string): Promise<EmailThread[]> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const inboxId = folderId ?? getInboxId(c);

    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: inboxId },
      collapseThreads: true,
      fetchBody: false,
    });

    return this.groupByThread(result.emails.map(parseJmapEmail));
  }

  async getThread(threadId: string): Promise<EmailThread> {
    const accountId = await this.client.getAccountId();
    const result = await getThread(this.client, { accountId, threadId });
    return {
      id: threadId,
      messages: result.messages,
      snippet: result.messages[0]?.snippet ?? "",
    };
  }

  async getThreadMessages(threadId: string): Promise<ParsedMessage[]> {
    const thread = await this.getThread(threadId);
    return thread.messages;
  }

  async getThreadMessagesInInbox(threadId: string): Promise<ParsedMessage[]> {
    const thread = await this.getThread(threadId);
    const c = await this.cache();
    const inboxId = getInboxId(c);
    return thread.messages.filter((m) => m.labelIds?.includes(inboxId));
  }

  async getThreadsWithParticipant(options: {
    participantEmail: string;
    maxThreads?: number;
  }): Promise<EmailThread[]> {
    const accountId = await this.client.getAccountId();
    const result = await queryEmails(this.client, {
      accountId,
      filter: { from: options.participantEmail },
      limit: options.maxThreads ?? 50,
      collapseThreads: true,
      fetchBody: false,
    });
    return this.groupByThread(result.emails.map(parseJmapEmail));
  }

  async getThreadsWithLabel(options: {
    labelId: string;
    maxResults?: number;
  }): Promise<EmailThread[]> {
    const accountId = await this.client.getAccountId();
    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: options.labelId },
      limit: options.maxResults ?? 50,
      collapseThreads: true,
      fetchBody: false,
    });
    return this.groupByThread(result.emails.map(parseJmapEmail));
  }

  async getThreadsWithQuery(options: {
    query?: ThreadsQuery;
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ threads: EmailThread[]; nextPageToken?: string }> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const filter = buildThreadQueryFilter(c, options.query);

    const result = await queryEmails(this.client, {
      accountId,
      filter,
      limit: options.maxResults ?? 50,
      position: options.pageToken ? Number(options.pageToken) : undefined,
      collapseThreads: true,
      fetchBody: false,
    });

    const threads = this.groupByThread(result.emails.map(parseJmapEmail));
    return {
      threads,
      nextPageToken: result.nextPosition?.toString(),
    };
  }

  async getThreadsFromSenderWithSubject(
    sender: string,
    limit: number,
  ): Promise<Array<{ id: string; snippet: string; subject: string }>> {
    const accountId = await this.client.getAccountId();
    const result = await queryEmails(this.client, {
      accountId,
      filter: { from: sender },
      limit,
      collapseThreads: true,
      fetchBody: false,
    });
    return result.emails.map((e) => {
      const parsed = parseJmapEmail(e);
      return {
        id: parsed.threadId,
        snippet: parsed.snippet ?? "",
        subject: parsed.headers.subject,
      };
    });
  }

  // --- messages ---

  async getMessage(messageId: string): Promise<ParsedMessage> {
    const accountId = await this.client.getAccountId();
    const email = await getEmail(this.client, {
      accountId,
      id: messageId,
      fetchBody: true,
    });
    return parseJmapEmail(email);
  }

  async getMessageByRfc822MessageId(
    rfc822MessageId: string,
  ): Promise<ParsedMessage | null> {
    const accountId = await this.client.getAccountId();
    const result = await queryEmails(this.client, {
      accountId,
      filter: { header: ["Message-ID", rfc822MessageId] },
      limit: 1,
      fetchBody: true,
    });
    if (result.emails.length === 0) return null;
    return parseJmapEmail(result.emails[0]);
  }

  async getSentMessages(maxResults?: number): Promise<ParsedMessage[]> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const sentId = getSentId(c);
    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: sentId },
      limit: maxResults ?? 50,
      fetchBody: false,
    });
    return result.emails.map(parseJmapEmail);
  }

  async getInboxMessages(maxResults?: number): Promise<ParsedMessage[]> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const inboxId = getInboxId(c);
    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: inboxId },
      limit: maxResults ?? 50,
      fetchBody: false,
    });
    return result.emails.map(parseJmapEmail);
  }

  async getMessagesBatch(messageIds: string[]): Promise<ParsedMessage[]> {
    if (messageIds.length === 0) return [];
    const accountId = await this.client.getAccountId();
    const emails = await getEmails(this.client, {
      accountId,
      ids: messageIds,
      fetchBody: true,
    });
    return emails.map(parseJmapEmail);
  }

  async getOriginalMessage(
    originalMessageId: string | undefined,
  ): Promise<ParsedMessage | null> {
    if (!originalMessageId) return null;
    try {
      return await this.getMessage(originalMessageId);
    } catch {
      return null;
    }
  }

  async getLatestMessageFromThreadSnapshot(
    thread: Pick<EmailThread, "id" | "messages">,
  ): Promise<ParsedMessage | null> {
    if (thread.messages.length > 0) {
      return thread.messages[thread.messages.length - 1];
    }
    return this.getLatestMessageInThread(thread.id);
  }

  async getLatestMessageInThread(
    threadId: string,
  ): Promise<ParsedMessage | null> {
    const thread = await this.getThread(threadId);
    if (thread.messages.length === 0) return null;
    return thread.messages[thread.messages.length - 1];
  }

  async getPreviousConversationMessages(
    messageIds: string[],
  ): Promise<ParsedMessage[]> {
    return this.getMessagesBatch(messageIds);
  }

  async getAttachment(
    _messageId: string,
    attachmentId: string,
  ): Promise<{ data: string; size: number }> {
    const session = await this.client.getSession();
    const accountId = await this.client.getAccountId();
    const downloadUrl = session.downloadUrl
      .replace("{accountId}", encodeURIComponent(accountId))
      .replace("{blobId}", encodeURIComponent(attachmentId))
      .replace("{name}", "attachment")
      .replace("{type}", "application/octet-stream");

    const response = await fetch(downloadUrl, {
      headers: { Authorization: `Bearer ${this.client.getAccessToken()}` },
    });
    const buffer = await response.arrayBuffer();
    const base64 = Buffer.from(buffer).toString("base64");
    return { data: base64, size: buffer.byteLength };
  }

  // --- pagination queries ---

  async getMessagesWithPagination(options: {
    query?: string;
    maxResults?: number;
    pageToken?: string;
    before?: Date;
    after?: Date;
    inboxOnly?: boolean;
    unreadOnly?: boolean;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const filter = buildPaginationFilter(c, options);

    const result = await queryEmails(this.client, {
      accountId,
      filter,
      limit: options.maxResults ?? 50,
      position: options.pageToken ? Number(options.pageToken) : undefined,
      fetchBody: true,
    });

    return {
      messages: result.emails.map(parseJmapEmail),
      nextPageToken: result.nextPosition?.toString(),
    };
  }

  async getMessagesWithAttachments(options: {
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    const accountId = await this.client.getAccountId();
    const result = await queryEmails(this.client, {
      accountId,
      filter: { hasAttachment: true },
      limit: options.maxResults ?? 50,
      position: options.pageToken ? Number(options.pageToken) : undefined,
      fetchBody: false,
    });
    return {
      messages: result.emails.map(parseJmapEmail),
      nextPageToken: result.nextPosition?.toString(),
    };
  }

  async getMessagesFromSender(options: {
    senderEmail: string;
    maxResults?: number;
    pageToken?: string;
    before?: Date;
    after?: Date;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    const accountId = await this.client.getAccountId();
    const filter: Record<string, unknown> = { from: options.senderEmail };
    if (options.before) filter.before = options.before.toISOString();
    if (options.after) filter.after = options.after.toISOString();

    const result = await queryEmails(this.client, {
      accountId,
      filter,
      limit: options.maxResults ?? 50,
      position: options.pageToken ? Number(options.pageToken) : undefined,
      fetchBody: false,
    });
    return {
      messages: result.emails.map(parseJmapEmail),
      nextPageToken: result.nextPosition?.toString(),
    };
  }

  async getSentMessageIds(options: {
    maxResults: number;
    after?: Date;
    before?: Date;
  }): Promise<{ id: string; threadId: string }[]> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const sentId = getSentId(c);
    const filter: Record<string, unknown> = { inMailbox: sentId };
    if (options.after) filter.after = options.after.toISOString();
    if (options.before) filter.before = options.before.toISOString();

    const result = await queryEmails(this.client, {
      accountId,
      filter,
      limit: options.maxResults,
      fetchBody: false,
    });
    return result.emails.map((e) => ({ id: e.id, threadId: e.threadId }));
  }

  async getSentThreadsExcluding(options: {
    excludeToEmails?: string[];
    excludeFromEmails?: string[];
    maxResults?: number;
  }): Promise<EmailThread[]> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const sentId = getSentId(c);
    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: sentId },
      limit: options.maxResults ?? 50,
      collapseThreads: true,
      fetchBody: false,
    });

    const excludeTo = new Set(
      options.excludeToEmails?.map((e) => e.toLowerCase()),
    );
    const excludeFrom = new Set(
      options.excludeFromEmails?.map((e) => e.toLowerCase()),
    );

    const messages = result.emails.map(parseJmapEmail).filter((m) => {
      if (excludeTo.size > 0) {
        const to = m.headers.to?.toLowerCase() ?? "";
        for (const email of excludeTo) {
          if (to.includes(email)) return false;
        }
      }
      if (excludeFrom.size > 0) {
        const from = m.headers.from?.toLowerCase() ?? "";
        for (const email of excludeFrom) {
          if (from.includes(email)) return false;
        }
      }
      return true;
    });

    return this.groupByThread(messages);
  }

  // --- actions ---

  async archiveThread(threadId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    await jmapArchiveThread(this.client, {
      accountId,
      threadId,
      inboxId: getInboxId(c),
      archiveId: getArchiveId(c),
    });
  }

  async archiveThreadWithLabel(
    threadId: string,
    _ownerEmail: string,
    labelId?: string,
  ): Promise<void> {
    if (labelId) {
      const emailIds = await getThreadEmailIds(this.client, {
        accountId: await this.client.getAccountId(),
        threadId,
      });
      for (const emailId of emailIds) {
        await jmapLabelMessage(this.client, {
          accountId: await this.client.getAccountId(),
          messageId: emailId,
          mailboxId: labelId,
        });
      }
    }
    await this.archiveThread(threadId);
  }

  async archiveMessage(messageId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    await jmapArchiveMessage(this.client, {
      accountId,
      messageId,
      inboxId: getInboxId(c),
      archiveId: getArchiveId(c),
    });
  }

  async trashThread(threadId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    await jmapTrashThread(this.client, {
      accountId,
      threadId,
      trashId: getTrashId(c),
    });
  }

  async markSpam(threadId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    await jmapMarkSpam(this.client, {
      accountId,
      threadId,
      junkId: getJunkId(c),
    });
  }

  async markRead(threadId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    await markThreadRead(this.client, { accountId, threadId, read: true });
  }

  async markReadThread(threadId: string, read: boolean): Promise<void> {
    const accountId = await this.client.getAccountId();
    await markThreadRead(this.client, { accountId, threadId, read });
  }

  async labelMessage(options: {
    messageId: string;
    labelId: string;
    labelName: string | null;
  }): Promise<{ usedFallback?: boolean; actualLabelId?: string }> {
    const accountId = await this.client.getAccountId();
    let mailboxId = options.labelId;

    if (!mailboxId && options.labelName) {
      const label = await this.getLabelByName(options.labelName);
      if (label) {
        mailboxId = label.id;
      } else {
        const created = await this.createLabel(options.labelName);
        mailboxId = created.id;
        return { usedFallback: true, actualLabelId: created.id };
      }
    }

    await jmapLabelMessage(this.client, {
      accountId,
      messageId: options.messageId,
      mailboxId,
    });

    return {};
  }

  async removeThreadLabel(threadId: string, labelId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    await removeLabel(this.client, {
      accountId,
      threadId,
      mailboxId: labelId,
    });
  }

  async removeThreadLabels(
    threadId: string,
    labelIds: string[],
  ): Promise<void> {
    const accountId = await this.client.getAccountId();
    await removeLabels(this.client, {
      accountId,
      threadId,
      mailboxIds: labelIds,
    });
  }

  async moveThreadToFolder(
    threadId: string,
    _ownerEmail: string,
    folderName: string,
  ): Promise<void> {
    const c = await this.cache();
    const mb = c.getByName(folderName);
    if (!mb) throw new Error(`Folder not found: "${folderName}"`);

    const accountId = await this.client.getAccountId();
    await jmapMoveThreadToFolder(this.client, {
      accountId,
      threadId,
      targetMailboxId: mb.id,
    });
  }

  async getOrCreateFolderIdByName(folderName: string): Promise<string> {
    const c = await this.cache();
    const mb = c.getByName(folderName);
    if (mb) return mb.id;
    const label = await this.createLabel(folderName);
    return label.id;
  }

  // --- bulk actions ---

  async bulkArchiveFromSenders(fromEmails: string[]): Promise<void> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    await jmapBulkArchiveFromSenders(this.client, {
      accountId,
      senders: fromEmails,
      inboxId: getInboxId(c),
      archiveId: getArchiveId(c),
    });
  }

  async bulkTrashFromSenders(fromEmails: string[]): Promise<void> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    await jmapBulkTrashFromSenders(this.client, {
      accountId,
      senders: fromEmails,
      inboxId: getInboxId(c),
      trashId: getTrashId(c),
    });
  }

  // --- drafts ---

  async getDrafts(options?: { maxResults?: number }): Promise<ParsedMessage[]> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const draftsId = getDraftsId(c);
    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: draftsId },
      limit: options?.maxResults ?? 50,
      fetchBody: true,
    });
    return result.emails.map(parseJmapEmail);
  }

  async getDraft(draftId: string): Promise<ParsedMessage | null> {
    try {
      return await this.getMessage(draftId);
    } catch {
      return null;
    }
  }

  async createDraft(params: {
    to: string;
    subject: string;
    messageHtml: string;
    replyToMessageId?: string;
  }): Promise<{ id: string }> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();

    let inReplyTo: string | undefined;
    let references: string | undefined;
    let threadId: string | undefined;
    if (params.replyToMessageId) {
      const original = await this.getMessage(params.replyToMessageId);
      inReplyTo = original.headers["message-id"];
      references = original.headers.references
        ? `${original.headers.references} ${inReplyTo}`
        : inReplyTo;
      threadId = original.threadId;
    }

    return jmapCreateDraft(this.client, {
      accountId,
      draftsMailboxId: getDraftsId(c),
      to: [params.to],
      subject: params.subject,
      htmlBody: params.messageHtml,
      inReplyTo,
      references,
      threadId,
    });
  }

  async updateDraft(
    draftId: string,
    params: { messageHtml?: string; subject?: string },
  ): Promise<void> {
    const accountId = await this.client.getAccountId();
    await jmapUpdateDraft(this.client, {
      accountId,
      draftId,
      htmlBody: params.messageHtml,
      subject: params.subject,
    });
  }

  async deleteDraft(draftId: string): Promise<void> {
    const accountId = await this.client.getAccountId();
    await jmapDeleteDraft(this.client, { accountId, draftId });
  }

  async sendDraft(
    draftId: string,
  ): Promise<{ messageId: string; threadId: string }> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    return jmapSendDraft(this.client, {
      accountId,
      draftId,
      draftsMailboxId: getDraftsId(c),
      sentMailboxId: getSentId(c),
    });
  }

  async draftEmail(
    email: ParsedMessage,
    args: {
      to?: string;
      subject?: string;
      content: string;
      cc?: string;
      bcc?: string;
    },
    _userEmail: string,
  ): Promise<{ draftId: string }> {
    const to = args.to ?? email.headers.from;
    const subject = args.subject ?? `Re: ${email.headers.subject}`;

    const result = await this.createDraft({
      to,
      subject,
      messageHtml: args.content,
      replyToMessageId: email.id,
    });
    return { draftId: result.id };
  }

  // --- sending ---

  async sendEmail(args: {
    to: string;
    cc?: string;
    bcc?: string;
    subject: string;
    messageText: string;
  }): Promise<void> {
    const accountId = await this.client.getAccountId();
    const identities = await getIdentities(this.client, { accountId });
    if (identities.length === 0) throw new Error("No identities available");

    await jmapSendEmail(this.client, {
      accountId,
      to: [args.to],
      cc: args.cc ? args.cc.split(",").map((e) => e.trim()) : undefined,
      bcc: args.bcc ? args.bcc.split(",").map((e) => e.trim()) : undefined,
      subject: args.subject,
      textBody: args.messageText,
      from: identities[0].email,
    });
  }

  async sendEmailWithHtml(body: {
    replyToEmail?: {
      threadId: string;
      headerMessageId: string;
      references?: string;
      messageId?: string;
    };
    to: string;
    from?: string;
    cc?: string;
    bcc?: string;
    replyTo?: string;
    subject: string;
    messageHtml: string;
    attachments?: Array<{
      filename: string;
      content: string;
      contentType: string;
    }>;
  }): Promise<{ messageId: string; threadId: string }> {
    const accountId = await this.client.getAccountId();
    const identities = await getIdentities(this.client, { accountId });
    const from =
      body.from ??
      identities.find((i) => i.email)?.email ??
      identities[0]?.email;
    if (!from) throw new Error("No sender identity available");

    return jmapSendEmail(this.client, {
      accountId,
      to: [body.to],
      cc: body.cc ? body.cc.split(",").map((e) => e.trim()) : undefined,
      bcc: body.bcc ? body.bcc.split(",").map((e) => e.trim()) : undefined,
      subject: body.subject,
      htmlBody: body.messageHtml,
      from,
      replyTo: body.replyTo,
      inReplyTo: body.replyToEmail?.headerMessageId,
      references: body.replyToEmail?.references,
      threadId: body.replyToEmail?.threadId,
      attachments: body.attachments?.map((a) => ({
        name: a.filename,
        contentType: a.contentType,
        content: a.content,
      })),
    });
  }

  async replyToEmail(
    email: ParsedMessage,
    content: string,
    options?: { replyTo?: string; from?: string },
  ): Promise<void> {
    await this.sendEmailWithHtml({
      replyToEmail: {
        threadId: email.threadId,
        headerMessageId: email.headers["message-id"] ?? "",
        references: email.headers.references,
        messageId: email.id,
      },
      to: email.headers.from,
      from: options?.from,
      replyTo: options?.replyTo,
      subject: `Re: ${email.headers.subject}`,
      messageHtml: content,
    });
  }

  async forwardEmail(
    email: ParsedMessage,
    args: { to: string; cc?: string; bcc?: string; content?: string },
  ): Promise<void> {
    const originalBody = email.textHtml ?? email.textPlain ?? "";
    const forwardBody = args.content
      ? `${args.content}<br/><br/>---------- Forwarded message ----------<br/>${originalBody}`
      : originalBody;

    await this.sendEmailWithHtml({
      to: args.to,
      cc: args.cc,
      bcc: args.bcc,
      subject: `Fwd: ${email.headers.subject}`,
      messageHtml: forwardBody,
    });
  }

  // --- metadata queries ---

  async checkIfReplySent(senderEmail: string): Promise<boolean> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const sentId = getSentId(c);
    const result = await queryEmails(this.client, {
      accountId,
      filter: { inMailbox: sentId, to: senderEmail },
      limit: 1,
      fetchBody: false,
    });
    return result.emails.length > 0;
  }

  async countReceivedMessages(
    senderEmail: string,
    threshold: number,
  ): Promise<number> {
    const accountId = await this.client.getAccountId();
    const result = await queryEmails(this.client, {
      accountId,
      filter: { from: senderEmail },
      limit: threshold,
      fetchBody: false,
    });
    return result.total;
  }

  async hasPreviousCommunicationsWithSenderOrDomain(options: {
    from: string;
    date: Date;
    messageId: string;
  }): Promise<boolean> {
    const accountId = await this.client.getAccountId();
    const c = await this.cache();
    const sentId = getSentId(c);
    const result = await queryEmails(this.client, {
      accountId,
      filter: {
        inMailbox: sentId,
        to: options.from,
        before: options.date.toISOString(),
      },
      limit: 1,
      fetchBody: false,
    });
    return result.emails.length > 0;
  }

  isSentMessage(message: ParsedMessage): boolean {
    const c = this.mailboxCache;
    if (!c) return false;
    const sentMb = c.getByRole("sent");
    if (!sentMb) return false;
    return message.labelIds?.includes(sentMb.id) ?? false;
  }

  isReplyInThread(message: ParsedMessage): boolean {
    return !!message.headers["in-reply-to"];
  }

  // --- signatures ---

  async getSignatures(): Promise<EmailSignature[]> {
    try {
      const accountId = await this.client.getAccountId();
      const identities = await getIdentities(this.client, { accountId });
      return identities
        .filter((i) => i.htmlSignature || i.textSignature)
        .map((i, index) => ({
          email: i.email,
          signature: i.htmlSignature || i.textSignature,
          isDefault: index === 0,
          displayName: i.name || undefined,
        }));
    } catch {
      return [];
    }
  }

  // --- stats ---

  async getInboxStats(): Promise<{ total: number; unread: number }> {
    const c = await this.cache();
    const inbox = c.getByRole("inbox");
    if (!inbox) return { total: 0, unread: 0 };
    return { total: inbox.totalEmails, unread: inbox.unreadEmails };
  }

  // --- filters (not supported in JMAP — return stubs) ---

  async getFiltersList(): Promise<EmailFilter[]> {
    return [];
  }

  async createFilter(): Promise<{ status: number }> {
    this.logger.warn("Filters not supported for Fastmail JMAP provider");
    return { status: 501 };
  }

  async deleteFilter(): Promise<{ status: number }> {
    this.logger.warn("Filters not supported for Fastmail JMAP provider");
    return { status: 501 };
  }

  async createAutoArchiveFilter(): Promise<{ status: number }> {
    this.logger.warn("Filters not supported for Fastmail JMAP provider");
    return { status: 501 };
  }

  // --- webhooks (stub — polling handled separately) ---

  async processHistory(): Promise<void> {
    // JMAP push/changes handled by polling route, not processHistory
  }

  async watchEmails(): Promise<{
    expirationDate: Date;
    subscriptionId?: string;
  } | null> {
    return null;
  }

  async unwatchEmails(): Promise<void> {
    // No-op for Fastmail
  }

  // --- unsubscribe ---

  async blockUnsubscribedEmail(messageId: string): Promise<void> {
    await this.archiveMessage(messageId);
  }

  // --- helpers ---

  private groupByThread(messages: ParsedMessage[]): EmailThread[] {
    const threadMap = new Map<string, ParsedMessage[]>();
    for (const msg of messages) {
      const existing = threadMap.get(msg.threadId);
      if (existing) {
        existing.push(msg);
      } else {
        threadMap.set(msg.threadId, [msg]);
      }
    }
    return Array.from(threadMap.entries()).map(([id, msgs]) => ({
      id,
      messages: msgs,
      snippet: msgs[0]?.snippet ?? "",
    }));
  }
}

// --- module-level helpers ---

function buildThreadQueryFilter(
  cache: MailboxCache,
  query?: ThreadsQuery,
): Record<string, unknown> {
  const conditions: Record<string, unknown>[] = [];

  if (query?.fromEmail) {
    conditions.push({ from: query.fromEmail });
  }
  if (query?.labelId) {
    conditions.push({ inMailbox: query.labelId });
  }
  if (query?.after) {
    conditions.push({ after: query.after.toISOString() });
  }
  if (query?.before) {
    conditions.push({ before: query.before.toISOString() });
  }
  if (query?.isUnread) {
    conditions.push({ notKeyword: "$seen" });
  }

  if (conditions.length === 0) {
    const inbox = cache.getByRole("inbox");
    return inbox ? { inMailbox: inbox.id } : {};
  }
  if (conditions.length === 1) return conditions[0];
  return { operator: "AND", conditions };
}

function buildPaginationFilter(
  cache: MailboxCache,
  options: {
    query?: string;
    inboxOnly?: boolean;
    unreadOnly?: boolean;
    before?: Date;
    after?: Date;
  },
): Record<string, unknown> {
  const conditions: Record<string, unknown>[] = [];

  if (options.inboxOnly) {
    const inbox = cache.getByRole("inbox");
    if (inbox) conditions.push({ inMailbox: inbox.id });
  }
  if (options.query) {
    conditions.push({ text: options.query });
  }
  if (options.unreadOnly) {
    conditions.push({ notKeyword: "$seen" });
  }
  if (options.before) {
    conditions.push({ before: options.before.toISOString() });
  }
  if (options.after) {
    conditions.push({ after: options.after.toISOString() });
  }

  if (conditions.length === 0) return {};
  if (conditions.length === 1) return conditions[0];
  return { operator: "AND", conditions };
}
