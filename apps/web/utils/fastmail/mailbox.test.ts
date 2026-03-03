import { describe, it, expect } from "vitest";
import {
  MailboxCache,
  getMailboxByRole,
  getInboxId,
  getArchiveId,
  getTrashId,
  getSentId,
  getDraftsId,
  getJunkId,
  parseMailboxToLabel,
  getLabels,
} from "./mailbox";
import type { JmapMailbox } from "./types";

function makeMailbox(overrides: Partial<JmapMailbox> = {}): JmapMailbox {
  return {
    id: "mb-1",
    name: "Inbox",
    parentId: null,
    role: "inbox",
    sortOrder: 1,
    totalEmails: 100,
    unreadEmails: 5,
    totalThreads: 80,
    unreadThreads: 3,
    ...overrides,
  };
}

const testMailboxes: JmapMailbox[] = [
  makeMailbox({ id: "mb-inbox", name: "Inbox", role: "inbox" }),
  makeMailbox({ id: "mb-archive", name: "Archive", role: "archive" }),
  makeMailbox({ id: "mb-trash", name: "Trash", role: "trash" }),
  makeMailbox({ id: "mb-sent", name: "Sent", role: "sent" }),
  makeMailbox({ id: "mb-drafts", name: "Drafts", role: "drafts" }),
  makeMailbox({ id: "mb-junk", name: "Junk Mail", role: "junk" }),
  makeMailbox({ id: "mb-custom", name: "My Folder", role: null }),
];

describe("MailboxCache", () => {
  const cache = new MailboxCache(testMailboxes);

  it("looks up by id", () => {
    expect(cache.getById("mb-inbox")?.name).toBe("Inbox");
    expect(cache.getById("nonexistent")).toBeUndefined();
  });

  it("looks up by role", () => {
    expect(cache.getByRole("inbox")?.id).toBe("mb-inbox");
    expect(cache.getByRole("archive")?.id).toBe("mb-archive");
    expect(cache.getByRole("nonexistent")).toBeUndefined();
  });

  it("looks up by name (case-insensitive)", () => {
    expect(cache.getByName("inbox")?.id).toBe("mb-inbox");
    expect(cache.getByName("INBOX")?.id).toBe("mb-inbox");
    expect(cache.getByName("My Folder")?.id).toBe("mb-custom");
    expect(cache.getByName("my folder")?.id).toBe("mb-custom");
    expect(cache.getByName("nonexistent")).toBeUndefined();
  });

  it("returns all mailboxes", () => {
    expect(cache.getAll()).toHaveLength(testMailboxes.length);
  });

  it("handles empty mailbox list", () => {
    const empty = new MailboxCache([]);
    expect(empty.getAll()).toHaveLength(0);
    expect(empty.getById("anything")).toBeUndefined();
    expect(empty.getByRole("inbox")).toBeUndefined();
  });

  it("handles mailboxes without roles", () => {
    const noRoles = new MailboxCache([
      makeMailbox({ id: "m1", name: "Custom", role: null }),
    ]);
    expect(noRoles.getByRole("inbox")).toBeUndefined();
    expect(noRoles.getByName("custom")?.id).toBe("m1");
  });
});

describe("role-based getters", () => {
  const cache = new MailboxCache(testMailboxes);

  it("getInboxId returns the inbox mailbox id", () => {
    expect(getInboxId(cache)).toBe("mb-inbox");
  });

  it("getArchiveId returns the archive mailbox id", () => {
    expect(getArchiveId(cache)).toBe("mb-archive");
  });

  it("getTrashId returns the trash mailbox id", () => {
    expect(getTrashId(cache)).toBe("mb-trash");
  });

  it("getSentId returns the sent mailbox id", () => {
    expect(getSentId(cache)).toBe("mb-sent");
  });

  it("getDraftsId returns the drafts mailbox id", () => {
    expect(getDraftsId(cache)).toBe("mb-drafts");
  });

  it("getJunkId returns the junk mailbox id", () => {
    expect(getJunkId(cache)).toBe("mb-junk");
  });

  it("getMailboxByRole throws for missing roles", () => {
    const cacheNoTrash = new MailboxCache([
      makeMailbox({ id: "mb-inbox", role: "inbox" }),
    ]);
    expect(() => getMailboxByRole(cacheNoTrash, "trash")).toThrow(
      'No mailbox found with role "trash"',
    );
  });
});

describe("parseMailboxToLabel", () => {
  it("converts system mailbox to label with system type", () => {
    const label = parseMailboxToLabel(
      makeMailbox({
        id: "mb-inbox",
        name: "Inbox",
        role: "inbox",
        totalThreads: 80,
      }),
    );
    expect(label).toEqual({
      id: "mb-inbox",
      name: "Inbox",
      type: "system",
      threadsTotal: 80,
    });
  });

  it("converts user mailbox to label with user type", () => {
    const label = parseMailboxToLabel(
      makeMailbox({
        id: "mb-custom",
        name: "My Folder",
        role: null,
        totalThreads: 10,
      }),
    );
    expect(label).toEqual({
      id: "mb-custom",
      name: "My Folder",
      type: "user",
      threadsTotal: 10,
    });
  });
});

describe("getLabels", () => {
  it("returns labels for all mailboxes", () => {
    const cache = new MailboxCache(testMailboxes);
    const labels = getLabels(cache);
    expect(labels).toHaveLength(testMailboxes.length);

    const inboxLabel = labels.find((l) => l.name === "Inbox");
    expect(inboxLabel?.type).toBe("system");

    const customLabel = labels.find((l) => l.name === "My Folder");
    expect(customLabel?.type).toBe("user");
  });
});
