import { describe, it, expect } from "vitest";
import { buildThreadQueryFilter, buildPaginationFilter } from "./fastmail";
import { MailboxCache } from "@/utils/fastmail/mailbox";
import type { JmapMailbox } from "@/utils/fastmail/types";

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

const cache = new MailboxCache([
  makeMailbox({ id: "mb-inbox", name: "Inbox", role: "inbox" }),
  makeMailbox({ id: "mb-sent", name: "Sent", role: "sent" }),
  makeMailbox({ id: "mb-archive", name: "Archive", role: "archive" }),
]);

describe("buildThreadQueryFilter", () => {
  it("defaults to inbox when no query provided", () => {
    expect(buildThreadQueryFilter(cache)).toEqual({
      inMailbox: "mb-inbox",
    });
  });

  it("defaults to inbox when query is empty object", () => {
    expect(buildThreadQueryFilter(cache, {})).toEqual({
      inMailbox: "mb-inbox",
    });
  });

  it("filters by fromEmail", () => {
    const result = buildThreadQueryFilter(cache, {
      fromEmail: "alice@example.com",
    });
    expect(result).toEqual({ from: "alice@example.com" });
  });

  it("filters by labelId", () => {
    const result = buildThreadQueryFilter(cache, {
      labelId: "mb-sent",
    });
    expect(result).toEqual({ inMailbox: "mb-sent" });
  });

  it("filters by unread", () => {
    const result = buildThreadQueryFilter(cache, { isUnread: true });
    expect(result).toEqual({ notKeyword: "$seen" });
  });

  it("filters by date range", () => {
    const after = new Date("2024-01-01");
    const before = new Date("2024-12-31");
    const result = buildThreadQueryFilter(cache, { after, before });
    expect(result).toEqual({
      operator: "AND",
      conditions: [
        { after: after.toISOString() },
        { before: before.toISOString() },
      ],
    });
  });

  it("combines multiple filters with AND operator", () => {
    const result = buildThreadQueryFilter(cache, {
      fromEmail: "alice@example.com",
      labelId: "mb-inbox",
      isUnread: true,
    });
    expect(result).toEqual({
      operator: "AND",
      conditions: [
        { from: "alice@example.com" },
        { inMailbox: "mb-inbox" },
        { notKeyword: "$seen" },
      ],
    });
  });
});

describe("buildPaginationFilter", () => {
  it("returns empty object with no options", () => {
    expect(buildPaginationFilter(cache, {})).toEqual({});
  });

  it("filters by inbox when inboxOnly is true", () => {
    const result = buildPaginationFilter(cache, { inboxOnly: true });
    expect(result).toEqual({ inMailbox: "mb-inbox" });
  });

  it("filters by text query", () => {
    const result = buildPaginationFilter(cache, { query: "hello" });
    expect(result).toEqual({ text: "hello" });
  });

  it("filters by unread", () => {
    const result = buildPaginationFilter(cache, { unreadOnly: true });
    expect(result).toEqual({ notKeyword: "$seen" });
  });

  it("filters by date range", () => {
    const before = new Date("2024-12-31");
    const after = new Date("2024-01-01");
    const result = buildPaginationFilter(cache, { before, after });
    expect(result).toEqual({
      operator: "AND",
      conditions: [
        { before: before.toISOString() },
        { after: after.toISOString() },
      ],
    });
  });

  it("combines multiple filters with AND operator", () => {
    const result = buildPaginationFilter(cache, {
      inboxOnly: true,
      query: "urgent",
      unreadOnly: true,
    });
    expect(result).toEqual({
      operator: "AND",
      conditions: [
        { inMailbox: "mb-inbox" },
        { text: "urgent" },
        { notKeyword: "$seen" },
      ],
    });
  });
});
