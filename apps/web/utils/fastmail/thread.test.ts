import { describe, it, expect, vi } from "vitest";
import { getThread, getThreadEmailIds, getThreadsBatch } from "./thread";
import type { FastmailClient } from "./client";
import type { JmapGetResponse, JmapThread, JmapEmail } from "./types";

function makeClient(
  responses: Array<[string, Record<string, unknown>, string]>,
): FastmailClient {
  return {
    request: vi.fn().mockResolvedValue(responses),
  } as unknown as FastmailClient;
}

function makeThreadResponse(
  threads: JmapThread[],
): [string, JmapGetResponse<JmapThread>, string] {
  return ["Thread/get", { list: threads, state: "s1", notFound: [] }, "t"];
}

function makeEmailResponse(
  emails: Partial<JmapEmail>[],
): [string, JmapGetResponse<JmapEmail>, string] {
  const defaults: JmapEmail = {
    id: "e1",
    blobId: "b1",
    threadId: "t1",
    mailboxIds: { inbox: true },
    keywords: { $seen: true },
    receivedAt: "2024-06-01T12:00:00Z",
    sentAt: "2024-06-01T11:59:00Z",
    size: 100,
    subject: "Test",
    from: [{ name: "Alice", email: "alice@example.com" }],
    to: [{ name: "Bob", email: "bob@example.com" }],
    cc: null,
    bcc: null,
    replyTo: null,
    messageId: ["<msg-1@example.com>"],
    inReplyTo: null,
    references: null,
    hasAttachment: false,
    preview: "Preview text",
    textBody: [{ partId: "text", type: "text/plain", size: 10 }],
    htmlBody: [{ partId: "html", type: "text/html", size: 20 }],
    attachments: [],
    bodyValues: {
      text: {
        value: "Plain text",
        isTruncated: false,
        isEncodingProblem: false,
      },
      html: {
        value: "<p>HTML</p>",
        isTruncated: false,
        isEncodingProblem: false,
      },
    },
  };

  return [
    "Email/get",
    {
      list: emails.map((e) => ({ ...defaults, ...e })),
      state: "s1",
      notFound: [],
    },
    "e",
  ];
}

describe("getThread", () => {
  it("returns thread and parsed messages", async () => {
    const thread: JmapThread = { id: "t1", emailIds: ["e1", "e2"] };
    const client = makeClient([
      makeThreadResponse([thread]),
      makeEmailResponse([
        { id: "e1", threadId: "t1", subject: "First" },
        { id: "e2", threadId: "t1", subject: "Second" },
      ]),
    ]);

    const result = await getThread(client, {
      accountId: "acc1",
      threadId: "t1",
    });

    expect(result.thread).toEqual(thread);
    expect(result.messages).toHaveLength(2);
    expect(result.messages[0].id).toBe("e1");
    expect(result.messages[0].subject).toBe("First");
    expect(result.messages[1].id).toBe("e2");
    expect(result.messages[1].subject).toBe("Second");
  });

  it("throws when thread is not found", async () => {
    const client = makeClient([makeThreadResponse([]), makeEmailResponse([])]);

    await expect(
      getThread(client, { accountId: "acc1", threadId: "t-missing" }),
    ).rejects.toThrow("Thread not found: t-missing");
  });

  it("uses backreference chaining in method calls", async () => {
    const client = makeClient([
      makeThreadResponse([{ id: "t1", emailIds: ["e1"] }]),
      makeEmailResponse([{ id: "e1", threadId: "t1" }]),
    ]);

    await getThread(client, { accountId: "acc1", threadId: "t1" });

    const calls = (client.request as ReturnType<typeof vi.fn>).mock.calls[0][0];
    expect(calls).toHaveLength(2);
    expect(calls[0][0]).toBe("Thread/get");
    expect(calls[1][0]).toBe("Email/get");
    expect(calls[1][1]["#ids"]).toEqual({
      resultOf: "t",
      name: "Thread/get",
      path: "/list/*/emailIds",
    });
  });

  it("parses messages with correct labels", async () => {
    const client = makeClient([
      makeThreadResponse([{ id: "t1", emailIds: ["e1"] }]),
      makeEmailResponse([
        {
          id: "e1",
          threadId: "t1",
          mailboxIds: { inbox: true, archive: true },
          keywords: { $flagged: true },
        },
      ]),
    ]);

    const result = await getThread(client, {
      accountId: "acc1",
      threadId: "t1",
    });

    expect(result.messages[0].labelIds).toContain("inbox");
    expect(result.messages[0].labelIds).toContain("archive");
    expect(result.messages[0].labelIds).toContain("UNREAD");
    expect(result.messages[0].labelIds).toContain("STARRED");
  });
});

describe("getThreadEmailIds", () => {
  it("returns email IDs from a thread", async () => {
    const client = makeClient([
      makeThreadResponse([{ id: "t1", emailIds: ["e1", "e2", "e3"] }]),
    ]);

    const ids = await getThreadEmailIds(client, {
      accountId: "acc1",
      threadId: "t1",
    });

    expect(ids).toEqual(["e1", "e2", "e3"]);
  });

  it("throws when thread is not found", async () => {
    const client = makeClient([makeThreadResponse([])]);

    await expect(
      getThreadEmailIds(client, { accountId: "acc1", threadId: "t-missing" }),
    ).rejects.toThrow("Thread not found: t-missing");
  });

  it("only makes one method call (no Email/get)", async () => {
    const client = makeClient([
      makeThreadResponse([{ id: "t1", emailIds: ["e1"] }]),
    ]);

    await getThreadEmailIds(client, { accountId: "acc1", threadId: "t1" });

    const calls = (client.request as ReturnType<typeof vi.fn>).mock.calls[0][0];
    expect(calls).toHaveLength(1);
    expect(calls[0][0]).toBe("Thread/get");
  });
});

describe("getThreadsBatch", () => {
  it("returns a map of thread IDs to email IDs", async () => {
    const client = makeClient([
      makeThreadResponse([
        { id: "t1", emailIds: ["e1", "e2"] },
        { id: "t2", emailIds: ["e3"] },
      ]),
    ]);

    const result = await getThreadsBatch(client, {
      accountId: "acc1",
      threadIds: ["t1", "t2"],
    });

    expect(result).toBeInstanceOf(Map);
    expect(result.get("t1")).toEqual(["e1", "e2"]);
    expect(result.get("t2")).toEqual(["e3"]);
  });

  it("returns empty map for empty input", async () => {
    const client = makeClient([]);

    const result = await getThreadsBatch(client, {
      accountId: "acc1",
      threadIds: [],
    });

    expect(result.size).toBe(0);
    expect(client.request).not.toHaveBeenCalled();
  });

  it("handles threads not found (missing from response)", async () => {
    const client = makeClient([
      makeThreadResponse([{ id: "t1", emailIds: ["e1"] }]),
    ]);

    const result = await getThreadsBatch(client, {
      accountId: "acc1",
      threadIds: ["t1", "t-missing"],
    });

    expect(result.get("t1")).toEqual(["e1"]);
    expect(result.has("t-missing")).toBe(false);
  });

  it("sends all thread IDs in a single request", async () => {
    const client = makeClient([
      makeThreadResponse([
        { id: "t1", emailIds: ["e1"] },
        { id: "t2", emailIds: ["e2"] },
        { id: "t3", emailIds: ["e3"] },
      ]),
    ]);

    await getThreadsBatch(client, {
      accountId: "acc1",
      threadIds: ["t1", "t2", "t3"],
    });

    const calls = (client.request as ReturnType<typeof vi.fn>).mock.calls[0][0];
    expect(calls).toHaveLength(1);
    expect(calls[0][1].ids).toEqual(["t1", "t2", "t3"]);
  });
});
