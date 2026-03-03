import { describe, it, expect } from "vitest";
import { parseJmapEmail, calculateNextPageToken } from "./message";
import type { JmapEmail } from "./types";

function makeEmail(overrides: Partial<JmapEmail> = {}): JmapEmail {
  return {
    id: "e1",
    blobId: "b1",
    threadId: "t1",
    mailboxIds: { inbox1: true },
    keywords: { $seen: true },
    receivedAt: "2024-06-01T12:00:00Z",
    sentAt: "2024-06-01T11:59:00Z",
    size: 1234,
    subject: "Hello",
    from: [{ name: "Alice", email: "alice@example.com" }],
    to: [{ name: "Bob", email: "bob@example.com" }],
    cc: null,
    bcc: null,
    replyTo: null,
    messageId: ["<msg-1@example.com>"],
    inReplyTo: null,
    references: null,
    hasAttachment: false,
    preview: "This is a preview",
    textBody: [{ partId: "text", type: "text/plain", size: 100 }],
    htmlBody: [{ partId: "html", type: "text/html", size: 200 }],
    attachments: [],
    bodyValues: {
      text: {
        value: "Hello plain",
        isTruncated: false,
        isEncodingProblem: false,
      },
      html: {
        value: "<p>Hello html</p>",
        isTruncated: false,
        isEncodingProblem: false,
      },
    },
    ...overrides,
  };
}

/** Simulates a JMAP response with only EMAIL_PROPERTIES_MINIMAL (no body/attachments). */
function makeMinimalEmail(overrides: Partial<JmapEmail> = {}): JmapEmail {
  return {
    id: "e1",
    blobId: "b1",
    threadId: "t1",
    mailboxIds: { inbox1: true },
    keywords: { $seen: true },
    receivedAt: "2024-06-01T12:00:00Z",
    sentAt: "2024-06-01T11:59:00Z",
    size: 1234,
    subject: "Hello",
    from: [{ name: "Alice", email: "alice@example.com" }],
    to: [{ name: "Bob", email: "bob@example.com" }],
    cc: null,
    bcc: null,
    replyTo: null,
    messageId: ["<msg-1@example.com>"],
    inReplyTo: null,
    references: null,
    hasAttachment: false,
    preview: "This is a preview",
    ...overrides,
  };
}

describe("parseJmapEmail", () => {
  it("parses basic fields correctly", () => {
    const result = parseJmapEmail(makeEmail());

    expect(result.id).toBe("e1");
    expect(result.threadId).toBe("t1");
    expect(result.subject).toBe("Hello");
    expect(result.snippet).toBe("This is a preview");
    expect(result.textPlain).toBe("Hello plain");
    expect(result.textHtml).toBe("<p>Hello html</p>");
    expect(result.headers.subject).toBe("Hello");
    expect(result.headers.from).toBe("Alice <alice@example.com>");
    expect(result.headers.to).toBe("Bob <bob@example.com>");
    expect(result.headers["message-id"]).toBe("<msg-1@example.com>");
  });

  it("uses sentAt for date, falling back to receivedAt", () => {
    const withSentAt = parseJmapEmail(makeEmail());
    expect(withSentAt.date).toBe("2024-06-01T11:59:00Z");

    const withoutSentAt = parseJmapEmail(makeEmail({ sentAt: null }));
    expect(withoutSentAt.date).toBe("2024-06-01T12:00:00Z");
  });

  it("injects UNREAD label when $seen keyword is missing", () => {
    const unread = parseJmapEmail(makeEmail({ keywords: {} }));
    expect(unread.labelIds).toContain("UNREAD");

    const read = parseJmapEmail(makeEmail({ keywords: { $seen: true } }));
    expect(read.labelIds).not.toContain("UNREAD");
  });

  it("injects STARRED label when $flagged keyword is present", () => {
    const starred = parseJmapEmail(
      makeEmail({ keywords: { $seen: true, $flagged: true } }),
    );
    expect(starred.labelIds).toContain("STARRED");

    const unstarred = parseJmapEmail(makeEmail({ keywords: { $seen: true } }));
    expect(unstarred.labelIds).not.toContain("STARRED");
  });

  it("includes mailbox IDs as label IDs", () => {
    const result = parseJmapEmail(
      makeEmail({ mailboxIds: { mb1: true, mb2: true } }),
    );
    expect(result.labelIds).toContain("mb1");
    expect(result.labelIds).toContain("mb2");
  });

  it("handles both UNREAD and STARRED simultaneously", () => {
    const result = parseJmapEmail(
      makeEmail({
        mailboxIds: { inbox: true },
        keywords: { $flagged: true },
      }),
    );
    expect(result.labelIds).toContain("UNREAD");
    expect(result.labelIds).toContain("STARRED");
    expect(result.labelIds).toContain("inbox");
  });

  it("formats address with name correctly", () => {
    const result = parseJmapEmail(
      makeEmail({
        from: [{ name: "Alice Smith", email: "alice@example.com" }],
      }),
    );
    expect(result.headers.from).toBe("Alice Smith <alice@example.com>");
  });

  it("formats address without name as just email", () => {
    const result = parseJmapEmail(
      makeEmail({
        from: [{ name: null, email: "alice@example.com" }],
      }),
    );
    expect(result.headers.from).toBe("alice@example.com");
  });

  it("formats address where name equals email as just email", () => {
    const result = parseJmapEmail(
      makeEmail({
        from: [{ name: "alice@example.com", email: "alice@example.com" }],
      }),
    );
    expect(result.headers.from).toBe("alice@example.com");
  });

  it("handles multiple addresses in to/cc/bcc", () => {
    const result = parseJmapEmail(
      makeEmail({
        to: [
          { name: "Bob", email: "bob@example.com" },
          { name: null, email: "carol@example.com" },
        ],
        cc: [{ name: "Dave", email: "dave@example.com" }],
      }),
    );
    expect(result.headers.to).toBe("Bob <bob@example.com>, carol@example.com");
    expect(result.headers.cc).toBe("Dave <dave@example.com>");
  });

  it("returns undefined for empty cc/bcc", () => {
    const result = parseJmapEmail(makeEmail({ cc: null, bcc: null }));
    expect(result.headers.cc).toBeUndefined();
    expect(result.headers.bcc).toBeUndefined();
  });

  it("converts attachments to the expected format", () => {
    const result = parseJmapEmail(
      makeEmail({
        attachments: [
          {
            partId: "att1",
            blobId: "blob-123",
            name: "document.pdf",
            type: "application/pdf",
            size: 5000,
          },
        ],
      }),
    );
    expect(result.attachments).toHaveLength(1);
    expect(result.attachments![0]).toEqual({
      filename: "document.pdf",
      mimeType: "application/pdf",
      size: 5000,
      attachmentId: "blob-123",
      headers: {
        "content-type": "application/pdf",
        "content-description": "",
        "content-transfer-encoding": "",
        "content-id": "",
      },
    });
  });

  it("handles attachment with null name", () => {
    const result = parseJmapEmail(
      makeEmail({
        attachments: [
          {
            partId: "att1",
            blobId: "blob-123",
            name: null,
            type: "application/octet-stream",
            size: 100,
          },
        ],
      }),
    );
    expect(result.attachments![0].filename).toBe("");
  });

  it("returns undefined attachments when empty", () => {
    const result = parseJmapEmail(makeEmail({ attachments: [] }));
    expect(result.attachments).toBeUndefined();
  });

  it("returns undefined textPlain/textHtml when body parts are empty", () => {
    const result = parseJmapEmail(
      makeEmail({
        textBody: [],
        htmlBody: [],
        bodyValues: {},
      }),
    );
    expect(result.textPlain).toBeUndefined();
    expect(result.textHtml).toBeUndefined();
  });

  it("sets historyId to empty string", () => {
    const result = parseJmapEmail(makeEmail());
    expect(result.historyId).toBe("");
  });

  it("parses inReplyTo and references headers", () => {
    const result = parseJmapEmail(
      makeEmail({
        inReplyTo: ["<reply-to@example.com>"],
        references: ["<ref1@example.com>", "<ref2@example.com>"],
      }),
    );
    expect(result.headers["in-reply-to"]).toBe("<reply-to@example.com>");
    expect(result.headers.references).toBe(
      "<ref1@example.com> <ref2@example.com>",
    );
  });

  it("handles minimal properties (no body/attachments fields)", () => {
    // When fetchBody is false, JMAP only returns EMAIL_PROPERTIES_MINIMAL.
    // textBody, htmlBody, attachments, and bodyValues are all undefined.
    const result = parseJmapEmail(makeMinimalEmail());

    expect(result.id).toBe("e1");
    expect(result.threadId).toBe("t1");
    expect(result.subject).toBe("Hello");
    expect(result.snippet).toBe("This is a preview");
    expect(result.textPlain).toBeUndefined();
    expect(result.textHtml).toBeUndefined();
    expect(result.attachments).toBeUndefined();
  });

  it("converts receivedAt to millisecond timestamp for internalDate", () => {
    const result = parseJmapEmail(
      makeEmail({ receivedAt: "2024-06-01T12:00:00Z" }),
    );
    expect(result.internalDate).toBe(
      String(new Date("2024-06-01T12:00:00Z").getTime()),
    );
  });
});

describe("calculateNextPageToken", () => {
  it("returns next position when more results exist", () => {
    expect(calculateNextPageToken(0, 50, 100)).toBe(50);
  });

  it("returns undefined when at the end", () => {
    expect(calculateNextPageToken(50, 50, 100)).toBeUndefined();
  });

  it("returns undefined when past the end", () => {
    expect(calculateNextPageToken(90, 50, 100)).toBeUndefined();
  });

  it("returns undefined when exactly at total", () => {
    expect(calculateNextPageToken(0, 100, 100)).toBeUndefined();
  });

  it("handles zero results", () => {
    expect(calculateNextPageToken(0, 0, 0)).toBeUndefined();
  });

  it("handles position in the middle", () => {
    expect(calculateNextPageToken(25, 25, 100)).toBe(50);
  });
});
