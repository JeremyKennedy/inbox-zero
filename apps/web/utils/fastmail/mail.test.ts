import { describe, it, expect } from "vitest";
import { buildEmailObject } from "./mail";

describe("buildEmailObject", () => {
  it("builds minimal email with required fields", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Hello",
      keywords: { $seen: true },
    });

    expect(result.mailboxIds).toEqual({ "mb-sent": true });
    expect(result.to).toEqual([{ name: null, email: "bob@example.com" }]);
    expect(result.subject).toBe("Hello");
    expect(result.keywords).toEqual({ $seen: true });
    expect(result.from).toBeUndefined();
    expect(result.cc).toBeUndefined();
    expect(result.bcc).toBeUndefined();
    expect(result.replyTo).toBeUndefined();
  });

  it("includes from when provided", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      from: "alice@example.com",
      to: ["bob@example.com"],
      subject: "Hello",
      keywords: { $seen: true },
    });

    expect(result.from).toEqual([{ name: null, email: "alice@example.com" }]);
  });

  it("includes cc and bcc when provided", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      cc: ["carol@example.com"],
      bcc: ["dave@example.com"],
      subject: "Hello",
      keywords: {},
    });

    expect(result.cc).toEqual([{ name: null, email: "carol@example.com" }]);
    expect(result.bcc).toEqual([{ name: null, email: "dave@example.com" }]);
  });

  it("omits cc and bcc when empty arrays", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      cc: [],
      bcc: [],
      subject: "Hello",
      keywords: {},
    });

    expect(result.cc).toBeUndefined();
    expect(result.bcc).toBeUndefined();
  });

  it("includes replyTo when provided", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      replyTo: "reply@example.com",
      subject: "Hello",
      keywords: {},
    });

    expect(result.replyTo).toEqual([
      { name: null, email: "reply@example.com" },
    ]);
  });

  it("includes text body", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Hello",
      textBody: "Plain text content",
      keywords: {},
    });

    expect(result.textBody).toEqual([{ partId: "text", type: "text/plain" }]);
    expect(
      (result.bodyValues as Record<string, { value: string }>).text.value,
    ).toBe("Plain text content");
  });

  it("includes HTML body", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Hello",
      htmlBody: "<p>HTML content</p>",
      keywords: {},
    });

    expect(result.htmlBody).toEqual([{ partId: "html", type: "text/html" }]);
    expect(
      (result.bodyValues as Record<string, { value: string }>).html.value,
    ).toBe("<p>HTML content</p>");
  });

  it("includes both text and HTML body", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Hello",
      textBody: "Plain text",
      htmlBody: "<p>HTML</p>",
      keywords: {},
    });

    expect(result.textBody).toBeDefined();
    expect(result.htmlBody).toBeDefined();
    const bodyValues = result.bodyValues as Record<string, { value: string }>;
    expect(bodyValues.text.value).toBe("Plain text");
    expect(bodyValues.html.value).toBe("<p>HTML</p>");
  });

  it("includes inReplyTo as array", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Re: Hello",
      inReplyTo: "<msg-123@example.com>",
      keywords: {},
    });

    expect(result.inReplyTo).toEqual(["<msg-123@example.com>"]);
  });

  it("splits references string into array", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Re: Hello",
      references: "<ref1@example.com> <ref2@example.com>",
      keywords: {},
    });

    expect(result.references).toEqual([
      "<ref1@example.com>",
      "<ref2@example.com>",
    ]);
  });

  it("includes threadId when provided", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "Re: Hello",
      threadId: "thread-123",
      keywords: {},
    });

    expect(result.threadId).toBe("thread-123");
  });

  it("includes attachments when provided", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "With attachment",
      keywords: {},
      attachmentBlobs: [
        {
          blobId: "blob-1",
          name: "file.pdf",
          contentType: "application/pdf",
        },
      ],
    });

    expect(result.attachments).toEqual([
      { blobId: "blob-1", name: "file.pdf", type: "application/pdf" },
    ]);
  });

  it("omits attachments when empty array", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com"],
      subject: "No attachment",
      keywords: {},
      attachmentBlobs: [],
    });

    expect(result.attachments).toBeUndefined();
  });

  it("handles draft keywords", () => {
    const result = buildEmailObject({
      mailboxId: "mb-drafts",
      to: ["bob@example.com"],
      subject: "Draft",
      keywords: { $seen: true, $draft: true },
    });

    expect(result.keywords).toEqual({ $seen: true, $draft: true });
  });

  it("handles multiple recipients", () => {
    const result = buildEmailObject({
      mailboxId: "mb-sent",
      to: ["bob@example.com", "carol@example.com"],
      subject: "Group email",
      keywords: {},
    });

    expect(result.to).toEqual([
      { name: null, email: "bob@example.com" },
      { name: null, email: "carol@example.com" },
    ]);
  });
});
