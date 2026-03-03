import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { FastmailClient, createFastmailClient } from "./client";
import type { Logger } from "@/utils/logger";

const MOCK_SESSION = {
  apiUrl: "https://api.fastmail.com/jmap/api/",
  downloadUrl:
    "https://api.fastmail.com/jmap/download/{accountId}/{blobId}/{name}",
  uploadUrl: "https://api.fastmail.com/jmap/upload/{accountId}/",
  eventSourceUrl: "https://api.fastmail.com/jmap/eventsource/",
  username: "test@example.com",
  primaryAccounts: { "urn:ietf:params:jmap:mail": "acc-123" },
  capabilities: {
    "urn:ietf:params:jmap:core": {},
    "urn:ietf:params:jmap:mail": {},
    "urn:ietf:params:jmap:submission": {},
  },
  state: "s1",
};

function makeLogger(): Logger {
  return {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    trace: vi.fn(),
    debug: vi.fn(),
    child: vi.fn().mockReturnThis(),
  } as unknown as Logger;
}

let fetchMock: ReturnType<typeof vi.fn>;

beforeEach(() => {
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
});

afterEach(() => {
  vi.restoreAllMocks();
});

function mockSessionFetch() {
  fetchMock.mockResolvedValueOnce({
    ok: true,
    json: () => Promise.resolve(MOCK_SESSION),
  });
}

function mockApiResponse(methodResponses: unknown[]) {
  fetchMock.mockResolvedValueOnce({
    ok: true,
    json: () => Promise.resolve({ methodResponses }),
  });
}

describe("createFastmailClient", () => {
  it("returns a FastmailClient instance", () => {
    const client = createFastmailClient("token-123", makeLogger());
    expect(client).toBeInstanceOf(FastmailClient);
  });
});

describe("FastmailClient", () => {
  describe("ensureSession", () => {
    it("fetches session on first call", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      await client.ensureSession();

      expect(fetchMock).toHaveBeenCalledWith(
        "https://api.fastmail.com/jmap/session",
        expect.objectContaining({
          headers: { Authorization: "Bearer token-123" },
        }),
      );
    });

    it("caches session across multiple calls", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      await client.ensureSession();
      await client.ensureSession();
      await client.ensureSession();

      // Only one fetch to the session URL
      expect(fetchMock).toHaveBeenCalledTimes(1);
    });

    it("deduplicates concurrent session fetches", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      // Fire three concurrent calls
      await Promise.all([
        client.ensureSession(),
        client.ensureSession(),
        client.ensureSession(),
      ]);

      expect(fetchMock).toHaveBeenCalledTimes(1);
    });

    it("throws on invalid token (non-ok response)", async () => {
      const client = createFastmailClient("bad-token", makeLogger());
      fetchMock.mockResolvedValueOnce({
        ok: false,
        status: 401,
        statusText: "Unauthorized",
      });

      await expect(client.ensureSession()).rejects.toThrow(
        "JMAP session fetch failed: 401 Unauthorized",
      );
    });

    it("throws when session has no mail account", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      fetchMock.mockResolvedValueOnce({
        ok: true,
        json: () =>
          Promise.resolve({
            ...MOCK_SESSION,
            primaryAccounts: {},
          }),
      });

      await expect(client.ensureSession()).rejects.toThrow(
        "JMAP session has no primary account",
      );
    });
  });

  describe("getAccountId", () => {
    it("returns the primary mail account ID", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      const accountId = await client.getAccountId();
      expect(accountId).toBe("acc-123");
    });
  });

  describe("getAccessToken", () => {
    it("returns the token without fetching session", () => {
      const client = createFastmailClient("token-123", makeLogger());
      expect(client.getAccessToken()).toBe("token-123");
      expect(fetchMock).not.toHaveBeenCalled();
    });
  });

  describe("hasCapability", () => {
    it("returns false before session is fetched", () => {
      const client = createFastmailClient("token-123", makeLogger());
      expect(client.hasCapability("urn:ietf:params:jmap:submission")).toBe(
        false,
      );
    });

    it("returns true for available capabilities after session", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      await client.ensureSession();

      expect(client.hasCapability("urn:ietf:params:jmap:core")).toBe(true);
      expect(client.hasCapability("urn:ietf:params:jmap:mail")).toBe(true);
      expect(client.hasCapability("urn:ietf:params:jmap:submission")).toBe(
        true,
      );
    });

    it("returns false for unavailable capabilities", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      await client.ensureSession();

      expect(client.hasCapability("urn:ietf:params:jmap:contacts")).toBe(false);
    });
  });

  describe("request", () => {
    it("sends method calls to the API URL", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      mockApiResponse([
        ["Mailbox/get", { list: [], state: "s1", notFound: [] }, "mb"],
      ]);

      await client.request([["Mailbox/get", { accountId: "acc-123" }, "mb"]]);

      // Second call is to the API URL
      expect(fetchMock).toHaveBeenCalledTimes(2);
      const apiCall = fetchMock.mock.calls[1];
      expect(apiCall[0]).toBe("https://api.fastmail.com/jmap/api/");
      const body = JSON.parse(apiCall[1].body);
      expect(body.methodCalls).toEqual([
        ["Mailbox/get", { accountId: "acc-123" }, "mb"],
      ]);
    });

    it("filters capabilities to only those available in session", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      // Session with only core and mail (no submission)
      fetchMock.mockResolvedValueOnce({
        ok: true,
        json: () =>
          Promise.resolve({
            ...MOCK_SESSION,
            capabilities: {
              "urn:ietf:params:jmap:core": {},
              "urn:ietf:params:jmap:mail": {},
            },
          }),
      });
      mockApiResponse([["Email/get", {}, "e"]]);

      await client.request([["Email/get", { accountId: "acc-123" }, "e"]]);

      const body = JSON.parse(fetchMock.mock.calls[1][1].body);
      expect(body.using).toEqual([
        "urn:ietf:params:jmap:core",
        "urn:ietf:params:jmap:mail",
      ]);
      expect(body.using).not.toContain("urn:ietf:params:jmap:submission");
    });

    it("throws on JMAP method-level errors", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      mockApiResponse([["error", { type: "serverFail" }, "mb"]]);

      await expect(
        client.request([["Mailbox/get", { accountId: "acc-123" }, "mb"]]),
      ).rejects.toThrow("JMAP error in mb: serverFail");
    });

    it("throws on JMAP error with unknown type", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      mockApiResponse([["error", {}, "mb"]]);

      await expect(
        client.request([["Mailbox/get", { accountId: "acc-123" }, "mb"]]),
      ).rejects.toThrow("JMAP error in mb: unknown");
    });
  });

  describe("retry logic", () => {
    it("retries on 429 status", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      // First attempt: 429
      fetchMock.mockResolvedValueOnce({
        ok: false,
        status: 429,
        statusText: "Too Many Requests",
        headers: new Map(),
      });
      // Second attempt: success
      mockApiResponse([["Email/get", {}, "e"]]);

      const result = await client.request([
        ["Email/get", { accountId: "acc-123" }, "e"],
      ]);

      // session + 429 + success = 3 fetches
      expect(fetchMock).toHaveBeenCalledTimes(3);
      expect(result).toEqual([["Email/get", {}, "e"]]);
    });

    it("retries on 502/503/504 status", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      fetchMock.mockResolvedValueOnce({
        ok: false,
        status: 502,
        statusText: "Bad Gateway",
        headers: new Map(),
      });
      mockApiResponse([["Email/get", {}, "e"]]);

      await client.request([["Email/get", { accountId: "acc-123" }, "e"]]);
      expect(fetchMock).toHaveBeenCalledTimes(3);
    });

    it("throws immediately on non-retryable status (e.g. 400)", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      fetchMock.mockResolvedValueOnce({
        ok: false,
        status: 400,
        statusText: "Bad Request",
      });

      await expect(
        client.request([["Email/get", { accountId: "acc-123" }, "e"]]),
      ).rejects.toThrow("JMAP request failed: 400 Bad Request");

      // session + one failed attempt (no retry)
      expect(fetchMock).toHaveBeenCalledTimes(2);
    });

    it("retries on network errors", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      fetchMock.mockRejectedValueOnce(new Error("network timeout"));
      mockApiResponse([["Email/get", {}, "e"]]);

      await client.request([["Email/get", { accountId: "acc-123" }, "e"]]);
      expect(fetchMock).toHaveBeenCalledTimes(3);
    });

    it("throws after exhausting retries", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();

      // All attempts fail with 503 (initial + 3 retries = 4 total)
      for (let i = 0; i <= 3; i++) {
        fetchMock.mockResolvedValueOnce({
          ok: false,
          status: 503,
          statusText: "Service Unavailable",
          headers: new Map(),
        });
      }

      await expect(
        client.request([["Email/get", { accountId: "acc-123" }, "e"]]),
      ).rejects.toThrow("JMAP request failed: 503 Service Unavailable");

      // session + 4 failed API attempts = 5
      expect(fetchMock).toHaveBeenCalledTimes(5);
    }, 15_000);
  });

  describe("uploadBlob", () => {
    it("uploads data to the upload URL", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      await client.ensureSession();

      fetchMock.mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ blobId: "blob-abc" }),
      });

      const data = new ArrayBuffer(10);
      const blobId = await client.uploadBlob(data, "application/pdf");

      expect(blobId).toBe("blob-abc");
      const uploadCall = fetchMock.mock.calls[1];
      expect(uploadCall[0]).toBe(
        "https://api.fastmail.com/jmap/upload/acc-123/",
      );
      expect(uploadCall[1].method).toBe("POST");
      expect(uploadCall[1].headers["Content-Type"]).toBe("application/pdf");
    });

    it("throws on upload failure", async () => {
      const client = createFastmailClient("token-123", makeLogger());
      mockSessionFetch();
      await client.ensureSession();

      fetchMock.mockResolvedValueOnce({
        ok: false,
        status: 413,
        statusText: "Payload Too Large",
      });

      await expect(
        client.uploadBlob(new ArrayBuffer(10), "application/pdf"),
      ).rejects.toThrow("JMAP blob upload failed: 413 Payload Too Large");
    });
  });
});
