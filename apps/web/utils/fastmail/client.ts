import type {
  JmapSession,
  JmapMethodCall,
  JmapMethodResponse,
} from "@/utils/fastmail/types";
import {
  JMAP_SESSION_URL,
  WANTED_CAPABILITIES,
} from "@/utils/fastmail/constants";
import type { Logger } from "@/utils/logger";

const MAX_RETRIES = 3;
const BASE_DELAY_MS = 1000;
const RETRYABLE_STATUS_CODES = new Set([429, 502, 503, 504]);

export class FastmailClient {
  private session: JmapSession | null = null;
  private sessionPromise: Promise<JmapSession> | null = null;
  private readonly accessToken: string;
  private readonly logger: Logger;

  constructor(accessToken: string, logger: Logger) {
    this.accessToken = accessToken;
    this.logger = logger;
  }

  async ensureSession(): Promise<void> {
    if (this.session) return;

    // Deduplicate concurrent session fetches
    if (!this.sessionPromise) {
      this.sessionPromise = this.fetchSession();
    }

    this.session = await this.sessionPromise;
    this.sessionPromise = null;
  }

  async request(methodCalls: JmapMethodCall[]): Promise<JmapMethodResponse[]> {
    await this.ensureSession();
    const session = this.session!;

    const body = JSON.stringify({
      using: this.getAvailableCapabilities(),
      methodCalls,
    });

    const response = await this.fetchWithRetry(session.apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${this.accessToken}`,
      },
      body,
    });

    const data = await response.json();
    const responses = data.methodResponses as JmapMethodResponse[];

    for (const methodResponse of responses) {
      if (methodResponse[0] === "error") {
        const errorBody = methodResponse[1] as { type?: string };
        const callId = methodResponse[2];
        throw new Error(
          `JMAP error in ${callId}: ${errorBody.type ?? "unknown"}`,
        );
      }
    }

    return responses;
  }

  async getAccountId(): Promise<string> {
    await this.ensureSession();
    return this.session!.primaryAccounts["urn:ietf:params:jmap:mail"]!;
  }

  async getSession(): Promise<JmapSession> {
    await this.ensureSession();
    return this.session!;
  }

  getAccessToken(): string {
    return this.accessToken;
  }

  hasCapability(capability: string): boolean {
    if (!this.session) return false;
    return capability in this.session.capabilities;
  }

  async uploadBlob(data: ArrayBuffer, type: string): Promise<string> {
    await this.ensureSession();
    const session = this.session!;
    const accountId = session.primaryAccounts["urn:ietf:params:jmap:mail"]!;

    const uploadUrl = session.uploadUrl.replace(
      "{accountId}",
      encodeURIComponent(accountId),
    );

    const response = await fetch(uploadUrl, {
      method: "POST",
      headers: {
        "Content-Type": type,
        Authorization: `Bearer ${this.accessToken}`,
      },
      body: data,
    });

    const result = await response.json();
    return result.blobId as string;
  }

  private getAvailableCapabilities(): string[] {
    if (!this.session) return [];
    return WANTED_CAPABILITIES.filter(
      (cap) => cap in this.session!.capabilities,
    );
  }

  private async fetchSession(): Promise<JmapSession> {
    this.logger.info("Fetching JMAP session");

    const response = await fetch(JMAP_SESSION_URL, {
      headers: { Authorization: `Bearer ${this.accessToken}` },
    });

    if (!response.ok) {
      throw new Error(
        `JMAP session fetch failed: ${response.status} ${response.statusText}`,
      );
    }

    const session = (await response.json()) as JmapSession;

    const accountId = session.primaryAccounts["urn:ietf:params:jmap:mail"];
    if (!accountId) {
      throw new Error(
        "JMAP session has no primary account for urn:ietf:params:jmap:mail",
      );
    }

    const available = WANTED_CAPABILITIES.filter(
      (cap) => cap in session.capabilities,
    );
    this.logger.info("JMAP session established", {
      capabilities: available.length,
    });
    this.logger.trace("JMAP session user", { username: session.username });

    return session;
  }

  private async fetchWithRetry(
    url: string,
    init: RequestInit,
  ): Promise<Response> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      if (attempt > 0) {
        const delay = BASE_DELAY_MS * 2 ** (attempt - 1);
        this.logger.warn(`JMAP request retry ${attempt}/${MAX_RETRIES}`, {
          delay,
        });
        await sleep(delay);
      }

      let response: Response;
      try {
        response = await fetch(url, init);
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));
        continue;
      }

      if (response.ok) return response;

      if (RETRYABLE_STATUS_CODES.has(response.status)) {
        const retryAfter = response.headers.get("Retry-After");
        if (retryAfter && attempt < MAX_RETRIES) {
          const retryMs = Number.parseInt(retryAfter, 10) * 1000;
          if (!Number.isNaN(retryMs) && retryMs > 0) {
            this.logger.warn("Respecting Retry-After header", {
              retryAfterSeconds: retryAfter,
            });
            await sleep(retryMs);
          }
        }
        lastError = new Error(
          `JMAP request failed: ${response.status} ${response.statusText}`,
        );
        continue;
      }

      throw new Error(
        `JMAP request failed: ${response.status} ${response.statusText}`,
      );
    }

    throw lastError ?? new Error("JMAP request failed after retries");
  }
}

export function createFastmailClient(
  token: string,
  logger: Logger,
): FastmailClient {
  return new FastmailClient(token, logger);
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
