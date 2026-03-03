import type {
  JmapThread,
  JmapEmail,
  JmapMethodCall,
  JmapGetResponse,
} from "@/utils/fastmail/types";
import type { FastmailClient } from "@/utils/fastmail/client";
import type { ParsedMessage } from "@/utils/types";
import {
  EMAIL_PROPERTIES_FULL,
  FETCH_BODY_OPTIONS,
} from "@/utils/fastmail/constants";
import { parseJmapEmail } from "@/utils/fastmail/message";

type GetThreadOptions = {
  accountId: string;
  threadId: string;
};

type GetThreadResult = {
  thread: JmapThread;
  messages: ParsedMessage[];
};

type GetThreadEmailIdsOptions = {
  accountId: string;
  threadId: string;
};

type GetThreadsBatchOptions = {
  accountId: string;
  threadIds: string[];
};

export async function getThread(
  client: FastmailClient,
  options: GetThreadOptions,
): Promise<GetThreadResult> {
  const { accountId, threadId } = options;

  const methodCalls: JmapMethodCall[] = [
    ["Thread/get", { accountId, ids: [threadId] }, "t"],
    [
      "Email/get",
      {
        accountId,
        "#ids": {
          resultOf: "t",
          name: "Thread/get",
          path: "/list/*/emailIds",
        },
        properties: EMAIL_PROPERTIES_FULL,
        ...FETCH_BODY_OPTIONS,
      },
      "e",
    ],
  ];

  const responses = await client.request(methodCalls);
  const threadData = responses[0][1] as JmapGetResponse<JmapThread>;
  const emailData = responses[1][1] as JmapGetResponse<JmapEmail>;

  if (threadData.list.length === 0) {
    throw new Error(`Thread not found: ${threadId}`);
  }

  const thread = threadData.list[0];
  const messages = emailData.list.map(parseJmapEmail);

  return { thread, messages };
}

export async function getThreadEmailIds(
  client: FastmailClient,
  options: GetThreadEmailIdsOptions,
): Promise<string[]> {
  const { accountId, threadId } = options;

  const methodCalls: JmapMethodCall[] = [
    ["Thread/get", { accountId, ids: [threadId] }, "t"],
  ];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapGetResponse<JmapThread>;

  if (data.list.length === 0) {
    throw new Error(`Thread not found: ${threadId}`);
  }

  return data.list[0].emailIds;
}

export async function getThreadsBatch(
  client: FastmailClient,
  options: GetThreadsBatchOptions,
): Promise<Map<string, string[]>> {
  const { accountId, threadIds } = options;

  if (threadIds.length === 0) return new Map();

  const methodCalls: JmapMethodCall[] = [
    ["Thread/get", { accountId, ids: threadIds }, "t"],
  ];

  const responses = await client.request(methodCalls);
  const data = responses[0][1] as JmapGetResponse<JmapThread>;

  const result = new Map<string, string[]>();
  for (const thread of data.list) {
    result.set(thread.id, thread.emailIds);
  }

  return result;
}
