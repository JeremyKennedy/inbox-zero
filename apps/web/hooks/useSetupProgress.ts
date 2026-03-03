import { useMemo } from "react";
import type { GetSetupProgressResponse } from "@/app/api/user/setup-progress/route";
import { useAccount } from "@/providers/EmailAccountProvider";
import { isFastmailProvider } from "@/utils/email/provider-types";
import { useSWRWithEmailAccount } from "@/utils/swr";

export function useSetupProgress() {
  const { provider } = useAccount();
  const result = useSWRWithEmailAccount<GetSetupProgressResponse>(
    "/api/user/setup-progress",
  );

  // Calendar integration only supports Google/Outlook OAuth — exclude for Fastmail
  const adjusted = useMemo(() => {
    if (!result.data || !isFastmailProvider(provider)) return result.data;

    const calendarWasCounted = result.data.steps.calendarConnected;
    const completed = calendarWasCounted
      ? result.data.completed - 1
      : result.data.completed;
    const total = result.data.total - 1;

    return {
      ...result.data,
      completed,
      total,
      isComplete: completed >= total,
    };
  }, [result.data, provider]);

  return { ...result, data: adjusted };
}
