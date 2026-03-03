import { APIError, createAuthEndpoint } from "better-auth/api";
import { setSessionCookie } from "better-auth/cookies";
import { z } from "zod";
import prisma from "@/utils/prisma";
import { WELCOME_PATH } from "@/utils/config";
import { isInternalPath } from "@/utils/path";

const JMAP_SESSION_URL = "https://api.fastmail.com/jmap/session";
const FASTMAIL_PROVIDER = "fastmail";
const FASTMAIL_PROVIDER_ACCOUNT_PREFIX = "fastmail:";

const fastmailSignInSchema = z.object({
  apiToken: z.string().min(1),
  callbackURL: z.string().optional(),
});

type FastmailUser = {
  id: string;
  email: string;
  name: string;
  emailVerified: boolean;
  createdAt: Date;
  updatedAt: Date;
  image?: string | null;
};

type FastmailInternalAdapter = {
  findUserByEmail: (email: string) => Promise<{ user: FastmailUser } | null>;
  createUser: (user: {
    email: string;
    name: string;
    emailVerified: boolean;
  }) => Promise<FastmailUser | null>;
  createSession: (
    userId: string,
  ) => Promise<{ id: string; token: string } | null>;
};

async function discoverJmapSession(
  token: string,
): Promise<{ email: string; name: string; accountId: string }> {
  const res = await fetch(JMAP_SESSION_URL, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    throw new APIError("UNAUTHORIZED", {
      message: "Invalid Fastmail API token",
    });
  }

  const session = await res.json();
  const email = session.username;
  const accountId = session.primaryAccounts["urn:ietf:params:jmap:mail"];

  if (!email || !accountId) {
    throw new APIError("INTERNAL_SERVER_ERROR", {
      message: "Could not determine email from JMAP session",
    });
  }

  return { email, name: email.split("@")[0], accountId };
}

export function fastmailAuthPlugin() {
  return {
    id: "fastmail-auth",
    endpoints: {
      signInFastmail: createAuthEndpoint(
        "/sign-in/fastmail",
        {
          method: "POST",
          body: fastmailSignInSchema,
        },
        async (ctx) => {
          const { apiToken, callbackURL } = ctx.body;

          const { email, name } = await discoverJmapSession(apiToken);

          const user = await getOrCreateUser(
            ctx.context.internalAdapter,
            email,
            name,
          );

          await ensureFastmailAccount(user, email, name, apiToken);

          const session = await ctx.context.internalAdapter.createSession(
            user.id,
          );
          if (!session) {
            throw new APIError("INTERNAL_SERVER_ERROR", {
              message: "Failed to create session",
            });
          }

          await setSessionCookie(ctx, { session, user }, false);

          const redirectUrl = isInternalPath(callbackURL)
            ? callbackURL
            : WELCOME_PATH;

          return ctx.json({ callbackURL: redirectUrl });
        },
      ),
    },
  };
}

async function getOrCreateUser(
  internalAdapter: FastmailInternalAdapter,
  email: string,
  name: string,
): Promise<FastmailUser> {
  const existing = await internalAdapter.findUserByEmail(email);
  if (existing?.user) return existing.user;

  const created = await internalAdapter.createUser({
    email,
    name,
    emailVerified: true,
  });

  if (!created) {
    throw new APIError("INTERNAL_SERVER_ERROR", {
      message: "Failed to create user",
    });
  }

  return created;
}

async function ensureFastmailAccount(
  user: FastmailUser,
  email: string,
  name: string,
  apiToken: string,
) {
  const providerAccountId = `${FASTMAIL_PROVIDER_ACCOUNT_PREFIX}${email}`;

  const account = await prisma.account.upsert({
    where: {
      provider_providerAccountId: {
        provider: FASTMAIL_PROVIDER,
        providerAccountId,
      },
    },
    update: {
      userId: user.id,
      disconnectedAt: null,
      access_token: apiToken,
      refresh_token: apiToken,
      expires_at: getFutureDate(),
      refreshTokenExpiresAt: getFutureDate(),
    },
    create: {
      userId: user.id,
      provider: FASTMAIL_PROVIDER,
      providerAccountId,
      access_token: apiToken,
      refresh_token: apiToken,
      expires_at: getFutureDate(),
      refreshTokenExpiresAt: getFutureDate(),
    },
    select: { id: true },
  });

  await prisma.emailAccount.upsert({
    where: { email },
    update: {
      userId: user.id,
      accountId: account.id,
      name,
      image: null,
    },
    create: {
      email,
      userId: user.id,
      accountId: account.id,
      name,
      image: null,
    },
  });
}

function getFutureDate() {
  const date = new Date();
  date.setFullYear(date.getFullYear() + 10);
  return date;
}
