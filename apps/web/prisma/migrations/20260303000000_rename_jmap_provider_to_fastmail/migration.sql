-- Matches the provider name change from "jmap" to "fastmail" in the codebase
UPDATE "Account" SET provider = 'fastmail' WHERE provider = 'jmap';

-- Update providerAccountId prefix from "jmap:" to "fastmail:" to match the auth plugin
UPDATE "Account"
  SET "providerAccountId" = 'fastmail:' || substring("providerAccountId" from 6)
  WHERE "providerAccountId" LIKE 'jmap:%';
