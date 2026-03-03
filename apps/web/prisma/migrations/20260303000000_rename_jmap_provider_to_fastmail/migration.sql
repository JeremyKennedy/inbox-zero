-- Matches the provider name change from "jmap" to "fastmail" in the codebase
UPDATE "Account" SET provider = 'fastmail' WHERE provider = 'jmap';
