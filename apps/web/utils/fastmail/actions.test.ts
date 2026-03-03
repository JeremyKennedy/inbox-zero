import { describe, it, expect } from "vitest";
import { buildUpdateForAll } from "./actions";

describe("buildUpdateForAll", () => {
  it("applies the same patch to all email IDs", () => {
    const result = buildUpdateForAll(["e1", "e2", "e3"], {
      "keywords/$seen": true,
    });
    expect(result).toEqual({
      e1: { "keywords/$seen": true },
      e2: { "keywords/$seen": true },
      e3: { "keywords/$seen": true },
    });
  });

  it("handles empty email ID list", () => {
    const result = buildUpdateForAll([], { "keywords/$seen": true });
    expect(result).toEqual({});
  });

  it("handles single email ID", () => {
    const result = buildUpdateForAll(["e1"], {
      "mailboxIds/inbox": null,
      "mailboxIds/archive": true,
    });
    expect(result).toEqual({
      e1: {
        "mailboxIds/inbox": null,
        "mailboxIds/archive": true,
      },
    });
  });

  it("handles complex patches with mixed null and true values", () => {
    const patch = {
      "mailboxIds/inbox-id": null,
      "mailboxIds/trash-id": true,
    };
    const result = buildUpdateForAll(["e1", "e2"], patch);
    expect(result.e1).toEqual(patch);
    expect(result.e2).toEqual(patch);
  });

  it("creates independent copies for each email ID", () => {
    const result = buildUpdateForAll(["e1", "e2"], {
      "keywords/$seen": true,
    });
    expect(result.e1).not.toBe(result.e2);
  });
});
