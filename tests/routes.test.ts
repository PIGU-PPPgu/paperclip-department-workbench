import assert from "node:assert/strict";
import test from "node:test";
import { officePath } from "../src/routes.ts";

test("dashboard and project links retain the active organization prefix", () => {
  assert.equal(officePath("CMP", "/CMP/dashboard"), "/CMP/office");
  assert.equal(officePath(null, "/CMP/projects/123", "?project=123&view=milestones"), "/CMP/office?project=123&view=milestones");
});
