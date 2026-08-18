// Server version as a literal. The stdio entry reads package.json at runtime,
// but the Worker has no filesystem, so the value is duplicated here and the
// test harness asserts the two stay in sync.
export const VERSION = "0.8.0";
