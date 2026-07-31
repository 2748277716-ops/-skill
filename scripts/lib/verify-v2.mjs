import { verifyLongAlignmentFast } from "./verify-fast.mjs";

export { verifyWideAlignment } from "./verify.mjs";

export function verifyLongAlignment(sourceModel, outputModel, config = {}) {
  return verifyLongAlignmentFast(sourceModel, outputModel, config.cityContext);
}
