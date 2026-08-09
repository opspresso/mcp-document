/**
 * The shape extracted text comes out in, shared by every reader that produces
 * lines rather than pages.
 *
 * One owner because the shape is a contract: a caller comparing two documents,
 * or a test asserting a round trip, is comparing this normalisation as much as
 * it is comparing the text. Two readers that trimmed differently would make
 * "the same document" depend on which one read it.
 */

/**
 * Trim each line, drop the blanks at either end, and never allow two in a row.
 *
 * A run of whitespace collapses to a single character, and to a **tab** when
 * the run held one. In these formats a tab is not spacing that survived from a
 * source file — it is an element somebody inserted (`w:tab`, `hp:tab`), and it
 * is how columns are laid out in a document that has no table. Flattening it to
 * a space merges the columns.
 */
export function normalize(lines: readonly string[]): string {
  const out: string[] = [];
  for (const raw of lines) {
    // A trailing cell separator is an artifact of the row's last cell closing,
    // not content — and a row whose last cells are empty produces several of
    // them in a row, so they all go rather than one of them.
    const line = raw
      .replace(/[^\S\n]+/g, (run) => (run.includes("\t") ? "\t" : " "))
      .trim()
      .replace(/(?:\s*\|)+$/, "")
      .trim();
    if (line === "") {
      if (out.length > 0 && out[out.length - 1] !== "") {
        out.push("");
      }
      continue;
    }
    out.push(line);
  }
  while (out.length > 0 && out[out.length - 1] === "") {
    out.pop();
  }
  return out.join("\n");
}
