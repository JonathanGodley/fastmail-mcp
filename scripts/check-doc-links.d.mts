// Types for check-doc-links.mjs. Hand-written: keep in step with its exports.

/** One link, with its file relative to the root (`/` separators) and 1-based line. */
export interface LinkFinding {
  file: string;
  line: number;
  target: string;
  problem: 'missing file' | 'no such anchor' | 'anchor on a non-.md target';
}

/** Checks every relative link in the docs under `root`. */
export declare function checkDocLinks(root: string): {
  checked: number;
  bad: LinkFinding[];
  notes: LinkFinding[];
};
