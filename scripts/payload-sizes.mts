// The payload-size baseline (#156). The work is in src/testing/payload-sizes/measure.ts.
//
// Run: npm run payload:sizes              (compare against the committed baseline)
//      npm run payload:sizes -- --update  (rewrite the baseline with the current figures)

import { runPayloadSizes } from '../src/testing/payload-sizes/measure.js';

process.exitCode = await runPayloadSizes(process.argv.slice(2));
