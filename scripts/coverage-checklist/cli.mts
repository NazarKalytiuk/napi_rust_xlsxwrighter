// npm run coverage:sync             — refresh test call sites, counts and reports; documents
//                                     the release with rustdoc when Cargo.lock pins a new one
// npm run coverage:sync -- --rustdoc — re-document the pinned release even if it is unchanged
// npm run coverage:check            — fail if the checklist is out of date (no rustdoc needed)
import { run } from './generate.mts';

run({ check: process.argv.includes('--check'), rustdoc: process.argv.includes('--rustdoc') });
