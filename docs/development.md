# Development

    npm install          # Node >= 24
    npm run dev          # tsx, no build step
    npm run typecheck
    npm test             # node --test, no test framework
    npm run demo         # the two demo sources, plus all profiles under build/profiles/
    npm run build        # tsc -p tsconfig.build.json (tests excluded from dist)
    npm start            # node dist/server.js, after a build

## Release

    npm version minor -m "chore: release %s"    # or patch / major
    git push && git push --tags

The version is stated in two places — `package.json`, which builds the image, and
`src/version.ts`, which is what a client is told on `initialize` and what every
produced document records as its producer. `npm version` keeps them in step: its
`version` lifecycle hook runs `scripts/sync-version.mjs` and stages the result,
so the release commit carries both. `src/version.test.ts` is the backstop for a
release made some other way.

Pushing the tag is what runs the release workflow: verify, then the ECR and GHCR
images, the GitHub release notes, and the GitOps dispatch.

## What the tests cover

Tests cover the pure decisions — format detection, the zip budget, the HWP
record walk and its control-character table, the spreadsheet's column
arithmetic and row budget, RTF's destinations and escapes, and the Markdown
parser — and, for every writer, a structural reopen. Markdown outputs are also
read back, and the spreadsheet writer is inspected for values, formulas and
cached results, so both directions fail together or not at all.

The PDF round trip is why `unpdf` is still a devDependency: the reader left with
the URL side, but rendering a PDF nothing can read is worth catching.

Nothing touches the network, and now nothing can: there is no client here.

What tests cannot cover is what a document *looks like*. Open the output — a
`.docx` in Word or Google Docs, a `.pdf` in a viewer, a `.hwpx` in 한글, a
`.pptx` in PowerPoint or Keynote — before trusting a change to a renderer. This
matters most for `pptx`, whose line counting is an estimate: a slide that
overflows is visible only on the screen.

`Verify` runs typecheck, the tests and a `docker build` on every pull request.
The release workflow runs the code checks again on the tag.
