Determine the next version tag and create a release, following these steps:

1. Run `git tag --sort=-v:refname | head -1` to find the latest existing tag (if any).
2. Run `git log --oneline` (and `git log <last_tag>..HEAD --oneline` if a tag exists) to see what commits are included in this release.
3. Determine the next semantic version:
   - If no tags exist, start at `v1.0.0`.
   - Bump the patch version by default (e.g. `v1.0.1`). Bump minor if new features are present, major if breaking changes are present.
4. Based on the commits, write concise release notes as a bullet list — group by feature, fix, and improvement. Focus on user-facing changes. Do not include refactor-only or internal-only commits.
5. Show the user the proposed version and release notes and ask for confirmation before proceeding.
6. Once confirmed, run:

```bash
git tag <version> -m "Release <version>

<release notes>" && git push origin <version>
```
