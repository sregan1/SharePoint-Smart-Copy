# Releasing

Releases are created by pushing a **version tag**. Pushing to `main` does not release anything.

1. Bump `<Version>` in `SharePointSmartCopy/SharePointSmartCopy.csproj` (for example `4.1.0`).
2. Add a `## 4.1.0` section to `CHANGELOG.md`. Its text becomes the GitHub release notes (if the section is missing, the notes link to the changelog).
3. Commit both changes and merge them to `main`.
4. Tag and push the tag:

   ```powershell
   git tag v4.1.0
   git push origin main v4.1.0
   ```

The **Release** workflow (`.github/workflows/release.yml`) runs on the tag, and:

- fails immediately if the tag does not equal `v` + the csproj `<Version>` (so a tag can never ship a binary stamped with a different version);
- publishes the self-contained single-file `SharePointSmartCopy.exe`;
- creates the GitHub release with the exe and a zip attached;
- never replaces an existing release (to republish, delete the release manually first, or bump the version).

CI (`.github/workflows/ci.yml`) still builds and tests every push and pull request to `main`.
