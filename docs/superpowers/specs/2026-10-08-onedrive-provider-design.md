# OneDrive cloud provider — design

## Context

This fork (`girishKM/ReFra`, forked from `IacobIonut01/ReFra`) replaces a
from-scratch gallery app (`onedrive-gallery`, Phases 0–3 done: MSAL sign-in,
recursive Graph listing, full-screen viewer) as the base for a personal
OneDrive gallery. ReFra already has a mature pluggable cloud-provider
architecture (Immich, Nextcloud, ownCloud, WebDAV, SMB, NFS, each its own
Gradle source set under `app/src/<provider>/kotlin/...`, gated by an
`app.properties` flag and wired through `ProviderRegistry`/`ProviderType`).
All of those are now disabled (`app.properties`); this spec adds OneDrive
as a new provider following the same pattern.

Checkpoint already done: the fork builds and runs on the dev machine
(Windows) with local photos visible, after three Windows-toolchain fixes to
the native build scripts (commit `a704d415`). This spec covers the OneDrive
provider implementation itself — no product code has been written yet.

## Decision: auth uses OAuth Device Code Flow, not MSAL

The existing `onedrive-gallery` Phase 1–3 code used MSAL (`PublicClientApplication`,
`BrowserTabActivity` redirect, interactive browser sign-in). ReFra's own auth
abstraction (`CloudInteractiveAuthHandler`: `begin(serverUrl) -> InteractiveAuthSession`,
`poll(session) -> InteractiveAuthPollResult`) is a **polling** model, built for
flows like Nextcloud's Login Flow v2 (open a browser URL, poll an endpoint
until the user finishes in-browser, get credentials back).

Microsoft identity platform's **OAuth 2.0 Device Authorization Grant** (device
code flow) is exactly this shape: call `POST /devicecode`, show the user a URL
+ short code, poll `POST /token` until they approve on another device/tab. It
needs only plain HTTP calls (no MSAL, no `BrowserTabActivity`, no redirect URI
registration, no app-side webview). This removes an entire class of problems
we hit in `onedrive-gallery` (redirect URI mismatch, single-account `signIn()`
vs `acquireToken()`) by construction — there's no redirect URI at all in this
flow.

**Decision: implement `OneDriveDeviceCodeClient` using the device code flow,
dropping the MSAL dependency entirely.** `CloudServerConfig.apiKey` stores the
refresh token (encrypted at rest like other providers' secrets, via
`CredentialEncryptor`); `username` stores the account's UPN/email; `serverUrl`
is the fixed constant `https://graph.microsoft.com` (OneDrive has no
user-configurable server, unlike Nextcloud/WebDAV).

Azure app registration changes: the existing registration
(`c623491e-a0a4-4e9e-9732-9d86d9733352`) can be reused, but device code flow
requires the registration's Authentication blade to have "Allow public client
flows" enabled, and does not need the Android platform/redirect URI entry at
all. This gets re-verified once the client is implemented.

## Components

New source set `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/`
(mirroring `immich`'s layout, the closest existing analog — both are
REST+JSON APIs with their own OAuth-shaped auth, unlike the WebDAV family):

- `OneDriveProvider.kt` — implements `RemoteMediaProvider` + `SyncCapableProvider`.
- `data/api/OneDriveApiService.kt` — Retrofit/OkHttp Graph v1.0 client:
  paginated `/me/drive/root/children` (and recursive descent into
  subfolders — ported logic from `onedrive-gallery`'s `GraphDriveApi`),
  `/me/drive/root/delta` for sync, upload endpoints.
- `data/api/OneDriveAuthInterceptor.kt` — attaches `Authorization: Bearer`,
  refreshes the access token from the stored refresh token on 401.
- `data/dto/OneDriveDtos.kt` — `driveItem`, `thumbnails`, `@microsoft.graph.downloadUrl`,
  delta response shapes.
- `auth/OneDriveDeviceCodeClient.kt` — the device code flow itself (begin/poll
  against Microsoft's `/devicecode` and `/token` endpoints).
- `auth/OneDriveInteractiveAuthHandler.kt` — thin `CloudInteractiveAuthHandler`
  adapter over the client (mirrors `NextcloudInteractiveAuthHandler`).
- `di/OneDriveModule.kt` — Hilt bindings (mirrors `ImmichModule.kt`).

Plus the usual plumbing other providers have:
- `ProviderType.ONEDRIVE` entry + `BuildConfig.ONEDRIVE_ENABLED` (mirrors the
  existing `IMMICH_ENABLED` etc. pattern in `app/build.gradle.kts`).
- `INCLUDE_ONEDRIVE` flag in `app.properties` (default `true` for this fork).
- `app/src/noonedrive/kotlin/com/dot/gallery/cloud/onedrive/di/OneDriveModule.kt`
  stub (mirrors `noimmich`), for when the flag is off.
- `app/src/main/res/drawable/ic_provider_onedrive.xml` (the repo already
  reserves `ic_provider_azure.xml`, unused — replace or add alongside it).

## Data mapping

Graph `driveItem` → `CloudMediaEntity`: `id` → `remoteId`, `name`, `size`,
`@microsoft.graph.downloadUrl` → original URL (`getOriginalUrl`), `thumbnails[0].medium.url`
→ thumbnail (`getThumbnailUrl`), `file.hashes.quickXorHash` or `sha256Hash` →
content hash for `bulkUploadCheck`/`remoteExists` dedup, `parentReference.path`
→ `relativePath` (feeds the default `downloadSubPath`).

`getRemoteAssets` / recursive listing: port `onedrive-gallery`'s
`GraphDriveApi`/`PhotoRepository` folder-queue traversal logic (BFS over
subfolders via `/children`, following `@odata.nextLink`) — this was the one
genuinely tricky bug from that project (root `/children` alone misses
nested folders like `Pictures/Camera Roll`) and should be ported as-is
rather than rediscovered.

`getSyncDelta`: Graph's `/me/drive/root/delta` endpoint natively returns
added/changed/deleted items since a `deltaLink` watermark — maps directly
onto `SyncDelta(items, deletedRemoteIds, completeRemoteIds)`, including the
`completeRemoteIds` full-reconciliation case on a from-scratch delta token.

`uploadAsset`: Graph simple upload (`PUT /items/{parent}:/name:/content`) for
files ≤4MB; upload session (`POST .../createUploadSession` + chunked `PUT`)
above that, matching Graph's documented limits.

## Testing

Unit tests (JVM, MockWebServer + fakes), mirroring the existing
`NextcloudLoginFlowClientTest`/`CloudUriTest`/`UploadTargetResolverTest`
patterns already in `app/src/test/java/com/dot/gallery/cloud/`:
- `OneDriveDeviceCodeClientTest` — begin/poll state machine against a fake
  Microsoft token endpoint (pending, complete, expired, denied).
- `OneDriveApiServiceTest` — paginated listing + recursive folder traversal
  (the ported `onedrive-gallery` logic), delta parsing, DTO mapping.
- `OneDriveProviderTest` — capability surface, thumbnail/original URL
  construction, upload-target resolution.

## Explicitly out of scope for this change

- Rebranding (app name, icon, README cleanup, stripping donation/Crowdin
  banners) — separate follow-up task once the provider works end-to-end.
- `onedrive-gallery` (the original from-scratch repo) is left as-is; it
  already works standalone and isn't being merged file-for-file into ReFra.
  Its Graph API logic is reference material for the port described above,
  not a drop-in dependency.
- People/OCR/map/memories capability interfaces — OneDrive personal accounts
  don't expose equivalent Graph APIs for these; `OneDriveProvider` simply
  doesn't implement those capability interfaces, same as the path-based
  providers (SMB/NFS/WebDAV) already don't.
