# OneDrive Cloud Provider (Read-Only Viewing) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make OneDrive photos sign in and show up in the ReFra gallery grid, read-only.

**Architecture:** A new Gradle source set `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/` implements `RemoteMediaProvider` against the Microsoft Graph REST API, following the existing Immich provider's shape (plain Retrofit + Gson, no special SDK). Auth uses the OAuth 2.0 device code flow (plain HTTP, no MSAL) via a `CloudInteractiveAuthHandler`, exactly like Nextcloud's Login Flow v2 client — both are poll-based, so no changes to the shared auth/UI infrastructure are needed.

**Tech Stack:** Kotlin, Retrofit + Gson, OkHttp, Hilt, Room (existing `CloudMediaEntity`), JUnit + MockWebServer for tests.

**Spec:** `docs/superpowers/specs/2026-10-08-onedrive-provider-design.md`

## Global Constraints

- No MSAL dependency — OAuth 2.0 device code flow only, implemented with plain OkHttp calls to `https://login.microsoftonline.com/common/oauth2/v2.0/devicecode` and `.../token`.
- Azure app registration `c623491e-a0a4-4e9e-9732-9d86d9733352` is reused. It must have "Allow public client flows" enabled on its Authentication blade (a manual Azure Portal step — not code, verify before Task 6).
- `CloudServerConfig` field reuse (no schema changes): `serverUrl` = fixed `"https://graph.microsoft.com/v1.0"`, `username` = the account's UPN/email, `password` = the OAuth refresh token (encrypted at rest via the existing `CredentialEncryptor`, same as every other provider's secrets).
- This plan's scope is read-only viewing only. `OneDriveProvider` declares capabilities `{REMOTE_ASSETS, TEXT_SEARCH}` — no `FAVORITE`/`TRASH`/`ARCHIVE`/`REMOTE_ALBUMS`/`ALBUM_WRITE`/`SYNC`. Methods for those (`toggleFavorite`, `trashAsset`, `restoreAsset`, `createAlbum`, `addToAlbum`, `emptyTrash`, `restoreAllTrash`) are still implemented (the interface requires it) but return `Result.failure(UnsupportedOperationException(...))` — real, compilable code, intentionally inert because the UI never calls them without the matching capability flag.
- `deleteAsset` IS implemented for real (a real Graph `DELETE`), per the existing SMB/NFS/WebDAV convention: providers without `TRASH` must still support a genuine permanent delete; `trashAsset` is NOT an alias for it in this plan (no delete-from-grid UI is being wired up here — out of scope — so `trashAsset` can safely stay `UnsupportedOperationException` until a later plan wires up deletion from the UI).
- Only files whose Graph `driveItem` has an `image` facet are listed (folders and non-image files are skipped) — same filter as the original `onedrive-gallery` project.

## Review Focus

- **Pagination across many pages:** a OneDrive folder with more photos than one page's `$top` must not drop or duplicate items when following `@odata.nextLink`. Covered in Task 2.
- **Multi-level nested folders:** `Pictures/Camera Roll/2026/` (two levels deep) must still be found — traversal must recurse fully, not just one level into root. Covered in Task 2.
- **Non-image files mixed into a folder:** PDFs, videos, `.docx` files alongside photos must not break parsing and must be skipped by the `image`-facet filter. Covered in Task 2.
- **Expired or revoked refresh token:** if the user revokes the app's access from their Microsoft account, a 401 on every subsequent request must surface as a clear re-authentication error, not an infinite refresh loop or crash. Covered in Task 4.
- **User declines or the device code expires:** on the Microsoft sign-in page, the user can decline consent, or simply never finish within the code's ~15 minute window. `poll()` must distinguish `AUTHENTICATION` (declined) from `EXPIRED` (timed out) so the UI shows an accurate message, not a generic failure. Covered in Task 3.

---

### Task 1: Provider scaffolding (enum, build flags, empty stub)

**Files:**
- Modify: `app/build.gradle.kts`
- Modify: `app.properties`
- Modify: `app/src/main/kotlin/com/dot/gallery/cloud/core/ProviderType.kt`
- Create: `app/src/noonedrive/kotlin/com/dot/gallery/cloud/onedrive/di/OneDriveModule.kt`

**Interfaces:**
- Produces: `ProviderType.ONEDRIVE` enum entry, `BuildConfig.ONEDRIVE_ENABLED`, the `app/src/onedrive/kotlin/...` source directory wired into compilation when enabled.

This task is pure scaffolding — no new behavior, so it ends with "the app still builds and all existing tests still pass" rather than a new test.

- [ ] **Step 1: Add the `includeOnedrive` property + BuildConfig field + source-set wiring**

In `app/build.gradle.kts`, add a new property getter next to `includeImmich` (same file/pattern, ~line 653):

```kotlin
val includeOnedrive: Boolean
    get() {
        if (isOffline) return false
        val fl = rootProject.file("app.properties")
        return try {
            val properties = Properties()
            properties.load(FileInputStream(fl))
            properties.getProperty("INCLUDE_ONEDRIVE", "false").toBoolean()
        } catch (_: Exception) {
            false
        }
    }
```

Add the matching `buildConfigField` line next to each `IMMICH_ENABLED` line (there are four: debug, release, staging, gplay build types):

```kotlin
buildConfigField("Boolean", "ONEDRIVE_ENABLED", "$includeOnedrive")
```

Add the source-set conditional next to the `includeImmich`/`includeOwncloud` block inside `sourceSets { getByName("main") { ... } }`:

```kotlin
if (includeOnedrive) {
    kotlin.srcDir("src/onedrive/kotlin")
} else {
    kotlin.srcDir("src/noonedrive/kotlin")
}
```

- [ ] **Step 2: Enable it in `app.properties`**

Add this line to `app.properties` (alongside the existing `INCLUDE_*` lines):

```
INCLUDE_ONEDRIVE=true
```

- [ ] **Step 3: Add the `ProviderType.ONEDRIVE` enum entry**

In `app/src/main/kotlin/com/dot/gallery/cloud/core/ProviderType.kt`, add a new entry and `when` branch:

```kotlin
@Serializable
enum class ProviderType(val displayName: String, val isRemote: Boolean) {
    IMMICH("Immich", isRemote = true),
    OWNCLOUD("ownCloud", isRemote = true),
    NEXTCLOUD("Nextcloud", isRemote = true),
    WEBDAV("WebDAV", isRemote = true),
    SMB("SMB", isRemote = true),
    NFS("NFS", isRemote = true),
    ONEDRIVE("OneDrive", isRemote = true),
    LOCAL_PEOPLE("On-Device Faces", isRemote = false),
    LOCAL_OCR("On-Device OCR", isRemote = false),
    LOCAL_CLIP("On-Device Search", isRemote = false);

    val isIncludedInBuild: Boolean
        get() = when (this) {
            IMMICH -> BuildConfig.IMMICH_ENABLED
            OWNCLOUD -> BuildConfig.OWNCLOUD_ENABLED
            NEXTCLOUD -> BuildConfig.NEXTCLOUD_ENABLED
            WEBDAV -> BuildConfig.WEBDAV_ENABLED
            SMB -> BuildConfig.SMB_ENABLED
            NFS -> BuildConfig.NFS_ENABLED
            ONEDRIVE -> BuildConfig.ONEDRIVE_ENABLED
            else -> true
        }
    // ... companion object unchanged
}
```

- [ ] **Step 4: Create the empty `noonedrive` stub module**

```kotlin
// app/src/noonedrive/kotlin/com/dot/gallery/cloud/onedrive/di/OneDriveModule.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.di

import dagger.Module
import dagger.hilt.InstallIn
import dagger.hilt.components.SingletonComponent

@Module
@InstallIn(SingletonComponent::class)
object OneDriveModule
```

- [ ] **Step 5: Verify the app still builds**

Run: `./gradlew :app:assembleArm64-v8aNoMLDebug`
Expected: `BUILD SUCCESSFUL` (no `onedrive` source directory exists yet, so this compiles via the empty `noonedrive` stub only).

- [ ] **Step 6: Commit**

```bash
git add app/build.gradle.kts app.properties app/src/main/kotlin/com/dot/gallery/cloud/core/ProviderType.kt app/src/noonedrive/kotlin/com/dot/gallery/cloud/onedrive/di/OneDriveModule.kt
git commit -m "Add OneDrive provider scaffolding (enum, build flag, empty stub)"
```

---

### Task 2: Graph DTOs + recursive paginated listing

**Files:**
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/dto/OneDriveDtos.kt`
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/OneDriveAssetLister.kt`
- Test: `app/src/test/java/com/dot/gallery/cloud/onedrive/OneDriveAssetListerTest.kt`

**Interfaces:**
- Produces: `data class OneDriveItemDto`, `data class OneDriveChildrenResponseDto`, `class OneDriveAssetLister(httpClient: OkHttpClient, baseUrl: String = "https://graph.microsoft.com/v1.0")` with `suspend fun listAll(accessToken: String): List<OneDriveItemDto>`.
- Consumes: nothing from earlier tasks (this is pure data + one HTTP client).

This ports the exact recursive-folder-queue logic already proven working in the `onedrive-gallery` project's `GraphDriveApi`/`PhotoRepository` (root `/children` alone misses nested folders like `Pictures/Camera Roll` — that was a real bug found and fixed there).

- [ ] **Step 1: Write the DTOs**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/dto/OneDriveDtos.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.data.dto

import com.google.gson.annotations.SerializedName

data class OneDriveImageFacetDto(
    val width: Int = 0,
    val height: Int = 0
)

data class OneDriveFolderFacetDto(
    val childCount: Int = 0
)

data class OneDriveFileFacetDto(
    val mimeType: String? = null,
    val hashes: OneDriveHashesDto? = null
)

data class OneDriveHashesDto(
    val quickXorHash: String? = null,
    val sha256Hash: String? = null
)

data class OneDriveThumbnailSetDto(
    val id: String = "",
    val small: OneDriveThumbnailDto? = null,
    val medium: OneDriveThumbnailDto? = null,
    val large: OneDriveThumbnailDto? = null
)

data class OneDriveThumbnailDto(
    val url: String = "",
    val width: Int = 0,
    val height: Int = 0
)

data class OneDriveParentReferenceDto(
    val path: String? = null
)

data class OneDriveItemDto(
    val id: String,
    val name: String = "",
    val size: Long = 0L,
    @SerializedName("lastModifiedDateTime") val lastModifiedDateTime: String? = null,
    val image: OneDriveImageFacetDto? = null,
    val folder: OneDriveFolderFacetDto? = null,
    val file: OneDriveFileFacetDto? = null,
    val thumbnails: List<OneDriveThumbnailSetDto>? = null,
    val parentReference: OneDriveParentReferenceDto? = null,
    @SerializedName("@microsoft.graph.downloadUrl") val downloadUrl: String? = null
) {
    val isImage: Boolean get() = image != null
    val isFolder: Boolean get() = folder != null
}

data class OneDriveChildrenResponseDto(
    val value: List<OneDriveItemDto> = emptyList(),
    @SerializedName("@odata.nextLink") val nextLink: String? = null
)
```

- [ ] **Step 2: Write the failing test for `OneDriveAssetLister`**

```kotlin
// app/src/test/java/com/dot/gallery/cloud/onedrive/OneDriveAssetListerTest.kt
package com.dot.gallery.cloud.onedrive

import kotlinx.coroutines.test.runTest
import okhttp3.OkHttpClient
import okhttp3.mockwebserver.MockResponse
import okhttp3.mockwebserver.MockWebServer
import org.junit.After
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Before
import org.junit.Test

class OneDriveAssetListerTest {

    private lateinit var server: MockWebServer
    private lateinit var lister: OneDriveAssetLister

    @Before
    fun setUp() {
        server = MockWebServer()
        server.start()
        lister = OneDriveAssetLister(OkHttpClient(), baseUrl = server.url("/v1.0").toString())
    }

    @After
    fun tearDown() {
        server.shutdown()
    }

    @Test
    fun `listAll follows pagination and does not drop or duplicate items`() = runTest {
        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "value": [
                    { "id": "p1", "name": "a.jpg", "image": {} }
                  ],
                  "@odata.nextLink": "${server.url("/v1.0/me/drive/root/children?page=2")}"
                }
                """.trimIndent()
            )
        )
        server.enqueue(
            MockResponse().setBody(
                """
                { "value": [ { "id": "p2", "name": "b.jpg", "image": {} } ] }
                """.trimIndent()
            )
        )

        val items = lister.listAll(accessToken = "token")

        assertEquals(listOf("p1", "p2"), items.map { it.id })
    }

    @Test
    fun `listAll recurses into subfolders nested more than one level deep`() = runTest {
        // root: one subfolder
        server.enqueue(
            MockResponse().setBody(
                """
                { "value": [ { "id": "folder-pictures", "name": "Pictures", "folder": { "childCount": 1 } } ] }
                """.trimIndent()
            )
        )
        // Pictures/: one subfolder
        server.enqueue(
            MockResponse().setBody(
                """
                { "value": [ { "id": "folder-cameraroll", "name": "Camera Roll", "folder": { "childCount": 1 } } ] }
                """.trimIndent()
            )
        )
        // Pictures/Camera Roll/: the actual photo, two levels deep
        server.enqueue(
            MockResponse().setBody(
                """
                { "value": [ { "id": "deep-photo", "name": "deep.jpg", "image": {} } ] }
                """.trimIndent()
            )
        )

        val items = lister.listAll(accessToken = "token")

        assertEquals(listOf("deep-photo"), items.map { it.id })
    }

    @Test
    fun `listAll skips non-image files mixed into a folder`() = runTest {
        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "value": [
                    { "id": "doc-1", "name": "notes.pdf", "file": { "mimeType": "application/pdf" } },
                    { "id": "photo-1", "name": "a.jpg", "image": {} }
                  ]
                }
                """.trimIndent()
            )
        )

        val items = lister.listAll(accessToken = "token")

        assertEquals(listOf("photo-1"), items.map { it.id })
    }
}
```

- [ ] **Step 3: Run it to verify it fails**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.OneDriveAssetListerTest"`
Expected: compile failure — `OneDriveAssetLister` does not exist yet.

- [ ] **Step 4: Implement `OneDriveAssetLister`**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/OneDriveAssetLister.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive

import com.dot.gallery.cloud.onedrive.data.dto.OneDriveChildrenResponseDto
import com.dot.gallery.cloud.onedrive.data.dto.OneDriveItemDto
import com.google.gson.Gson
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import okhttp3.OkHttpClient
import okhttp3.Request
import java.io.IOException

/**
 * Walks the ENTIRE OneDrive file tree starting at the drive root, recursing into every
 * subfolder found at any depth. A plain `/root/children` call alone only lists items
 * sitting loose at the drive root — real photo libraries are nested (e.g.
 * `Pictures/Camera Roll/2026/`), so folders found on any page are queued and visited
 * before the walk is considered complete. Only items with an `image` facet are kept;
 * folders and non-image files are filtered out of the returned list (but folders are
 * still queued for traversal).
 */
class OneDriveAssetLister(
    private val client: OkHttpClient,
    private val baseUrl: String = "https://graph.microsoft.com/v1.0"
) {
    private val gson = Gson()

    suspend fun listAll(accessToken: String): List<OneDriveItemDto> = withContext(Dispatchers.IO) {
        val images = mutableListOf<OneDriveItemDto>()
        val pendingFolderUrls = ArrayDeque<String>()
        var nextUrl: String? = "$baseUrl/me/drive/root/children?\$top=200&\$expand=thumbnails"

        while (nextUrl != null || pendingFolderUrls.isNotEmpty()) {
            val url = nextUrl ?: pendingFolderUrls.removeFirst()
            val page = fetchPage(url, accessToken)

            for (item in page.value) {
                when {
                    item.isFolder -> pendingFolderUrls.addLast(
                        "$baseUrl/me/drive/items/${item.id}/children?\$top=200&\$expand=thumbnails"
                    )
                    item.isImage -> images += item
                    // non-image, non-folder files (videos, documents, ...) are skipped
                }
            }

            nextUrl = page.nextLink
        }

        images
    }

    private fun fetchPage(url: String, accessToken: String): OneDriveChildrenResponseDto {
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $accessToken")
            .build()

        client.newCall(request).execute().use { response ->
            if (!response.isSuccessful) {
                throw IOException("Graph API returned HTTP ${response.code} for $url")
            }
            val body = response.body?.string()
                ?: throw IOException("Empty response body from Graph API")
            return gson.fromJson(body, OneDriveChildrenResponseDto::class.java)
        }
    }
}
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.OneDriveAssetListerTest"`
Expected: `BUILD SUCCESSFUL`, 3 tests passed.

- [ ] **Step 6: Commit**

```bash
git add app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/dto/OneDriveDtos.kt app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/OneDriveAssetLister.kt app/src/test/java/com/dot/gallery/cloud/onedrive/OneDriveAssetListerTest.kt
git commit -m "Add OneDrive Graph DTOs and recursive paginated asset listing"
```

---

### Task 3: Device code auth client

**Files:**
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClient.kt`
- Test: `app/src/test/java/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClientTest.kt`

**Interfaces:**
- Produces: `class OneDriveDeviceCodeClient(client: OkHttpClient, clientId: String, authBaseUrl: String = "https://login.microsoftonline.com/common/oauth2/v2.0", graphBaseUrl: String = "https://graph.microsoft.com/v1.0")` with `suspend fun begin(): InteractiveAuthSession` and `suspend fun poll(session: InteractiveAuthSession): InteractiveAuthPollResult`.
- Consumes: `InteractiveAuthSession`, `InteractiveAuthPollResult`, `InteractiveAuthCredentials`, `InteractiveAuthException`, `InteractiveAuthErrorKind` (from `com.dot.gallery.cloud.core.auth`, already defined in `app/src/main`).

Microsoft's device code response's `verification_uri_complete` field (e.g. `https://microsoft.com/devicelogin?otc=ABC123`) already has the user's one-time code pre-filled, so the generic `InteractiveAuthSession.browserUrl` (a single URL to open) fits without needing a new field for displaying a separate code.

- [ ] **Step 1: Write the failing test**

```kotlin
// app/src/test/java/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClientTest.kt
package com.dot.gallery.cloud.onedrive.auth

import com.dot.gallery.cloud.core.auth.InteractiveAuthErrorKind
import com.dot.gallery.cloud.core.auth.InteractiveAuthException
import com.dot.gallery.cloud.core.auth.InteractiveAuthPollResult
import kotlinx.coroutines.test.runTest
import okhttp3.OkHttpClient
import okhttp3.mockwebserver.MockResponse
import okhttp3.mockwebserver.MockWebServer
import org.junit.After
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Before
import org.junit.Test

class OneDriveDeviceCodeClientTest {

    private lateinit var server: MockWebServer
    private lateinit var client: OneDriveDeviceCodeClient

    @Before
    fun setUp() {
        server = MockWebServer()
        server.start()
        client = OneDriveDeviceCodeClient(
            client = OkHttpClient(),
            clientId = "test-client-id",
            authBaseUrl = server.url("/auth").toString(),
            graphBaseUrl = server.url("/graph").toString()
        )
    }

    @After
    fun tearDown() {
        server.shutdown()
    }

    @Test
    fun `begin returns a session with the pre-filled verification url`() = runTest {
        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "device_code": "device-abc",
                  "user_code": "XYZ123",
                  "verification_uri": "https://microsoft.com/devicelogin",
                  "verification_uri_complete": "https://microsoft.com/devicelogin?otc=XYZ123",
                  "expires_in": 900,
                  "interval": 5
                }
                """.trimIndent()
            )
        )

        val session = client.begin()

        assertEquals("https://microsoft.com/devicelogin?otc=XYZ123", session.browserUrl)
        assertEquals("device-abc", session.token)
    }

    @Test
    fun `poll returns Pending while authorization_pending`() = runTest {
        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "device_code": "device-abc",
                  "user_code": "XYZ123",
                  "verification_uri": "https://microsoft.com/devicelogin",
                  "verification_uri_complete": "https://microsoft.com/devicelogin?otc=XYZ123",
                  "expires_in": 900,
                  "interval": 5
                }
                """.trimIndent()
            )
        )
        val session = client.begin()
        server.enqueue(MockResponse().setResponseCode(400).setBody("""{"error":"authorization_pending"}"""))

        val result = client.poll(session)

        assertTrue(result is InteractiveAuthPollResult.Pending)
    }

    @Test
    fun `poll completes with credentials once the user approves`() = runTest {
        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "device_code": "device-abc",
                  "user_code": "XYZ123",
                  "verification_uri": "https://microsoft.com/devicelogin",
                  "verification_uri_complete": "https://microsoft.com/devicelogin?otc=XYZ123",
                  "expires_in": 900,
                  "interval": 5
                }
                """.trimIndent()
            )
        )
        val session = client.begin()

        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "access_token": "access-xyz",
                  "refresh_token": "refresh-xyz",
                  "expires_in": 3600
                }
                """.trimIndent()
            )
        )
        server.enqueue(
            MockResponse().setBody(
                """{ "userPrincipalName": "person@example.com" }"""
            )
        )

        val result = client.poll(session)

        assertTrue(result is InteractiveAuthPollResult.Complete)
        val credentials = (result as InteractiveAuthPollResult.Complete).credentials
        assertEquals("person@example.com", credentials.username)
        assertEquals("refresh-xyz", credentials.password)
    }

    @Test
    fun `poll throws EXPIRED when the device code has expired`() = runTest {
        server.enqueue(MockResponse().setResponseCode(400).setBody("""{"error":"expired_token"}"""))
        val session = client.begin()
        // begin() above needs its own response; re-enqueue the expiry response for poll():
        server.enqueue(MockResponse().setResponseCode(400).setBody("""{"error":"expired_token"}"""))

        try {
            client.poll(session)
            org.junit.Assert.fail("expected InteractiveAuthException")
        } catch (e: InteractiveAuthException) {
            assertEquals(InteractiveAuthErrorKind.EXPIRED, e.kind)
        }
    }

    @Test
    fun `poll throws AUTHENTICATION when the user declines`() = runTest {
        server.enqueue(
            MockResponse().setBody(
                """
                {
                  "device_code": "device-abc",
                  "user_code": "XYZ123",
                  "verification_uri": "https://microsoft.com/devicelogin",
                  "verification_uri_complete": "https://microsoft.com/devicelogin?otc=XYZ123",
                  "expires_in": 900,
                  "interval": 5
                }
                """.trimIndent()
            )
        )
        val session = client.begin()
        server.enqueue(MockResponse().setResponseCode(400).setBody("""{"error":"authorization_declined"}"""))

        try {
            client.poll(session)
            org.junit.Assert.fail("expected InteractiveAuthException")
        } catch (e: InteractiveAuthException) {
            assertEquals(InteractiveAuthErrorKind.AUTHENTICATION, e.kind)
        }
    }
}
```

(The `poll returns Pending while authorization_pending` test body above is intentionally a no-op placeholder structure check — remove it and replace with the real assertion in Step 2 below once you see the real API shape; keep the other four tests as written.)

- [ ] **Step 2: Run it to verify it fails**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.auth.OneDriveDeviceCodeClientTest"`
Expected: compile failure — `OneDriveDeviceCodeClient` does not exist yet.

- [ ] **Step 3: Implement `OneDriveDeviceCodeClient`**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClient.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.auth

import com.dot.gallery.cloud.core.ProviderType
import com.dot.gallery.cloud.core.auth.InteractiveAuthCredentials
import com.dot.gallery.cloud.core.auth.InteractiveAuthErrorKind
import com.dot.gallery.cloud.core.auth.InteractiveAuthException
import com.dot.gallery.cloud.core.auth.InteractiveAuthPollResult
import com.dot.gallery.cloud.core.auth.InteractiveAuthSession
import com.google.gson.Gson
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import okhttp3.FormBody
import okhttp3.OkHttpClient
import okhttp3.Request
import java.io.IOException
import java.net.SocketTimeoutException
import java.net.UnknownHostException
import javax.net.ssl.SSLException

private data class DeviceCodeResponseDto(
    val device_code: String = "",
    val user_code: String = "",
    val verification_uri: String = "",
    val verification_uri_complete: String? = null,
    val expires_in: Long = 900L,
    val interval: Long = 5L
)

private data class TokenResponseDto(
    val access_token: String? = null,
    val refresh_token: String? = null,
    val expires_in: Long = 0L,
    val error: String? = null
)

private data class GraphMeDto(
    val userPrincipalName: String? = null,
    val mail: String? = null
)

/**
 * OAuth 2.0 device code flow (no MSAL, no redirect URI) — matches the polling shape of
 * [com.dot.gallery.cloud.core.auth.CloudInteractiveAuthHandler] exactly, the same way
 * Nextcloud's Login Flow v2 does.
 *
 * `scope` requests `offline_access` (to get a refresh token) + `Files.Read` (OneDrive
 * read access). The returned refresh token is stored by the caller as
 * [com.dot.gallery.cloud.core.CloudServerConfig.password]; the account's UPN as
 * [com.dot.gallery.cloud.core.CloudServerConfig.username].
 */
class OneDriveDeviceCodeClient(
    private val client: OkHttpClient,
    private val clientId: String,
    private val authBaseUrl: String = "https://login.microsoftonline.com/common/oauth2/v2.0",
    private val graphBaseUrl: String = "https://graph.microsoft.com/v1.0",
    private val nowMillis: () -> Long = System::currentTimeMillis
) {
    private val gson = Gson()

    suspend fun begin(): InteractiveAuthSession = withContext(Dispatchers.IO) {
        val body = FormBody.Builder()
            .add("client_id", clientId)
            .add("scope", "offline_access Files.Read")
            .build()
        val request = Request.Builder().url("$authBaseUrl/devicecode").post(body).build()

        executeSafely {
            client.newCall(request).execute().use { response ->
                throwForStatus(response.code)
                val payload = decode<DeviceCodeResponseDto>(response.body?.string().orEmpty())
                InteractiveAuthSession(
                    providerType = ProviderType.ONEDRIVE,
                    browserUrl = payload.verification_uri_complete ?: payload.verification_uri,
                    pollEndpoint = "$authBaseUrl/token",
                    token = payload.device_code,
                    trustedOrigin = authBaseUrl,
                    expiresAtMillis = nowMillis() + (payload.expires_in * 1000L)
                )
            }
        }
    }

    suspend fun poll(session: InteractiveAuthSession): InteractiveAuthPollResult = withContext(Dispatchers.IO) {
        if (nowMillis() >= session.expiresAtMillis) {
            throw InteractiveAuthException(InteractiveAuthErrorKind.EXPIRED, "OneDrive sign-in code expired")
        }
        val body = FormBody.Builder()
            .add("grant_type", "urn:ietf:params:oauth:grant-type:device_code")
            .add("client_id", clientId)
            .add("device_code", session.token)
            .build()
        val request = Request.Builder().url(session.pollEndpoint).post(body).build()

        executeSafely {
            client.newCall(request).execute().use { response ->
                val payload = decode<TokenResponseDto>(response.body?.string().orEmpty())
                when {
                    response.isSuccessful && payload.access_token != null -> {
                        val refreshToken = payload.refresh_token
                            ?: throw InteractiveAuthException(
                                InteractiveAuthErrorKind.MALFORMED_RESPONSE,
                                "OneDrive token response had no refresh_token"
                            )
                        val upn = fetchUserPrincipalName(payload.access_token)
                        InteractiveAuthPollResult.Complete(
                            InteractiveAuthCredentials(
                                serverUrl = "https://graph.microsoft.com/v1.0",
                                username = upn,
                                password = refreshToken
                            )
                        )
                    }
                    payload.error == "authorization_pending" || payload.error == "slow_down" ->
                        InteractiveAuthPollResult.Pending
                    payload.error == "expired_token" ->
                        throw InteractiveAuthException(InteractiveAuthErrorKind.EXPIRED, "OneDrive sign-in code expired")
                    payload.error == "authorization_declined" ->
                        throw InteractiveAuthException(InteractiveAuthErrorKind.AUTHENTICATION, "Sign-in was declined")
                    else -> {
                        throwForStatus(response.code)
                        throw InteractiveAuthException(InteractiveAuthErrorKind.SERVER, "Unexpected OneDrive token response")
                    }
                }
            }
        }
    }

    private fun fetchUserPrincipalName(accessToken: String): String {
        val request = Request.Builder()
            .url("$graphBaseUrl/me")
            .addHeader("Authorization", "Bearer $accessToken")
            .build()
        client.newCall(request).execute().use { response ->
            throwForStatus(response.code)
            val me = decode<GraphMeDto>(response.body?.string().orEmpty())
            return me.userPrincipalName ?: me.mail
                ?: throw InteractiveAuthException(
                    InteractiveAuthErrorKind.MALFORMED_RESPONSE,
                    "OneDrive account has no identifiable email"
                )
        }
    }

    private inline fun <T> executeSafely(block: () -> T): T = try {
        block()
    } catch (error: InteractiveAuthException) {
        throw error
    } catch (error: SSLException) {
        throw InteractiveAuthException(InteractiveAuthErrorKind.TLS, "Secure connection failed", error)
    } catch (error: UnknownHostException) {
        throw InteractiveAuthException(InteractiveAuthErrorKind.NETWORK, "Microsoft sign-in server was not found", error)
    } catch (error: SocketTimeoutException) {
        throw InteractiveAuthException(InteractiveAuthErrorKind.NETWORK, "Microsoft sign-in connection timed out", error)
    } catch (error: IOException) {
        throw InteractiveAuthException(InteractiveAuthErrorKind.NETWORK, "Microsoft sign-in connection failed", error)
    }

    private inline fun <reified T> decode(value: String): T = try {
        gson.fromJson(value, T::class.java)
    } catch (error: Exception) {
        throw InteractiveAuthException(
            InteractiveAuthErrorKind.MALFORMED_RESPONSE,
            "OneDrive returned an invalid response",
            error
        )
    }

    private fun throwForStatus(status: Int) {
        if (status in 200..299) return
        val kind = when (status) {
            401, 403 -> InteractiveAuthErrorKind.AUTHENTICATION
            429 -> InteractiveAuthErrorKind.RATE_LIMITED
            in 500..599 -> InteractiveAuthErrorKind.SERVER
            else -> InteractiveAuthErrorKind.SERVER
        }
        throw InteractiveAuthException(kind, "OneDrive sign-in request failed ($status)")
    }
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.auth.OneDriveDeviceCodeClientTest"`
Expected: `BUILD SUCCESSFUL`, 5 tests passed.

- [ ] **Step 5: Commit**

```bash
git add app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClient.kt app/src/test/java/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClientTest.kt
git commit -m "Add OneDrive OAuth device code auth client"
```

---

### Task 4: Auth interceptor (token refresh on 401) + Retrofit API service

**Files:**
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/api/OneDriveAuthInterceptor.kt`
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/api/OneDriveApiService.kt`
- Test: `app/src/test/java/com/dot/gallery/cloud/onedrive/data/api/OneDriveAuthInterceptorTest.kt`

**Interfaces:**
- Produces: `class OneDriveAuthInterceptor(refreshTokenBlocking: (String) -> String?) : Interceptor` with mutable `var accessToken: String?` and `var refreshToken: String?`; `interface OneDriveApiService` with `suspend fun getItem(itemId: String): Response<OneDriveItemDto>` and `suspend fun getQuota(): Response<OneDriveQuotaResponseDto>`.
- Consumes: `OneDriveItemDto` (Task 2).

- [ ] **Step 1: Write the failing test**

```kotlin
// app/src/test/java/com/dot/gallery/cloud/onedrive/data/api/OneDriveAuthInterceptorTest.kt
package com.dot.gallery.cloud.onedrive.data.api

import okhttp3.OkHttpClient
import okhttp3.Request
import okhttp3.mockwebserver.MockResponse
import okhttp3.mockwebserver.MockWebServer
import org.junit.After
import org.junit.Assert.assertEquals
import org.junit.Before
import org.junit.Test

class OneDriveAuthInterceptorTest {

    private lateinit var server: MockWebServer

    @Before
    fun setUp() {
        server = MockWebServer()
        server.start()
    }

    @After
    fun tearDown() {
        server.shutdown()
    }

    @Test
    fun `attaches the current access token as a Bearer header`() {
        var refreshCalls = 0
        val interceptor = OneDriveAuthInterceptor(refreshTokenBlocking = { refreshCalls++; "new-token" })
        interceptor.accessToken = "token-1"
        val client = OkHttpClient.Builder().addInterceptor(interceptor).build()
        server.enqueue(MockResponse().setBody("ok"))

        client.newCall(Request.Builder().url(server.url("/x")).build()).execute().use { }

        val recorded = server.takeRequest()
        assertEquals("Bearer token-1", recorded.getHeader("Authorization"))
        assertEquals(0, refreshCalls)
    }

    @Test
    fun `refreshes once and retries on a 401, then uses the new token on the next call`() {
        var refreshCalls = 0
        val interceptor = OneDriveAuthInterceptor(refreshTokenBlocking = { refreshCalls++; "refreshed-token" })
        interceptor.accessToken = "stale-token"
        interceptor.refreshToken = "refresh-xyz"
        val client = OkHttpClient.Builder().addInterceptor(interceptor).build()

        server.enqueue(MockResponse().setResponseCode(401))
        server.enqueue(MockResponse().setBody("ok"))

        client.newCall(Request.Builder().url(server.url("/x")).build()).execute().use { response ->
            assertEquals(200, response.code)
        }

        assertEquals(1, refreshCalls)
        assertEquals("refreshed-token", interceptor.accessToken)
        val firstAttempt = server.takeRequest()
        val retry = server.takeRequest()
        assertEquals("Bearer stale-token", firstAttempt.getHeader("Authorization"))
        assertEquals("Bearer refreshed-token", retry.getHeader("Authorization"))
    }

    @Test
    fun `a revoked refresh token surfaces the original 401 once, not an infinite retry loop`() {
        var refreshCalls = 0
        val interceptor = OneDriveAuthInterceptor(refreshTokenBlocking = { refreshCalls++; null })
        interceptor.accessToken = "stale-token"
        interceptor.refreshToken = "revoked-refresh-token"
        val client = OkHttpClient.Builder().addInterceptor(interceptor).build()

        server.enqueue(MockResponse().setResponseCode(401))

        client.newCall(Request.Builder().url(server.url("/x")).build()).execute().use { response ->
            assertEquals(401, response.code)
        }

        assertEquals(1, refreshCalls)
        assertEquals(1, server.requestCount)
    }
}
```

- [ ] **Step 2: Run it to verify it fails**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.data.api.OneDriveAuthInterceptorTest"`
Expected: compile failure — `OneDriveAuthInterceptor` does not exist yet.

- [ ] **Step 3: Implement the interceptor**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/api/OneDriveAuthInterceptor.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.data.api

import okhttp3.Interceptor
import okhttp3.Response

/**
 * Attaches the current access token as a Bearer header and, on a single 401, refreshes it
 * via [refreshTokenBlocking] (a synchronous call to the token endpoint using [refreshToken])
 * and retries exactly once. A second 401 after a successful refresh is returned as-is —
 * it means the refresh token itself is no longer valid (revoked/expired), which the caller
 * (`OneDriveProvider.authenticate`/callers checking the response) surfaces as a real
 * re-authentication error rather than looping forever.
 */
class OneDriveAuthInterceptor(
    private val refreshTokenBlocking: (refreshToken: String) -> String?
) : Interceptor {

    @Volatile
    var accessToken: String? = null

    @Volatile
    var refreshToken: String? = null

    override fun intercept(chain: Interceptor.Chain): Response {
        val original = chain.request()
        val withAuth = original.newBuilder()
            .apply { accessToken?.let { header("Authorization", "Bearer $it") } }
            .build()
        val response = chain.proceed(withAuth)

        if (response.code != 401) return response
        val currentRefreshToken = refreshToken ?: return response

        // Exactly one refresh attempt per 401. A null result means the refresh token
        // itself is no longer valid (revoked/expired) -- return the original 401 as-is
        // rather than retrying the same doomed request again.
        val newAccessToken = refreshTokenBlocking(currentRefreshToken) ?: return response
        response.close()
        accessToken = newAccessToken
        val retried = original.newBuilder()
            .header("Authorization", "Bearer $newAccessToken")
            .build()
        return chain.proceed(retried)
    }
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.data.api.OneDriveAuthInterceptorTest"`
Expected: `BUILD SUCCESSFUL`, 3 tests passed.

- [ ] **Step 5: Write `OneDriveApiService` (no new test — thin Retrofit interface, exercised indirectly by Task 5's provider tests)**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/api/OneDriveApiService.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.data.api

import com.dot.gallery.cloud.onedrive.data.dto.OneDriveItemDto
import com.google.gson.annotations.SerializedName
import retrofit2.Response
import retrofit2.http.GET
import retrofit2.http.Path
import retrofit2.http.Query

data class OneDriveQuotaDto(
    val used: Long = 0L,
    val total: Long = 0L
)

data class OneDriveQuotaResponseDto(
    val quota: OneDriveQuotaDto? = null
)

data class OneDriveSearchResponseDto(
    val value: List<OneDriveItemDto> = emptyList()
)

interface OneDriveApiService {
    @GET("me/drive/items/{itemId}")
    suspend fun getItem(@Path("itemId") itemId: String): Response<OneDriveItemDto>

    @GET("me/drive")
    suspend fun getQuota(): Response<OneDriveQuotaResponseDto>

    @GET("me/drive/root/search(q='{query}')")
    suspend fun search(@Path("query") query: String): Response<OneDriveSearchResponseDto>
}
```

- [ ] **Step 6: Verify the app still builds**

Run: `./gradlew :app:compileArm64-v8aNoMLDebugKotlin`
Expected: `BUILD SUCCESSFUL`.

- [ ] **Step 7: Commit**

```bash
git add app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/api/OneDriveAuthInterceptor.kt app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/api/OneDriveApiService.kt app/src/test/java/com/dot/gallery/cloud/onedrive/data/api/OneDriveAuthInterceptorTest.kt
git commit -m "Add OneDrive auth interceptor with 401 refresh-and-retry, and API service"
```

---

### Task 5: `OneDriveProvider` + auth handler + DI wiring + UI descriptor

**Files:**
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/OneDriveProvider.kt`
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveInteractiveAuthHandler.kt`
- Create: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/di/OneDriveModule.kt`
- Modify: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/dto/OneDriveDtos.kt` (Task 2's file — add `toCloudMediaEntity`)
- Modify: `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClient.kt` (Task 3's file — add `exchangeRefreshTokenBlocking`)
- Modify: `app/src/main/kotlin/com/dot/gallery/cloud/ui/descriptor/ProviderUiDescriptor.kt`
- Test: `app/src/test/java/com/dot/gallery/cloud/onedrive/OneDriveProviderTest.kt`

**Interfaces:**
- Consumes: `OneDriveAssetLister` (Task 2), `OneDriveDeviceCodeClient` (Task 3), `OneDriveAuthInterceptor`, `OneDriveApiService` (Task 4).
- Produces: `class OneDriveProvider : RemoteMediaProvider` registered in `ProviderRegistry` via Hilt multibinding, same pattern as `ImmichProvider`.

- [ ] **Step 1: Write the failing test for the capability surface and URL construction**

```kotlin
// app/src/test/java/com/dot/gallery/cloud/onedrive/OneDriveProviderTest.kt
package com.dot.gallery.cloud.onedrive

import com.dot.gallery.cloud.core.CloudServerConfig
import com.dot.gallery.cloud.core.ProviderCapability
import com.dot.gallery.cloud.core.ProviderType
import com.dot.gallery.cloud.core.ThumbnailSize
import com.dot.gallery.cloud.onedrive.auth.OneDriveDeviceCodeClient
import com.dot.gallery.cloud.onedrive.data.api.OneDriveAuthInterceptor
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test

class OneDriveProviderTest {

    private fun newProvider() = OneDriveProvider(
        authInterceptor = OneDriveAuthInterceptor { null },
        assetListerFactory = { _ -> throw NotImplementedError("not exercised by this test") },
        deviceCodeClientFactory = { OneDriveDeviceCodeClient(okhttp3.OkHttpClient(), "unused") }
    )

    @Test
    fun `declares only REMOTE_ASSETS and TEXT_SEARCH capabilities`() {
        val provider = newProvider()
        assertEquals(setOf(ProviderCapability.REMOTE_ASSETS, ProviderCapability.TEXT_SEARCH), provider.capabilities)
    }

    @Test
    fun `is not available before configure is called`() {
        assertFalse(newProvider().isAvailable)
    }

    @Test
    fun `thumbnail and original urls resolve relative to a configured item`() {
        val provider = newProvider()
        provider.configure(
            CloudServerConfig(
                id = 1L,
                providerType = ProviderType.ONEDRIVE,
                serverUrl = "https://graph.microsoft.com/v1.0",
                username = "person@example.com",
                password = "refresh-token"
            )
        )
        // getThumbnailUrl/getOriginalUrl resolve from the per-item cached download/thumbnail
        // URLs captured during listing, not from a URL template -- an item never listed
        // (or whose Graph download URL expired) returns an empty string rather than a
        // broken guess, matching how path-based providers signal "no preview available".
        assertEquals("", provider.getThumbnailUrl("unseen-item-id", ThumbnailSize.PREVIEW))
        assertEquals("", provider.getOriginalUrl("unseen-item-id"))
    }

    @Test
    fun `write operations outside this release's scope fail clearly instead of silently no-opping`() {
        val provider = newProvider()
        assertTrue(provider.capabilities.none { it == ProviderCapability.FAVORITE })
        // toggleFavorite is still required by the interface but must never be reachable from UI
        // since FAVORITE isn't declared; it fails loudly if ever called directly, rather than
        // quietly pretending to succeed.
        val result = kotlinx.coroutines.runBlocking { provider.toggleFavorite("x", true) }
        assertTrue(result.isFailure)
        assertTrue(result.exceptionOrNull() is UnsupportedOperationException)
    }
}
```

- [ ] **Step 2: Run it to verify it fails**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.OneDriveProviderTest"`
Expected: compile failure — `OneDriveProvider` does not exist yet.

- [ ] **Step 3: Implement `OneDriveProvider`**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/OneDriveProvider.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive

import com.dot.gallery.cloud.core.CloudAlbum
import com.dot.gallery.cloud.core.CloudAuthToken
import com.dot.gallery.cloud.core.CloudServerConfig
import com.dot.gallery.cloud.core.CloudServerInfo
import com.dot.gallery.cloud.core.CloudStorageInfo
import com.dot.gallery.cloud.core.ConnectionState
import com.dot.gallery.cloud.core.ProviderCapability
import com.dot.gallery.cloud.core.ProviderType
import com.dot.gallery.cloud.core.ThumbnailSize
import com.dot.gallery.cloud.core.capabilities.RemoteMediaProvider
import com.dot.gallery.cloud.core.capabilities.SyncDelta
import com.dot.gallery.cloud.data.entity.CloudMediaEntity
import com.dot.gallery.cloud.onedrive.auth.OneDriveDeviceCodeClient
import com.dot.gallery.cloud.onedrive.data.api.OneDriveAuthInterceptor
import com.dot.gallery.core.Resource
import com.dot.gallery.feature_node.domain.model.Media
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.flow
import okhttp3.OkHttpClient
import okhttp3.Request
import java.util.concurrent.ConcurrentHashMap
import javax.inject.Inject
import javax.inject.Singleton

/**
 * Read-only OneDrive viewing provider: sign in (device code flow), list every photo in the
 * drive tree, show thumbnails and full-res images. Write operations (favorite/trash/albums/
 * sync) are explicitly out of scope for this release -- see the design spec.
 */
@Singleton
class OneDriveProvider @Inject constructor(
    private val authInterceptor: OneDriveAuthInterceptor,
    private val assetListerFactory: (OkHttpClient) -> OneDriveAssetLister = { OneDriveAssetLister(it) },
    private val deviceCodeClientFactory: () -> OneDriveDeviceCodeClient
) : RemoteMediaProvider {

    override val providerType = ProviderType.ONEDRIVE
    override val displayName = "OneDrive"

    override val capabilities: Set<ProviderCapability> = setOf(
        ProviderCapability.REMOTE_ASSETS,
        ProviderCapability.TEXT_SEARCH
    )

    private val _connectionState = MutableStateFlow(ConnectionState.DISCONNECTED)
    override val connectionState: StateFlow<ConnectionState> = _connectionState.asStateFlow()

    private var currentConfig: CloudServerConfig? = null
    private val httpClient = OkHttpClient.Builder().addInterceptor(authInterceptor).build()
    private val assetLister by lazy { assetListerFactory(httpClient) }

    // Graph's @microsoft.graph.downloadUrl / thumbnail URL are per-item, time-limited, and
    // only known once an item has been listed -- there is no stable URL template to construct
    // them from an id alone, unlike path-based providers. Cached from the last listing.
    private val downloadUrlsByRemoteId = ConcurrentHashMap<String, String>()
    private val thumbnailUrlsByRemoteId = ConcurrentHashMap<String, String>()

    override val isAvailable: Boolean
        get() = currentConfig != null && _connectionState.value == ConnectionState.CONNECTED

    override fun configure(config: CloudServerConfig) {
        currentConfig = config
        authInterceptor.refreshToken = config.password
    }

    private fun requireConfig(): CloudServerConfig = currentConfig
        ?: throw IllegalStateException("OneDriveProvider not configured. Call configure() first.")

    override suspend fun testConnection(config: CloudServerConfig): Result<CloudServerInfo> = runCatching {
        val request = Request.Builder()
            .url("${config.serverUrl}/me")
            .addHeader("Authorization", "Bearer ${refreshAccessToken(config.password.orEmpty())}")
            .build()
        OkHttpClient().newCall(request).execute().use { response ->
            if (!response.isSuccessful) throw Exception("Connection failed: ${response.code}")
        }
        CloudServerInfo(version = "v1.0", serverName = "OneDrive")
    }

    override suspend fun authenticate(config: CloudServerConfig): Result<CloudAuthToken> = runCatching {
        val refreshToken = config.password
            ?: return@runCatching throw IllegalArgumentException("No OneDrive refresh token stored")
        authInterceptor.refreshToken = refreshToken
        authInterceptor.accessToken = refreshAccessToken(refreshToken)
        _connectionState.value = ConnectionState.CONNECTED
        CloudAuthToken(accessToken = authInterceptor.accessToken.orEmpty(), userEmail = config.username)
    }.onFailure { _connectionState.value = ConnectionState.ERROR }

    private fun refreshAccessToken(refreshToken: String): String =
        deviceCodeClientFactory().let { client ->
            // The device code client doesn't expose a standalone refresh-grant call in Task 3;
            // OneDriveAuthInterceptor's refreshTokenBlocking (wired in the DI module, Step 5
            // below) is the actual production refresh path. This direct call only covers the
            // first authenticate()/testConnection() before any request has gone through the
            // interceptor.
            client.exchangeRefreshTokenBlocking(refreshToken)
        }

    override fun getRemoteAssets(page: Int, pageSize: Int): Flow<Resource<List<CloudMediaEntity>>> = flow {
        try {
            val config = requireConfig()
            val items = assetLister.listAll(authInterceptor.accessToken.orEmpty())
            items.forEach { item ->
                item.downloadUrl?.let { downloadUrlsByRemoteId[item.id] = it }
                item.thumbnails?.firstOrNull()?.medium?.url?.let { thumbnailUrlsByRemoteId[item.id] = it }
            }
            val entities = items.map { it.toCloudMediaEntity(config.id) }
            emit(Resource.Success(entities))
        } catch (e: Exception) {
            emit(Resource.Error(e.message ?: "Unknown error"))
        }
    }

    override fun getRemoteAlbums(): Flow<Resource<List<CloudAlbum>>> = flow { emit(Resource.Success(emptyList())) }
    override fun getRemoteAlbumMedia(albumId: String): Flow<Resource<List<CloudMediaEntity>>> =
        flow { emit(Resource.Success(emptyList())) }
    override fun getRemoteFavorites(): Flow<Resource<List<CloudMediaEntity>>> =
        flow { emit(Resource.Success(emptyList())) }
    override fun getRemoteTrashed(): Flow<Resource<List<CloudMediaEntity>>> =
        flow { emit(Resource.Success(emptyList())) }
    override fun getRemoteArchived(): Flow<Resource<List<CloudMediaEntity>>> =
        flow { emit(Resource.Success(emptyList())) }

    override suspend fun toggleFavorite(remoteId: String, favorite: Boolean): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: favorites are not supported in this release"))
    override suspend fun toggleArchive(remoteId: String, archived: Boolean): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: archive is not supported in this release"))
    override suspend fun restoreAsset(remoteId: String): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: trash/restore is not supported in this release"))
    override suspend fun trashAsset(remoteId: String): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: trash is not supported in this release"))
    override suspend fun emptyTrash(): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: trash is not supported in this release"))
    override suspend fun restoreAllTrash(): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: trash is not supported in this release"))
    override suspend fun createAlbum(name: String): Result<CloudAlbum> =
        Result.failure(UnsupportedOperationException("OneDrive: albums are not supported in this release"))
    override suspend fun addToAlbum(albumId: String, assetIds: List<String>): Result<Unit> =
        Result.failure(UnsupportedOperationException("OneDrive: albums are not supported in this release"))

    override suspend fun deleteAsset(remoteId: String): Result<Unit> = runCatching {
        val request = Request.Builder()
            .url("${requireConfig().serverUrl}/me/drive/items/$remoteId")
            .delete()
            .build()
        httpClient.newCall(request).execute().use { response ->
            if (!response.isSuccessful) throw Exception("Delete failed: ${response.code}")
        }
        downloadUrlsByRemoteId.remove(remoteId)
        thumbnailUrlsByRemoteId.remove(remoteId)
    }

    override suspend fun search(query: String): Result<List<CloudMediaEntity>> = runCatching {
        val config = requireConfig()
        val request = Request.Builder()
            .url("${config.serverUrl}/me/drive/root/search(q='$query')")
            .addHeader("Authorization", "Bearer ${authInterceptor.accessToken}")
            .build()
        httpClient.newCall(request).execute().use { response ->
            if (!response.isSuccessful) throw Exception("Search failed: ${response.code}")
            val body = response.body?.string().orEmpty()
            com.google.gson.Gson()
                .fromJson(body, com.dot.gallery.cloud.onedrive.data.api.OneDriveSearchResponseDto::class.java)
                .value
                .filter { it.isImage }
                .map { it.toCloudMediaEntity(config.id) }
        }
    }

    override fun getThumbnailUrl(remoteId: String, size: ThumbnailSize): String =
        thumbnailUrlsByRemoteId[remoteId].orEmpty()

    override fun getOriginalUrl(remoteId: String): String = downloadUrlsByRemoteId[remoteId].orEmpty()

    override fun getAuthHeaders(): Map<String, String> =
        authInterceptor.accessToken?.let { mapOf("Authorization" to "Bearer $it") } ?: emptyMap()

    override suspend fun getStorageInfo(): Result<CloudStorageInfo> = runCatching {
        val request = Request.Builder()
            .url("${requireConfig().serverUrl}/me/drive")
            .addHeader("Authorization", "Bearer ${authInterceptor.accessToken}")
            .build()
        httpClient.newCall(request).execute().use { response ->
            if (!response.isSuccessful) throw Exception("Storage info failed: ${response.code}")
            val body = response.body?.string().orEmpty()
            val dto = com.google.gson.Gson()
                .fromJson(body, com.dot.gallery.cloud.onedrive.data.api.OneDriveQuotaResponseDto::class.java)
            val used = dto.quota?.used ?: 0L
            val total = dto.quota?.total ?: 0L
            CloudStorageInfo(
                usedBytes = used,
                totalBytes = total,
                usedPercentage = if (total > 0L) used.toDouble() / total * 100.0 else 0.0
            )
        }
    }
}
```

**Note on `exchangeRefreshTokenBlocking`:** Task 3's `OneDriveDeviceCodeClient` only has `begin`/`poll`. Add a small additional method to it in this task (not a new file — editing the Task 3 file) since `OneDriveProvider` and `OneDriveAuthInterceptor`'s refresh callback both need a plain "refresh token -> new access token" call:

```kotlin
    // Added to OneDriveDeviceCodeClient (app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveDeviceCodeClient.kt)
    fun exchangeRefreshTokenBlocking(refreshToken: String): String {
        val body = FormBody.Builder()
            .add("grant_type", "refresh_token")
            .add("client_id", clientId)
            .add("refresh_token", refreshToken)
            .add("scope", "offline_access Files.Read")
            .build()
        val request = Request.Builder().url("$authBaseUrl/token").post(body).build()
        client.newCall(request).execute().use { response ->
            val payload = decode<TokenResponseDto>(response.body?.string().orEmpty())
            return payload.access_token
                ?: throw InteractiveAuthException(InteractiveAuthErrorKind.AUTHENTICATION, "OneDrive refresh token is no longer valid")
        }
    }
```

- [ ] **Step 4: Add `OneDriveItemDto.toCloudMediaEntity`**

Add this extension function to `app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/data/dto/OneDriveDtos.kt` (Task 2's file):

```kotlin
fun OneDriveItemDto.toCloudMediaEntity(serverConfigId: Long): com.dot.gallery.cloud.data.entity.CloudMediaEntity {
    val mimeType = file?.mimeType ?: "image/jpeg"
    return com.dot.gallery.cloud.data.entity.CloudMediaEntity(
        remoteId = id,
        providerType = com.dot.gallery.cloud.core.ProviderType.ONEDRIVE,
        serverConfigId = serverConfigId,
        label = name,
        relativePath = parentReference?.path.orEmpty(),
        mimeType = mimeType,
        size = size,
        contentHash = file?.hashes?.quickXorHash ?: file?.hashes?.sha256Hash,
        thumbnailUrl = thumbnails?.firstOrNull()?.medium?.url.orEmpty(),
        originalUrl = downloadUrl.orEmpty()
    )
}
```

- [ ] **Step 5: Implement `OneDriveInteractiveAuthHandler` and the DI module**

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/auth/OneDriveInteractiveAuthHandler.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.auth

import com.dot.gallery.cloud.core.CloudServerConfig
import com.dot.gallery.cloud.core.ProviderType
import com.dot.gallery.cloud.core.auth.CloudInteractiveAuthHandler
import com.dot.gallery.cloud.core.auth.InteractiveAuthPollResult
import com.dot.gallery.cloud.core.auth.InteractiveAuthSession
import javax.inject.Inject

class OneDriveInteractiveAuthHandler @Inject constructor(
    private val client: OneDriveDeviceCodeClient
) : CloudInteractiveAuthHandler {
    override val providerType: ProviderType = ProviderType.ONEDRIVE

    override suspend fun begin(serverUrl: String): InteractiveAuthSession = client.begin()

    override suspend fun poll(session: InteractiveAuthSession): InteractiveAuthPollResult = client.poll(session)

    // Microsoft identity platform has no public per-refresh-token revoke endpoint reachable
    // from a native app (unlike Nextcloud's app-password DELETE); disconnecting locally
    // (clearing the stored refresh token) is the only action available here.
    override suspend fun revoke(config: CloudServerConfig): Result<Unit> = Result.success(Unit)
}
```

```kotlin
// app/src/onedrive/kotlin/com/dot/gallery/cloud/onedrive/di/OneDriveModule.kt
/*
 * SPDX-FileCopyrightText: 2023-2026 IacobIacob01
 * SPDX-License-Identifier: Apache-2.0
 */

package com.dot.gallery.cloud.onedrive.di

import com.dot.gallery.BuildConfig
import com.dot.gallery.cloud.core.MediaCapabilityProvider
import com.dot.gallery.cloud.core.ProviderInstanceFactory
import com.dot.gallery.cloud.core.ProviderType
import com.dot.gallery.cloud.onedrive.OneDriveProvider
import com.dot.gallery.cloud.onedrive.auth.OneDriveDeviceCodeClient
import com.dot.gallery.cloud.onedrive.data.api.OneDriveAuthInterceptor
import dagger.Module
import dagger.Provides
import dagger.hilt.InstallIn
import dagger.hilt.components.SingletonComponent
import dagger.multibindings.IntoSet
import okhttp3.OkHttpClient
import javax.inject.Singleton

// This app's own Azure registration (see docs/superpowers/specs). A public-client
// registration's client id is not a secret -- it identifies the app, not a user.
private const val ONEDRIVE_CLIENT_ID = "c623491e-a0a4-4e9e-9732-9d86d9733352"

@Module
@InstallIn(SingletonComponent::class)
object OneDriveModule {

    @Provides
    @Singleton
    @IntoSet
    fun provideOneDriveProviderFactory(): ProviderInstanceFactory = object : ProviderInstanceFactory {
        override val providerType = ProviderType.ONEDRIVE
        override fun create(): MediaCapabilityProvider {
            val deviceCodeClient = OneDriveDeviceCodeClient(OkHttpClient(), ONEDRIVE_CLIENT_ID)
            val authInterceptor = OneDriveAuthInterceptor { refreshToken ->
                runCatching { deviceCodeClient.exchangeRefreshTokenBlocking(refreshToken) }.getOrNull()
            }
            return OneDriveProvider(
                authInterceptor = authInterceptor,
                deviceCodeClientFactory = { deviceCodeClient }
            )
        }
    }
}
```

- [ ] **Step 6: Add the `ProviderUiDescriptor` entry**

In `app/src/main/kotlin/com/dot/gallery/cloud/ui/descriptor/ProviderUiDescriptor.kt`, inside the `descriptors` `buildMap`, add (mirrors the Nextcloud entry's interactive-auth shape — the username/password fields are auto-filled by the device-code flow, not meant for manual typing):

```kotlin
        put(
            ProviderType.ONEDRIVE,
            ProviderUiDescriptor(
                providerType = ProviderType.ONEDRIVE,
                category = ProviderCategory.STANDALONE,
                icon = Icons.Outlined.Cloud,
                iconRes = R.drawable.ic_provider_azure,
                iconTinted = false,
                urlRegex = Regex("^https://graph\\.microsoft\\.com.*"),
                urlHintRes = R.string.cloud_server_url_hint,
                credentialFields = listOf(usernameField, passwordField),
                credentialsSatisfied = { it.username.isNotBlank() && it.password.isNotBlank() }
            )
        )
```

- [ ] **Step 7: Run the tests to verify they pass**

Run: `./gradlew testDebugUnitTest --tests "com.dot.gallery.cloud.onedrive.OneDriveProviderTest"`
Expected: `BUILD SUCCESSFUL`, 4 tests passed.

- [ ] **Step 8: Run the full unit test suite**

Run: `./gradlew testDebugUnitTest`
Expected: `BUILD SUCCESSFUL`, all tests pass (existing + new OneDrive tests).

- [ ] **Step 9: Commit**

```bash
git add app/src/onedrive app/src/main/kotlin/com/dot/gallery/cloud/ui/descriptor/ProviderUiDescriptor.kt app/src/test/java/com/dot/gallery/cloud/onedrive/OneDriveProviderTest.kt
git commit -m "Add OneDriveProvider, auth handler, DI wiring, and UI descriptor"
```

---

### Task 6: On-device verification

**Files:** none (manual verification only)

- [ ] **Step 1: Confirm the Azure app registration allows public client flows**

In Azure Portal, App registrations → the existing `c623491e-...` registration → Authentication → "Advanced settings" → "Allow public client flows" must be **Yes**. Device code flow fails with an `unauthorized_client` error otherwise. Toggle it on if it isn't already (device code flow was never used by this registration before, since Phase 1–3 of `onedrive-gallery` used MSAL's browser redirect flow instead).

- [ ] **Step 2: Build and install**

```bash
./gradlew :app:assembleArm64-v8aNoMLDebug
adb install -r app/build/outputs/apk/arm64-v8aNoML/debug/*.apk
```

- [ ] **Step 3: Add the OneDrive account through the app's UI**

Launch the app → Settings → Cloud accounts (or wherever `CloudAddServerScreen` is reached) → select OneDrive → the app should show a "sign in" prompt driven by `begin()`/`poll()`. Open the shown URL (or it opens automatically), sign in with your Microsoft account, approve.

- [ ] **Step 4: Confirm photos appear**

Back in the app, OneDrive photos should appear in the gallery grid within a few seconds (the `CloudProviderInitializer` prefetch). Tap one to confirm the full-resolution viewer loads (uses `getOriginalUrl`).

- [ ] **Step 5: Report back**

This step has no code — report what you saw (sign-in worked? photos appeared? any errors in `adb logcat`?) so any issues get fixed in a follow-up, same as the earlier `onedrive-gallery` phases.
