package app.cells

import android.content.Intent
import android.graphics.Color
import android.net.Uri
import android.os.Bundle
import android.provider.OpenableColumns
import android.view.View
import android.webkit.JavascriptInterface
import android.webkit.WebView
import androidx.activity.OnBackPressedCallback
import androidx.activity.SystemBarStyle
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContract
import androidx.activity.result.contract.ActivityResultContracts
import android.app.Activity
import android.content.Context
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import org.json.JSONArray
import org.json.JSONObject

// After Inkspire's (src-tauri/gen/android/.../MainActivity.kt on its android branch).
class MainActivity : TauriActivity() {
  private var webView: WebView? = null
  private var barStrip: View? = null

  /** Files handed to Cells by other apps ("Open with", Share), until the page takes them. */
  internal val opened = mutableListOf<String>()

  // Open and Save As go through the system's document picker with read and write access that lasts across
  // restarts, so Save writes back to the file and Recent can open it again. (The dialog plugin's picker grants a
  // read-only copy.)
  private val openDocument =
    registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri: Uri? -> picked(uri) }
  private val createDocument =
    registerForActivityResult(CreateTypedDocument()) { uri: Uri? -> picked(uri) }

  private fun picked(uri: Uri?) {
    if (uri != null) keep(uri)
    val value = if (uri == null) "null" else JSONObject.quote(uri.toString())
    webView?.post { webView?.evaluateJavascript("window.__cellsPicked && window.__cellsPicked($value)", null) }
  }

  /** Asks to keep access to a document after a restart; a provider that offers none still opens it now. */
  private fun keep(uri: Uri) {
    try {
      contentResolver.takePersistableUriPermission(
        uri,
        Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION,
      )
    } catch (error: SecurityException) {
      try {
        contentResolver.takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION)
      } catch (error: SecurityException) {
        // Access for this session only.
      }
    }
  }

  override fun onCreate(savedInstanceState: Bundle?) {
    // Light status and navigation icons, over the green strip below.
    enableEdgeToEdge(
      statusBarStyle = SystemBarStyle.dark(Color.TRANSPARENT),
      navigationBarStyle = SystemBarStyle.dark(Color.TRANSPARENT),
    )
    super.onCreate(savedInstanceState)
    // Android 15 draws apps under the status and navigation bars. Keep Cells clear of them, and of the on-screen
    // keyboard, so the cell being typed in stays above it. The strip behind the status bar takes the title bar's
    // green (--app-title-fill).
    val content = findViewById<View>(android.R.id.content)
    content.setBackgroundColor(Color.parseColor("#217346"))
    barStrip = content
    ViewCompat.setOnApplyWindowInsetsListener(content) { view, insets ->
      val bars = insets.getInsets(WindowInsetsCompat.Type.systemBars() or WindowInsetsCompat.Type.displayCutout())
      val keyboard = insets.getInsets(WindowInsetsCompat.Type.ime())
      view.setPadding(bars.left, bars.top, bars.right, maxOf(bars.bottom, keyboard.bottom))
      // The page's height shrinks by the keyboard; it adds it back when choosing the phone or tablet layout
      // (app/src/phone/shell-choice.ts), so opening the keyboard never switches layouts.
      val covered = maxOf(0, keyboard.bottom - bars.bottom) / resources.displayMetrics.density
      webView?.post {
        webView?.evaluateJavascript(
          "window.__cellsKeyboardHeight = $covered; window.dispatchEvent(new Event('resize'))",
          null,
        )
      }
      WindowInsetsCompat.CONSUMED
    }
    if (savedInstanceState == null) receive(intent)
  }

  override fun onNewIntent(intent: Intent) {
    super.onNewIntent(intent)
    receive(intent)
  }

  /** A file from another app: View ("Open with") carries it as the data, Share as a stream. */
  private fun receive(intent: Intent?) {
    val uris = mutableListOf<Uri>()
    when (intent?.action) {
      Intent.ACTION_VIEW -> intent.data?.let { uris.add(it) }
      Intent.ACTION_SEND -> streamOf(intent)?.let { uris.add(it) }
      Intent.ACTION_SEND_MULTIPLE -> streamsOf(intent).forEach { uris.add(it) }
    }
    // Only content links: a bare file:// path is outside what Cells may read on Android.
    val files = uris.filter { it.scheme == "content" }
    if (files.isEmpty()) return
    for (uri in files) {
      keep(uri)
      synchronized(opened) { opened.add(uri.toString()) }
    }
    webView?.post { webView?.evaluateJavascript("window.__cellsOpened && window.__cellsOpened()", null) }
  }

  @Suppress("DEPRECATION")
  private fun streamOf(intent: Intent): Uri? = intent.getParcelableExtra(Intent.EXTRA_STREAM)

  @Suppress("DEPRECATION")
  private fun streamsOf(intent: Intent): List<Uri> =
    intent.getParcelableArrayListExtra<Uri>(Intent.EXTRA_STREAM) ?: emptyList()

  override fun onWebViewCreate(webView: WebView) {
    this.webView = webView
    webView.addJavascriptInterface(AndroidFiles(this), "CellsFiles")
    webView.addJavascriptInterface(AndroidOcr(this, webView), "CellsOcr")
    // Back closes what is open over the sheet (a panel, dialog, menu, the File view or the cell being edited), as Esc
    // does; with nothing open, Cells goes to the background instead of ending with unsaved work. Added after Wry's
    // own handler (web history), so it comes first.
    onBackPressedDispatcher.addCallback(
      this,
      object : OnBackPressedCallback(true) {
        override fun handleOnBackPressed() {
          webView.evaluateJavascript("window.__cellsBack ? window.__cellsBack() : false") { handled ->
            if (handled != "true") moveTaskToBack(true)
          }
        }
      },
    )
    // Pinching zooms the sheet (the grid handles two fingers), never the whole interface.
    webView.settings.setSupportZoom(false)
    webView.settings.builtInZoomControls = false
    webView.settings.displayZoomControls = false
  }

  /** The strip behind the system bars follows the title bar (dark mode has a darker green). */
  fun setBarColor(color: String) {
    val parsed = try { Color.parseColor(color) } catch (error: IllegalArgumentException) { return }
    runOnUiThread { barStrip?.setBackgroundColor(parsed) }
  }

  fun pickOpen(mimeTypes: Array<String>) {
    runOnUiThread { openDocument.launch(mimeTypes) }
  }

  fun pickSave(name: String, mimeType: String) {
    runOnUiThread { createDocument.launch(name to mimeType) }
  }
}

/** Save As: a new document with this name and MIME type (the stock contract fixes the type once). */
class CreateTypedDocument : ActivityResultContract<Pair<String, String>, Uri?>() {
  override fun createIntent(context: Context, input: Pair<String, String>): Intent =
    Intent(Intent.ACTION_CREATE_DOCUMENT)
      .addCategory(Intent.CATEGORY_OPENABLE)
      .setType(input.second)
      .putExtra(Intent.EXTRA_TITLE, input.first)

  override fun parseResult(resultCode: Int, intent: Intent?): Uri? =
    if (resultCode == Activity.RESULT_OK) intent?.data else null
}

/** What the page asks of Android's files: `window.CellsFiles` (app/src/platform/android-files.ts). */
class AndroidFiles(private val activity: MainActivity) {
  /**
   * The name a document shows in the file manager. A content link's last part is often only an id (a Downloads
   * file is "msf:1000000018"); the content resolver knows the real name.
   */
  @JavascriptInterface
  fun displayName(uri: String): String? =
    try {
      activity.contentResolver
        .query(Uri.parse(uri), arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)
        ?.use { cursor -> if (cursor.moveToFirst()) cursor.getString(0) else null }
    } catch (error: Exception) {
      null
    }

  /** Whether the document is still there and Cells may still read it (Recent asks before opening). */
  @JavascriptInterface
  fun exists(uri: String): Boolean =
    try {
      activity.contentResolver.openFileDescriptor(Uri.parse(uri), "r")?.use { true } ?: false
    } catch (error: Exception) {
      false
    }

  /** Shows the Open picker for these MIME types (a JSON array); the choice comes back to `__cellsPicked`. */
  @JavascriptInterface
  fun openDocument(mimeTypesJson: String) {
    val list = JSONArray(mimeTypesJson)
    val types = Array(list.length()) { list.getString(it) }
    activity.pickOpen(if (types.isEmpty()) arrayOf("*/*") else types)
  }

  /** Shows the Save As picker with this name and type; the new document comes back to `__cellsPicked`. */
  @JavascriptInterface
  fun createDocument(name: String, mimeType: String) {
    activity.pickSave(name, mimeType.ifEmpty { "application/octet-stream" })
  }

  /** The colour behind the status and navigation bars, as `#rrggbb` (the title bar's). */
  @JavascriptInterface
  fun setBarColor(color: String) {
    activity.setBarColor(color)
  }

  /** The files other apps handed over since the last call, as a JSON array. */
  @JavascriptInterface
  fun takeOpened(): String =
    synchronized(activity.opened) {
      val list = JSONArray(activity.opened)
      activity.opened.clear()
      list.toString()
    }
}
