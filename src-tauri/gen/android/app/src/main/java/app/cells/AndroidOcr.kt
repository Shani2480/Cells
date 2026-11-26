package app.cells

import android.content.Context
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.graphics.Color
import android.graphics.Matrix
import android.graphics.Rect
import android.graphics.pdf.PdfRenderer
import android.net.Uri
import android.os.ParcelFileDescriptor
import android.util.Base64
import android.webkit.JavascriptInterface
import android.webkit.WebView
import com.google.mlkit.vision.common.InputImage
import com.google.mlkit.vision.text.Text
import com.google.mlkit.vision.text.TextRecognition
import com.google.mlkit.vision.text.latin.TextRecognizerOptions
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.util.concurrent.Executors
import kotlin.math.roundToInt

/**
 * Text recognition for scanned PDF pages (opt-in in Cells' settings): Google's ML Kit with its Latin model bundled in
 * the app, so it runs on this device without Google Play services or a download, and nothing leaves the device. ML
 * Kit picks the device's fastest way to run its model itself. `window.CellsOcr` (app/src/pdf/ocr-android.ts) hands
 * over the PDF once (`openPdf`), and Android's own PDF renderer draws each scanned page for ML Kit (much faster than
 * drawing it in the web view and sending a picture); a picture can be read too (`recognize`). The words come back to
 * `window.__cellsOcr(id, json)`, as lines with their words' boxes in pixels and a confidence from 0 to 100 (the shape
 * Tesseract's are read in), and each line's angle.
 */
class AndroidOcr(private val context: Context, private val webView: WebView) {
  private val recognizer by lazy { TextRecognition.getClient(TextRecognizerOptions.DEFAULT_OPTIONS) }
  /** Pages are drawn one at a time, away from the page's bridge thread. */
  private val drawing = Executors.newSingleThreadExecutor()
  /** The PDF being read, a copy in the app's cache until `closePdf` (only Cells can read it there). */
  private val source get() = File(context.cacheDir, "ocr-source.pdf")

  /** Whether recognition can run here (the page asks before offering it). */
  @JavascriptInterface
  fun available(): Boolean = true

  /** Starts reading a page; `jpeg` is base64. The answer (or `{"error": …}`) comes back to `__cellsOcr`. */
  @JavascriptInterface
  fun recognize(id: String, jpeg: String) {
    val bitmap =
      try {
        val bytes = Base64.decode(jpeg, Base64.DEFAULT)
        BitmapFactory.decodeByteArray(bytes, 0, bytes.size)
      } catch (error: IllegalArgumentException) {
        null
      }
    if (bitmap == null) return answer(id, JSONObject().put("error", "The page couldn't be drawn for recognition."))
    read(id, bitmap)
  }

  /**
   * Starts reading a picture file's bytes (base64; JPEG, PNG, WebP, HEIC…), for Data › From Picture. ML Kit opens it
   * itself, so a photo is turned upright as its camera noted. The copy in the cache is removed once it's read.
   */
  @JavascriptInterface
  fun recognizeImage(id: String, image: String) {
    drawing.execute {
      val file = File(context.cacheDir, "ocr-picture-$id")
      val input =
        try {
          file.writeBytes(Base64.decode(image, Base64.DEFAULT))
          InputImage.fromFilePath(context, Uri.fromFile(file))
        } catch (error: Exception) {
          null
        }
      if (input == null) {
        file.delete()
        answer(id, JSONObject().put("error", "Cells can't open this kind of picture. Try a PNG or JPEG."))
      } else
        recognizer
          .process(input)
          .addOnSuccessListener { text -> answer(id, JSONObject().put("lines", lines(text))) }
          .addOnFailureListener { error ->
            answer(id, JSONObject().put("error", error.message ?: "Text recognition failed."))
          }
          .addOnCompleteListener { file.delete() }
    }
  }

  /** Keeps the PDF (base64) to draw its pages from. */
  @JavascriptInterface
  fun openPdf(pdf: String): Boolean =
    try {
      source.writeBytes(Base64.decode(pdf, Base64.DEFAULT))
      true
    } catch (error: Exception) {
      false
    }

  /** Removes the kept copy of the PDF. */
  @JavascriptInterface
  fun closePdf() {
    drawing.execute { source.delete() }
  }

  /** Starts reading page `index` (from 0) of the kept PDF, drawn at `scale` pixels a point. */
  @JavascriptInterface
  fun recognizePdfPage(id: String, index: Int, scale: Double) {
    drawing.execute {
      val bitmap =
        try {
          draw(index, scale)
        } catch (error: Exception) {
          null
        }
      if (bitmap == null) answer(id, JSONObject().put("error", "The page couldn't be drawn for recognition."))
      else read(id, bitmap)
    }
  }

  private fun draw(index: Int, scale: Double): Bitmap {
    val fd = ParcelFileDescriptor.open(source, ParcelFileDescriptor.MODE_READ_ONLY)
    val renderer = PdfRenderer(fd)
    try {
      val page = renderer.openPage(index)
      try {
        val bitmap =
          Bitmap.createBitmap(
            (page.width * scale).roundToInt(),
            (page.height * scale).roundToInt(),
            Bitmap.Config.ARGB_8888,
          )
        bitmap.eraseColor(Color.WHITE)
        val matrix = Matrix().apply { setScale(scale.toFloat(), scale.toFloat()) }
        page.render(bitmap, null, matrix, PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY)
        return bitmap
      } finally {
        page.close()
      }
    } finally {
      renderer.close()
      fd.close()
    }
  }

  private fun read(id: String, bitmap: Bitmap) {
    recognizer
      .process(InputImage.fromBitmap(bitmap, 0))
      .addOnSuccessListener { text -> answer(id, JSONObject().put("lines", lines(text))) }
      .addOnFailureListener { error ->
        answer(id, JSONObject().put("error", error.message ?: "Text recognition failed."))
      }
      .addOnCompleteListener { bitmap.recycle() }
  }

  private fun lines(text: Text): JSONArray {
    val out = JSONArray()
    for (block in text.textBlocks)
      for (line in block.lines) {
        val box = line.boundingBox ?: continue
        val words = JSONArray()
        for (element in line.elements) {
          val wordBox = element.boundingBox ?: continue
          words.put(
            JSONObject()
              .put("text", element.text)
              .put("confidence", (element.confidence * 100).toDouble())
              .put("bbox", bbox(wordBox))
          )
        }
        // The line's slant (degrees clockwise): the page uses it to straighten a scan before finding columns.
        out.put(JSONObject().put("bbox", bbox(box)).put("words", words).put("angle", line.angle.toDouble()))
      }
    return out
  }

  private fun bbox(r: Rect) = JSONObject().put("x0", r.left).put("y0", r.top).put("x1", r.right).put("y1", r.bottom)

  private fun answer(id: String, result: JSONObject) {
    val call = "window.__cellsOcr && window.__cellsOcr(${JSONObject.quote(id)}, ${JSONObject.quote(result.toString())})"
    webView.post { webView.evaluateJavascript(call, null) }
  }
}
