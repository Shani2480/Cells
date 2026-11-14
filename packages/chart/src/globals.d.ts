// TextEncoder exists in browsers, workers and Node; the package's ES-only lib doesn't declare it.
declare class TextEncoder {
  encode(input?: string): Uint8Array
}
