// TextEncoder and TextDecoder exist in browsers, workers and Node; the package's ES-only lib doesn't declare them.
declare class TextDecoder {
  constructor(label?: string, options?: { fatal?: boolean })
  decode(input?: Uint8Array): string
}
declare class TextEncoder {
  encode(input?: string): Uint8Array
}
