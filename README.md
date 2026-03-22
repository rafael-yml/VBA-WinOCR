# VBA-WinOCR

VBA class for extracting text from images using the native Windows OCR engine (`Windows.Media.Ocr`), via direct WinRT vtable calls. No external tools, no shell execution, no admin rights, no installs.

Originally based on [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR). VBA-WinOCR functions in many similar ways but has since diverged significantly: it removes the `Kernel32.dll` dependency which is known to be blocked in some corporate environments, adds automatic image scaling with high-quality interpolation, and adds `BytesToText` for in-memory pipeline use.

Thanks to [Jaafar Tribak - vtblCall](https://www.mrexcel.com/board/threads/late-bound-windows-media-player-going-out-of-scope.1245903/post-6110097) & [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR)

---

## Features

- Extract text from any image file (PNG, JPEG, BMP, TIFF, GIF)
- Accept raw `Byte()` arrays directly via `BytesToText`. Writes a GUID-named temp file internally, no visible disk activity from the caller's perspective
- Optional line-by-line or word-by-word output with bounding box coordinates
- Automatic image scaling when input exceeds OCR engine limits (preserves accuracy)
- Language detection from system locale; supports any language installed in Windows
- No external dependencies beyond what ships with Windows 10/11
- No `Kernel32.dll` usage, safe in environments where Defender flags low-level API calls

---

## Usage

### Basic full text from image file

```vb
Dim ocr As New WinOCR
MsgBox ocr.ImageToText("C:\images\scan.png")(0)
```

### Full text from in-memory bytes (pipeline use)

```vb
Dim ocr As New WinOCR
Dim pdf As New PdfWRT

Dim oPages As Collection
Set oPages = pdf.RenderPDFToBytes("C:\docs\scan.pdf")

Dim page As Variant
Dim sText As String
For Each page In oPages
    Dim aBytes() As Byte
    aBytes = page
    sText = sText & ocr.BytesToText(aBytes) & vbLf
Next page
```

### With explicit language

```vb
Dim ocr As New WinOCR
MsgBox ocr.ImageToText("C:\images\scan.png", "en-US")(0)
```

### Line-by-line output

```vb
Dim ocr As New WinOCR
Dim lines As Variant
lines = ocr.ImageToText("C:\images\scan.png", "", True)
Dim i As Long
For i = 0 To UBound(lines)
    Debug.Print lines(i)
Next i
```

### Word array with bounding boxes

```vb
Dim ocr As New WinOCR
Dim words As Variant
words = ocr.ImageToText("C:\images\scan.png", "", True, True)
' Each element: Array(text, x, y, width, height)
Dim w As Variant
For Each w In words
    Debug.Print w(0), w(1), w(2), w(3), w(4)
Next w
```

### List supported languages

```vb
Dim ocr As New WinOCR
Dim langs As Collection
Set langs = ocr.GetSupportedLanguages
Dim l As Variant
For Each l In langs
    Debug.Print l(0), l(1)   ' tag, display name
Next l
```

---

## Public API

| Function | Returns | Description |
|---|---|---|
| `ImageToText(PathImage, [Language], [UseLines], [ReturnWordsArray])` | `Variant()` | OCR an image file. `result(0)` is always the full text string. |
| `BytesToText(aImageBytes(), [Language], [UseLines])` | `String` | OCR a PNG supplied as a `Byte()` array. Returns plain text. For pipeline use with `RenderPDFToBytes` (VBA-PdfWRT) or `GetImages` (VBA-WdCOM). |
| `GetSupportedLanguages()` | `Collection` | Returns installed OCR languages. Each item is `Array(tag, displayName)`. |

### ImageToText return value

`ImageToText` always returns a `Variant()` array. `ResultArray(0)` always contains the full concatenated text string regardless of other options. When `ReturnWordsArray = True`, subsequent elements contain word info arrays as `Array(text, x, y, width, height)`.

---

---

## Status codes

`LastStatus` is set after every `ImageToText` and `BytesToText` call.

| Constant | Value | Meaning |
|---|---|---|
| `WINOCR_OK` | 0 | Text recognised successfully |
| `WINOCR_NO_ENGINE` | 1 | Language not installed or engine unavailable |
| `WINOCR_TOO_LARGE` | 2 | Image exceeds `OcrEngine.MaxImageDimension` |
| `WINOCR_EMPTY` | 3 | OCR completed but returned no text |
| `WINOCR_FAIL` | 4 | File missing, decode error, or init failure |

```vb
Dim ocr As New WinOCR
ocr.ImageToText "C:\scan.png"
Select Case ocr.LastStatus
    Case WINOCR_OK:        ' text in result(0)
    Case WINOCR_TOO_LARGE: ' image too large, try lower render width
    Case WINOCR_EMPTY:     ' page was blank
End Select
```

## Resolution and OCR accuracy

`Windows.Media.Ocr` enforces two hard limits on its input bitmap:

| Limit | Value | Effect if exceeded |
|---|---|---|
| Per-dimension cap (`MaxImageDimension`) | 5000 px on Win10 1903+, 2048 px on older builds | Engine refuses the call |
| Total pixel budget | ~5 megapixels (internal) | Engine silently downsamples with nearest-neighbour interpolation before recognition |

The nearest-neighbour downsampling the engine performs internally is low quality. It produces aliasing on diagonal strokes and blurs fine text, which is the documented source of accuracy losses of up to ~40% on high-resolution inputs.

**WinOCR handles this automatically.** When the input image exceeds either limit, it downsamples the bitmap itself using **Fant interpolation** (bicubic area-average, the highest quality mode available in `Windows.Graphics.Imaging`) before passing it to `RecognizeAsync`. The engine then receives a clean, properly-sampled image.

You do not need to pre-scale images before calling `ImageToText` or `BytesToText`. Any resolution input is handled correctly.

**Practical implication for the PdfWRT pipeline:** When using WinOCR together with VBA-PdfWRT, you can render at any width. The recommended default of `2480px` (300 dpi for A4) produces an 8.7MP image that WinOCR will automatically scale to approximately 1876x2655 (~4.98MP) before recognition. Rendering at `4960px` (600 dpi) or higher is fine, WinOCR will scale it down correctly regardless.

---

## DLL dependencies

All DLLs are present on every modern Windows installation.

| DLL | Usage |
|---|---|
| `Combase.dll` | `RoGetActivationFactory`, `RoActivateInstance`, `WindowsCreateString`, `WindowsDeleteString`, `WindowsGetStringRawBuffer` |
| `Shcore.dll` | `CreateRandomAccessStreamOnFile` |
| `oleAut32.dll` | `DispCallFunc` (vtable call dispatcher) |
| `ole32.dll` | `CLSIDFromString`, `CoCreateGuid` |
| `msvcrt.dll` | `memcpy` (wide-char copy), `_sleep` (non-blocking async poll) |

---

## Tested environments

| OS | Edition | Version | Build |
|---|---|---|---|
| Windows 10 | Enterprise | 21H2 | 19044 |
| Windows 11 | Enterprise | 23H2 | 22631 |

---

## Changes from DanysysTeam/VBA-UWPOCR

- Replaced `Kernel32.dll` (`RtlMoveMemory`) with `msvcrt.dll` (`memcpy`) to avoid Defender triggers in corporate environments
- Added `BytesToText` for in-memory pipeline use (accepts `Byte()` array, returns plain `String`)
- Added timeout/retry counter to `WaitForAsyncInterface`
- Fixed `RECT` type fields from `Single` to `Long`
- Fixed `ResultArray` indexing bug in word array mode
- Automatic image scaling before OCR using Fant interpolation, eliminates accuracy loss from oversized inputs
- Added `msvcrt.dll _sleep` to async poll loops, eliminates 100% CPU spin during recognition
- Fixed `hString` leak in `RoGetActivationIFactory`, `hStringLanguage` leak in `CreateOcrEngine`
- Fixed COM object leaks (`pIAsyncInfo`, `pILanguage`, `pIRandomAccessStream`, `pIBitmapTransform`)
- Removed unreachable code after `Err.Raise` in `WaitForAsyncInterface`
- `LastStatus` property and `WINOCR_*` status codes added
- `WINOCR_TOO_LARGE` reported when image exceeds `MaxImageDimension`
- `pFIVLanguages` null guard added in `GetSupportedLanguages`

---

## License

MIT License. See [LICENSE](LICENSE) for details.

Copyright © 2024, [Danysys](https://www.danysys.com)

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)
