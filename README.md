# VBA-WinOCR

VBA-WinOCR is a simple library to use Universal Windows Platform Optical character recognition API.
Based on 

## Features

* Get Text From Image File.
* Easy to use.

## Usage

##### Basic use:

```VB
    Dim ocr As New WinOCR

    MsgBox ocr.ImageToText(ThisWorkbook.Path & "\Images\Image1.png")(0)

```

<!-- ##### More examples [here.](/Examples) -->

## Changes from original
Forked from [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR).

- Replaced `Kernel32.dll` (`RtlMoveMemory`) with `msvcrt.dll` (`memcpy`) 
  to avoid triggering Windows Defender on certain corporate environments
- Added timeout/retry counter to `WaitForAsyncInterface`
- Fixed `RECT` type fields from `Single` to `Long`
- Fixed `ResultArray` indexing bug in word array mode

## Release History

See [CHANGELOG.md](CHANGELOG.md)

<!-- ## Acknowledgments & Credits -->

## License

Usage is provided under the [MIT](https://choosealicense.com/licenses/mit/) License.

Copyright © 2024, [Danysys.](https://www.danysys.com)
Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)
