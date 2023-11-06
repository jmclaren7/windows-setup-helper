
This archive contains two precompiled DLL's for 7-Zip v17.00 or higher.

## Installation

1. download the codec archiv from https://mcmilk.de/projects/7-Zip-zstd/
2. create a new directory named "Codecs"
3. put in there the CODEC-x32.dll or the CODEC-x64.dll, depending on your 7-Zip installation
   - normally, the x32 should go to: "C:\Program Files (x86)\7-Zip\Codecs"
   - and the x64 version should go in here: "C:\Program Files\7-Zip\Codecs"
4. then, you may check if the dll is correctly installed via this command: `7z.exe i`

The output should look like this:
```

7-Zip 17.00 beta (x64) : Copyright (c) 1999-2017 Igor Pavlov : 2017-04-29


Libs:
 0  c:\Program Files\7-Zip\7z.dll
 1  c:\Program Files\7-Zip\Codecs\brotli-x64.dll
 2  c:\Program Files\7-Zip\Codecs\lizard-x64.dll
 3  c:\Program Files\7-Zip\Codecs\lz4-x64.dll
 4  c:\Program Files\7-Zip\Codecs\lz5-x64.dll
 5  c:\Program Files\7-Zip\Codecs\zstd-x64.dll

Formats:
 0               APM      apm           E R
 ...

Codecs:
 0 4ED  303011B BCJ2
 0  ED  3030103 BCJ
 0  ED  3030205 PPC
 0  ED  3030401 IA64
 0  ED  3030501 ARM
 0  ED  3030701 ARMT
 0  ED  3030805 SPARC
 0  ED    20302 Swap2
 0  ED    20304 Swap4
 0  ED    40202 BZip2
 0  ED        0 Copy
 0  ED    40109 Deflate64
 0  ED    40108 Deflate
 0  ED        3 Delta
 0  ED       21 LZMA2
 0  ED    30101 LZMA
 0  ED    30401 PPMD
 0   D    40301 Rar1
 0   D    40302 Rar2
 0   D    40303 Rar3
 0   D    40305 Rar5
 0  ED  6F10701 7zAES
 0  ED  6F00181 AES256CBC
 1  ED  4F71102 BROTLI
 2  ED  4F71106 LIZARD
 3  ED  4F71104 LZ4
 4  ED  4F71105 LZ5
 5  ED  4F71101 ZSTD

Hashers:
 0    4        1 CRC32
 0   20      201 SHA1
 0   32        A SHA256
 0    8        4 CRC64
 0   32      202 BLAKE2sp
```

## Usage

- when compressing binaries (*.exe, *.dll), you have to explicitly disable
  the bcj2 filter via `-m0=bcj`
- so the usage should look like this:

```
7z a archiv.7z -m0=bcj -m1=zstd -mx1   ...Fast mode, with BCJ preprocessor on executables
7z a archiv.7z -m0=bcj -m1=zstd -mx..  ...
7z a archiv.7z -m0=bcj -m1=zstd -mx21  ...2nd Slowest Mode, with BCJ preprocessor on executables
7z a archiv.7z -m0=bcj -m1=zstd -mx22  ...Ultra Mode, with BCJ preprocessor on executables
```

## License and redistribution

- the same as the original 7-Zip, which means GNU LGPL

/TR 2017-05-25
