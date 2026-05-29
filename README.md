# Speed Checker — Android

## Files
- `main.py` — the Kivy app
- `buildozer.spec` — packaging config
- `Dockerfile` — for building with Docker

## How to build the APK

### Using Docker Desktop (recommended)

1. Open a terminal in this folder
2. Run:
   ```
   docker build -t speedchecker .
   docker run --rm -v "%cd%/bin:/app/bin" speedchecker
   ```
   (On Mac/Linux replace `%cd%` with `$(pwd)`)

3. The APK appears in the `bin/` folder as something like:
   `speedchecker-1.0-arm64-v8a-debug.apk`

First build takes ~20 minutes (downloads Android SDK/NDK). Subsequent builds are much faster.

### Install on your phone

1. Email the APK to yourself, or copy via USB
2. On your Android device: Settings → Apps → Special app access → Install unknown apps → allow your browser/file manager
3. Open the APK and install

## OneDrive setup

Put your files in OneDrive on your PC and let them sync:
```
OneDrive/
  seediq/
    metadata_edit.xlsx
    sentences/
      a_1.1.wav
      abuh_1.1.wav
      ...
    words/
      a.wav
      abuh.wav
      ...
```

On Android, OneDrive syncs to:
`/storage/emulated/0/OneDrive/`

When the app opens, tap **📂 Excel** and navigate to your xlsx file,
then tap **🎵 Audio folder** and navigate to the sentences or words folder.
The app remembers your last location between taps.

## Notes

- Speed control uses Kivy's `pitch` property — works on Android's default audio backend
- If pitch doesn't work for speed (some devices), the slider still changes the value for future use; we can add a Python-level resample fallback
- Save writes back to the xlsx file directly — make sure OneDrive has synced before editing
