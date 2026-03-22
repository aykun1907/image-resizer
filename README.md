# Image Resizer

A lightweight Windows tool for resizing and converting images. Integrates into the Windows **Send To** right-click menu for quick access.

## Features

- **Resize** images by percentage or custom dimensions
- **Convert** between JPEG, PNG, WebP, and AVIF without resizing
- **Batch processing** — select multiple images at once
- **Windows Send To integration** — right-click images to resize/convert
- **Format-specific options:**
  - JPEG: chroma subsampling (4:2:0 / 4:4:4), Huffman optimization
  - WebP: compression effort (0-6)
  - AVIF: speed/quality tradeoff
- **Smart file handling** — auto-appends `(1)`, `(2)` to avoid overwriting files
- **EXIF orientation** — automatically corrects rotated images before processing
- **Portable** — single `.exe`, no installation required

## Download

Grab the latest `ImageResizer.exe` from the [Releases](../../releases) page.

## Usage

### Quick start
1. Double-click `ImageResizer.exe`
2. Click **Enable Integration** to add it to your right-click Send To menu
3. Select images in Explorer, right-click > **Send To** > **Resize Images**

### Standalone
Double-click the exe and use **Browse** to pick images.

## Building from source

Requires Python 3.9+.

```bash
pip install -r requirements.txt
python build_script.py
```

The exe will be in `dist/ImageResizer.exe`.

## Contributing

1. Fork the repo
2. Create a branch (`git checkout -b my-feature`)
3. Commit your changes
4. Push and open a Pull Request

## License

[MIT](LICENSE)
