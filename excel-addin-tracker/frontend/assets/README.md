# Icons

The placeholder icon files need to be replaced with actual PNG images:

- icon-16.png (16x16 pixels)
- icon-32.png (32x32 pixels)
- icon-64.png (64x64 pixels)
- icon-80.png (80x80 pixels)

You can create these using any image editing tool. Recommended:
- Use your company logo or a spreadsheet/chart icon
- Ensure transparent background
- Save as PNG format

For quick testing, you can use https://www.flaticon.com/ or create simple colored squares.

## Quick Icon Generation

You can use ImageMagick to create simple placeholder icons:

```bash
# Install ImageMagick
brew install imagemagick  # Mac
apt-get install imagemagick  # Linux

# Create simple blue squares
convert -size 16x16 xc:#0078D4 icon-16.png
convert -size 32x32 xc:#0078D4 icon-32.png
convert -size 64x64 xc:#0078D4 icon-64.png
convert -size 80x80 xc:#0078D4 icon-80.png
```
