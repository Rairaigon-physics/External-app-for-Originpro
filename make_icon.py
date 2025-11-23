from PIL import Image

def create_high_quality_ico(source_png, output_ico):
    img = Image.open(source_png)

    # The standard sizes Windows looks for
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

    # Save as ICO including all sizes
    img.save(output_ico, format='ICO', sizes=icon_sizes)
    print(f"Success! High-res icon saved as {output_ico}")

if __name__ == '__main__':
    # Make sure 'logo_source.png' exists in this folder!
    create_high_quality_ico("logo_source.png", "logo.ico")