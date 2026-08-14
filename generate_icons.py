from PIL import Image, ImageDraw, ImageFilter

def draw_color_icon():
    size = 192
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Background (rounded rect)
    bg_color = (20, 20, 45, 255)
    draw.rounded_rectangle([10, 10, size-10, size-10], radius=30, fill=bg_color)
    
    # Gradient-like envelope/airplane shape
    points = [
        (40, size//2),
        (size-40, 40),
        (size-40, size-40)
    ]
    draw.polygon(points, fill=(124, 92, 252, 255))
    
    points2 = [
        (40, size//2),
        (size-40, size//2),
        (size-40, size-40)
    ]
    draw.polygon(points2, fill=(56, 182, 255, 255))
    
    img.save('color.png')

def draw_outline_icon():
    size = 32
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Simple outline of an envelope/airplane
    points = [
        (4, size//2),
        (size-4, 4),
        (size-4, size-4),
        (4, size//2)
    ]
    draw.line(points, fill=(255, 255, 255, 255), width=2, joint="curve")
    draw.line([(4, size//2), (size-4, size//2)], fill=(255, 255, 255, 255), width=2)
    
    img.save('outline.png')

if __name__ == '__main__':
    draw_color_icon()
    draw_outline_icon()
    print("Icons generated successfully.")
