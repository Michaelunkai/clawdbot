"""Create an icon file for the exe"""
from PIL import Image, ImageDraw, ImageFont

def create_icon():
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    images = []
    
    for size in sizes:
        img = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # Draw green circle
        margin = size[0] // 8
        draw.ellipse([margin, margin, size[0]-margin, size[1]-margin], fill='#4CAF50')
        
        # Draw "C" in white
        font_size = size[0] // 2
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
        
        # Center the text
        bbox = draw.textbbox((0, 0), "C", font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        x = (size[0] - text_width) // 2
        y = (size[1] - text_height) // 2 - bbox[1]
        draw.text((x, y), "C", fill='white', font=font)
        
        images.append(img)
    
    # Save as ICO
    images[0].save('clawdbot.ico', format='ICO', sizes=[(s[0], s[1]) for s in sizes], append_images=images[1:])
    print("Created clawdbot.ico")

if __name__ == '__main__':
    create_icon()
