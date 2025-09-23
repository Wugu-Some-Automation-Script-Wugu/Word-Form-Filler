#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Create application icon
"""

import sys
import os

# Set UTF-8 encoding for Windows
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
    sys.stderr = codecs.getwriter("utf-8")(sys.stderr.detach())

try:
    from PIL import Image, ImageDraw, ImageFont
    
    # Create 256x256 icon
    size = (256, 256)
    img = Image.new('RGBA', size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw background circle
    margin = 20
    draw.ellipse([margin, margin, size[0]-margin, size[1]-margin], 
                 fill=(52, 152, 219, 255), outline=(41, 128, 185, 255), width=4)
    
    # Draw text "W" (Word)
    try:
        # Try to use system font
        font = ImageFont.truetype("arial.ttf", 120)
    except:
        # If font not found, use default font
        font = ImageFont.load_default()
    
    text = "W"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    x = (size[0] - text_width) // 2
    y = (size[1] - text_height) // 2 - 10
    
    draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)
    
    # Save as ICO format
    img.save('icon.ico', format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("Icon file icon.ico created successfully")
    
except ImportError:
    print("PIL not installed, skipping icon creation")
    print("To create icon, install Pillow: pip install Pillow")
except Exception as e:
    print(f"Error creating icon: {e}")
    print("Will use default icon")
