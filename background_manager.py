import os
import requests
import io
import tkinter as tk
from PIL import Image, ImageTk, ImageFilter, ImageEnhance

class BackgroundManager:
    """Manages high-resolution background images for the JARVIS interface"""
    
    def __init__(self):
        self.background_image = None
        self.background_photo = None
        self.background_id = None
        
    def load_background_image(self, canvas, width, height):
        """Load a colorful 4K background image (local or remote)."""
        try:
            # Clear previous background if any
            if self.background_id:
                try:
                    canvas.delete(self.background_id)
                except Exception:
                    pass
                self.background_id = None

            # Prefer local asset if available
            local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "background.jpg")

            img = None
            if os.path.exists(local_path):
                img = Image.open(local_path)
            else:
                # Colorful 4K fallbacks
                urls = [
                    "https://images.unsplash.com/photo-1500530855697-b586d89ba3ee?w=3840",  # colorful gradient sky
                    "https://images.unsplash.com/photo-1520975916090-3105956dac38?w=3840",  # neon colors
                    "https://images.unsplash.com/photo-1517816743773-6e0fd518b4a6?w=3840",  # abstract colors
                ]
                for url in urls:
                    try:
                        resp = requests.get(url, timeout=10)
                        if resp.status_code == 200:
                            img = Image.open(io.BytesIO(resp.content))
                            # Cache locally for next run
                            os.makedirs(os.path.dirname(local_path), exist_ok=True)
                            try:
                                img.save(local_path)
                            except Exception:
                                pass
                            break
                    except Exception:
                        continue

            if img is None:
                # Fallback: gradient background
                return self._create_gradient_background(canvas, width, height)

            # Process and draw
            img = self._process_image(img, width, height)
            self.background_photo = ImageTk.PhotoImage(img)
            self.background_id = canvas.create_image(width // 2, height // 2, image=self.background_photo)
            canvas.tag_lower(self.background_id)
            return True
        except Exception as e:
            print(f"Error loading background image: {e}")
            return self._create_gradient_background(canvas, width, height)
    
    def _process_image(self, img, width, height):
        """Process the image to fit the interface with 4K optimization"""
        # For 4K displays, maintain high quality
        if width >= 3840 or height >= 2160:
            # Use high-quality resizing for 4K
            img = img.resize((width, height), Image.LANCZOS)
            # Apply minimal blur for 4K displays
            img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
        else:
            # Standard processing for lower resolutions
            img = img.resize((width, height), Image.LANCZOS)
            img = img.filter(ImageFilter.GaussianBlur(radius=1))
        
        # Apply advanced image enhancement for better UI visibility
        # Brightness adjustment
        brightness_enhancer = ImageEnhance.Brightness(img)
        img = brightness_enhancer.enhance(0.8)  # Slightly brighter for 4K
        
        # Contrast enhancement
        contrast_enhancer = ImageEnhance.Contrast(img)
        img = contrast_enhancer.enhance(1.2)  # Increase contrast
        
        # Color enhancement
        color_enhancer = ImageEnhance.Color(img)
        img = color_enhancer.enhance(1.1)  # Slightly enhance colors
        
        # Sharpness enhancement for 4K
        if width >= 3840 or height >= 2160:
            sharpness_enhancer = ImageEnhance.Sharpness(img)
            img = sharpness_enhancer.enhance(1.3)  # Sharper for 4K displays
        
        return img
    
    def _create_gradient_background(self, canvas, width, height):
        """Create a fallback gradient background"""
        # Create gradient from dark blue to black
        for i in range(height):
            # Calculate color based on position
            r = int(0 * (1 - i/height))
            g = int(10 * (1 - i/height))
            b = int(30 * (1 - i/height))
            color = f'#{r:02x}{g:02x}{b:02x}'
            
            # Draw a line with the calculated color
            line_id = canvas.create_line(0, i, width, i, fill=color, width=1)
            canvas.tag_lower(line_id)  # Move to bottom
        
        return True
    
    def update_background(self, canvas, width, height):
        """Update the background if window size changes"""
        if self.background_id:
            canvas.delete(self.background_id)
        
        return self.load_background_image(canvas, width, height)
    
    def optimize_for_4k(self, canvas, width, height):
        """Optimize background specifically for 4K displays"""
        if width >= 3840 or height >= 2160:
            # Reload with 4K optimizations
            if self.background_id:
                canvas.delete(self.background_id)
            self.load_background_image(canvas, width, height)
            return True
        return False