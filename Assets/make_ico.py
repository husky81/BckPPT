"""icon.png의 둥근사각형 바깥 흰색을 투명으로 → ICO 생성"""
from PIL import Image
import numpy as np
from collections import deque

src = r'C:\Users\bckim\BckSoft\cellbox2\apps\tray\BckPPT\Assets\icon.png'
img = Image.open(src).convert('RGBA')
data = np.array(img)

r, g, b = data[:,:,0], data[:,:,1], data[:,:,2]
is_white = (r > 240) & (g > 240) & (b > 240)

h, w = is_white.shape
mask = np.zeros((h, w), dtype=bool)

# BFS flood fill from 4 corners through white pixels
queue = deque()
for sy, sx in [(0,0), (0,w-1), (h-1,0), (h-1,w-1)]:
    if is_white[sy, sx]:
        queue.append((sy, sx))
        mask[sy, sx] = True

while queue:
    y, x = queue.popleft()
    for dy, dx in [(-1,0),(1,0),(0,-1),(0,1)]:
        ny, nx = y+dy, x+dx
        if 0 <= ny < h and 0 <= nx < w and not mask[ny, nx] and is_white[ny, nx]:
            mask[ny, nx] = True
            queue.append((ny, nx))

# Make outer white pixels transparent
data[mask, 3] = 0

result = Image.fromarray(data)

# Save transparent PNG
png_path = src.replace('icon.png', 'icon_transparent.png')
result.save(png_path)
print(f'Saved: {png_path}')

# Generate ICO with multiple sizes
ico_path = src.replace('icon.png', 'icon.ico')
sizes = [16, 24, 32, 48, 64, 128, 256]
ico_images = []
for s in sizes:
    resized = result.resize((s, s), Image.LANCZOS)
    ico_images.append(resized)

ico_images[0].save(ico_path, format='ICO', sizes=[(s,s) for s in sizes], append_images=ico_images[1:])
print(f'Saved: {ico_path} (sizes: {sizes})')
