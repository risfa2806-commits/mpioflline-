import math
import struct
import zlib

def create_png(width, height, draw_fn, filename):
    # RGBA buffer
    pixels = bytearray(width * height * 4)

    def set_pixel(x, y, r, g, b, a):
        if 0 <= x < width and 0 <= y < height:
            idx = (y * width + x) * 4
            # Alpha blend
            existing_a = pixels[idx + 3] / 255.0
            new_a = a / 255.0
            out_a = new_a + existing_a * (1.0 - new_a)
            if out_a > 0:
                out_r = (r * new_a + pixels[idx] * existing_a * (1.0 - new_a)) / out_a
                out_g = (g * new_a + pixels[idx + 1] * existing_a * (1.0 - new_a)) / out_a
                out_b = (b * new_a + pixels[idx + 2] * existing_a * (1.0 - new_a)) / out_a
                pixels[idx] = int(out_r)
                pixels[idx + 1] = int(out_g)
                pixels[idx + 2] = int(out_b)
                pixels[idx + 3] = int(out_a * 255)

    def draw_thick_line(x0, y0, x1, y1, thickness=2.5, r=15, g=23, b=42):
        dist = math.hypot(x1 - x0, y1 - y0)
        steps = max(int(dist * 3), 1)
        radius = thickness / 2.0
        r_ceil = int(math.ceil(radius + 1))
        for s in range(steps + 1):
            t = s / float(steps)
            cx = x0 + t * (x1 - x0)
            cy = y0 + t * (y1 - y0)
            for dy in range(-r_ceil, r_ceil + 1):
                for dx in range(-r_ceil, r_ceil + 1):
                    d = math.hypot(dx, dy)
                    if d <= radius:
                        alpha = 255
                    elif d <= radius + 1.0:
                        alpha = int(255 * (1.0 - (d - radius)))
                    else:
                        continue
                    set_pixel(int(cx + dx), int(cy + dy), r, g, b, alpha)

    def draw_filled_circle(cx, cy, radius, r=15, g=23, b=42):
        r_ceil = int(math.ceil(radius + 1))
        for dy in range(-r_ceil, r_ceil + 1):
            for dx in range(-r_ceil, r_ceil + 1):
                d = math.hypot(dx, dy)
                if d <= radius:
                    alpha = 255
                elif d <= radius + 1.0:
                    alpha = int(255 * (1.0 - (d - radius)))
                else:
                    continue
                set_pixel(int(cx + dx), int(cy + dy), r, g, b, alpha)

    draw_fn(draw_thick_line, draw_filled_circle)

    # Encode PNG
    raw_data = bytearray()
    for y in range(height):
        raw_data.append(0) # Filter byte: None
        row_start = y * width * 4
        raw_data.extend(pixels[row_start:row_start + width * 4])

    compressed = zlib.compress(bytes(raw_data), 9)

    png = bytearray(b'\x89PNG\r\n\x1a\n')

    # IHDR
    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data)
    png.extend(struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc))

    # IDAT
    idat_crc = zlib.crc32(b'IDAT' + compressed)
    png.extend(struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', idat_crc))

    # IEND
    iend_crc = zlib.crc32(b'IEND')
    png.extend(struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc))

    with open(filename, 'wb') as f:
        f.write(png)
    print(f"Wrote {filename}, size {len(png)} bytes")

def bezier_points(p0, p1, p2, p3, num_points=40):
    pts = []
    for i in range(num_points + 1):
        t = i / float(num_points)
        x = (1-t)**3 * p0[0] + 3*(1-t)**2 * t * p1[0] + 3*(1-t) * t**2 * p2[0] + t**3 * p3[0]
        y = (1-t)**3 * p0[1] + 3*(1-t)**2 * t * p1[1] + 3*(1-t) * t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts

def render_signature(draw_line, draw_circle):
    # 1. Main vertical stem
    draw_line(68, 25, 78, 140, thickness=2.8)
    # small dot / accent on vertical stem
    draw_circle(67, 72, radius=1.6)

    # 2. Large sweeping oval loop on the left
    pts1 = bezier_points((75, 115), (45, 122), (18, 95), (15, 60), 30)
    pts2 = bezier_points((15, 60), (12, 28), (40, 16), (72, 22), 30)
    pts3 = bezier_points((72, 22), (90, 26), (88, 70), (68, 115), 30)
    for pts in [pts1, pts2, pts3]:
        for i in range(len(pts) - 1):
            draw_line(pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], thickness=2.4)

    # 3. Small superscript '2' at top right of vertical stem
    pts_two = bezier_points((102, 38), (102, 32), (112, 30), (115, 36), 15)
    for i in range(len(pts_two) - 1):
        draw_line(pts_two[i][0], pts_two[i][1], pts_two[i+1][0], pts_two[i+1][1], thickness=1.8)
    draw_line(115, 36, 103, 46, thickness=1.8)
    draw_line(103, 46, 116, 46, thickness=1.8)

    # 4. Belly of D / main letter
    d_pts1 = bezier_points((70, 48), (105, 42), (122, 60), (115, 82), 25)
    d_pts2 = bezier_points((115, 82), (105, 105), (75, 118), (62, 112), 25)
    for pts in [d_pts1, d_pts2]:
        for i in range(len(pts) - 1):
            draw_line(pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], thickness=2.6)

    # 5. Sharp vertical zig-zags (teeth/scribbles)
    # Peak 1
    draw_line(125, 78, 131, 38, thickness=2.2)
    draw_line(131, 38, 136, 80, thickness=2.2)
    # Peak 2
    draw_line(136, 80, 142, 28, thickness=2.4)
    draw_line(142, 28, 146, 84, thickness=2.4)
    # Peak 3
    draw_line(146, 84, 150, 32, thickness=2.4)
    draw_line(150, 32, 154, 82, thickness=2.2)
    # Peak 4
    draw_line(154, 82, 158, 42, thickness=2.2)
    draw_line(158, 42, 161, 80, thickness=2.0)
    # Peak 5
    draw_line(161, 80, 164, 48, thickness=2.0)
    draw_line(164, 48, 167, 78, thickness=2.0)
    # Peak 6 (trailing)
    draw_line(167, 78, 170, 56, thickness=1.8)
    draw_line(170, 56, 172, 75, thickness=1.8)

    # 6. Horizontal cross-strike line through the zigzags
    draw_line(68, 70, 225, 72, thickness=2.5)

    # 7. Bottom flourish curve starting below D
    flourish1 = bezier_points((98, 120), (106, 148), (124, 158), (142, 148), 25)
    flourish2 = bezier_points((142, 148), (155, 140), (168, 136), (175, 138), 20)
    for pts in [flourish1, flourish2]:
        for i in range(len(pts) - 1):
            draw_line(pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], thickness=2.6)

    # 8. Accent dot at end of flourish
    draw_circle(192, 142, radius=2.5)

if __name__ == '__main__':
    create_png(240, 170, render_signature, 'public/ttd_dosen.png')
