import pandas as pd
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont

def excel_to_image(excel_file, output_image='assignments.png'):
    # --- Load Excel ---
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()

    # Detect the "Due" column automatically
    due_col = [c for c in df.columns if 'due' in c.lower()][0]

    # --- Robust date parser ---
    def parse_date(val):
        if pd.isna(val):
            return None
        s = str(val).strip().upper()
        if s in ('TBD', 'TO BE DETERMINED', 'TBA', 'NONE', ''):
            return None
        if isinstance(val, (datetime, pd.Timestamp)):
            return val
        if isinstance(val, (int, float)):
            return datetime(1899, 12, 30) + pd.to_timedelta(int(val), unit='D')
        try:
            return datetime.strptime(str(val), '%d/%m/%Y')
        except:
            try:
                return pd.to_datetime(val, dayfirst=True)
            except:
                return None

    df['DueDate'] = df[due_col].apply(parse_date)
    df['SortKey'] = df['DueDate'].apply(lambda d: d if d is not None else datetime.max)
    df_sorted = df.sort_values(by='SortKey')

    # --- Helper for ordinal suffix ---
    def ordinal(n):
        if 10 <= n % 100 <= 20:
            return 'th'
        else:
            return {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')

    # --- Prepare text lines ---
    raw_lines = ["Assignments:"]
    for idx, row in enumerate(df_sorted.itertuples(index=False), start=1):
        dt = row[-1]  # DueDate is last column
        if not pd.isna(dt):
            day = dt.day
            suffix = ordinal(day)
            due_text = f"{day}{suffix} {dt.strftime('%B %Y')}"
        else:
            due_text = 'TBD'

        unit_nr = str(row[0])
        unit_name = str(row[1])
        assignment_nr = str(row[2])
        raw_lines.append(f"{idx}. Due {due_text} --- Unit {unit_nr} {unit_name} Assignment {assignment_nr}")

    # --- Image settings ---
    width = 816
    padding = 20
    font_size = 26
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()

    # --- Wrap lines by pixel width ---
    def wrap_by_pixel(text, font, max_width):
        words = text.split()
        lines = []
        current_line = ""
        for w in words:
            test_line = current_line + (" " if current_line else "") + w
            if font.getlength(test_line) <= max_width:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = w
        if current_line:
            lines.append(current_line)
        return lines

    max_text_width = width - padding * 2
    all_lines = []
    for line in raw_lines:
        all_lines.extend(wrap_by_pixel(line, font, max_text_width))

    # --- Calculate dynamic height ---
    line_height = font.getbbox("Ay")[3] - font.getbbox("Ay")[1] + 6
    total_height = padding * 2 + line_height * len(all_lines)

    # --- Draw image ---
    img = Image.new('RGB', (width, total_height), color='black')
    draw = ImageDraw.Draw(img)

    y = padding
    for line in all_lines:
        draw.text((padding, y), line, font=font, fill='white')
        y += line_height

    img.save(output_image)
    print(f"Saved image to {output_image} (width={width}, height={total_height})")

# --- Example usage ---
excel_to_image(r"C:\Users\iusti\OneDrive\Desktop\assignment_bot\assignments.xlsx")
