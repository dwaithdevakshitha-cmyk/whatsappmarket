import codecs

with codecs.open('whatsappmarket.py', 'r', 'utf-8') as f:
    lines = f.readlines()

out_lines = []
in_media = False
in_msg = False

for line in lines:
    if line.startswith('MEDIA_PATHS = ['):
        in_media = True
        out_lines.append('MEDIA_PATHS = [\n')
        out_lines.append('    r"C:\\Users\\HP\\Documents\\marketing excels images\\image.jpeg"\n')
        continue
    if in_media:
        if ']' in line:
            in_media = False
            out_lines.append(']\n')
        continue

    if line.startswith('MESSAGE_TEXT = """'):
        in_msg = True
        out_lines.append('MESSAGE_TEXT = """Villa for sale Finding a budget-friendly villa within a 40-minute drive\n')
        out_lines.append("from Hyderabad's Financial District and HITEC City can\n")
        out_lines.append("be challenging, as these areas are prime real estate\n")
        out_lines.append("zones with higher property prices. However, exploring\n")
        out_lines.append("nearby localities may offer more affordable options.\n")
        out_lines.append('Contact Villa price start with 99,00,000/"""\n')
        continue
    if in_msg:
        if '"""' in line:
            in_msg = False
        continue

    out_lines.append(line)

with codecs.open('whatsappmarket.py', 'w', 'utf-8') as f:
    f.writelines(out_lines)
print("Updated successfully")
