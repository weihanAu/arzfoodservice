from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics

def draw_text_in_container(c, x, y, container_width, container_height, text, align_bottom=False):
    # 初始字体大小
    font_size = 10
    c.setFont("Helvetica", font_size)

    # 计算文本行
    text_lines = []
    words = text.split(' ')
    current_line = ''

    for word in words:
        test_line = current_line + (word + ' ')  # 加上空格

        # 检查当前行的宽度是否超出容器宽度
        if c.stringWidth(test_line, "Helvetica", font_size) <= container_width:
            current_line = test_line  # 更新当前行
        else:
            # 如果当前行超出宽度，保存当前行并开始新的一行
            if current_line:  # 确保不保存空行
                text_lines.append(current_line)
            current_line = word + ' '  # 重置为当前单词开始新行

    # 添加最后一行
    if current_line:
        text_lines.append(current_line)

    # 动态调整字体大小以适应容器高度和宽度
    while True:
        # 检查是否所有行都适合容器的宽度
        fits_width = all(c.stringWidth(line, "Helvetica", font_size) <= container_width for line in text_lines)

        # 检查容器高度
        total_height = len(text_lines) * (font_size * 1.2)  # 每行大约占用1.2倍的字体大小
        fits_height = total_height <= container_height

        if fits_width and fits_height:
            break  # 如果宽度和高度都适合，则退出循环

        # 如果不适合，则缩小字体
        font_size -= 1
        if font_size < 1:  # 确保字体大小不小于1
            break

        c.setFont("Helvetica", font_size)

        # 重新计算每行
        text_lines.clear()
        current_line = ''
        for word in words:
            test_line = current_line + (word + ' ')
            if c.stringWidth(test_line, "Helvetica", font_size) <= container_width:
                current_line = test_line
            else:
                if current_line:  # 确保不保存空行
                    text_lines.append(current_line)
                current_line = word + ' '

        if current_line:
            text_lines.append(current_line)

    # 在容器内绘制文本，左对齐
    text_object = c.beginText(x, y + container_height)  # 从容器顶部开始绘制
    text_object.setFont("Helvetica", font_size)

    if align_bottom:
        # 如果对齐底部，计算文本的起始Y坐标
        y_start = y + container_height - total_height
        text_object = c.beginText(x, y_start)  # 从容器底部开始绘制
    else:
        # 否则，保持顶部对齐
        y_start = y + container_height

    for line in text_lines:
        text_object.textLine(line.strip())  # 去掉多余的空格

    c.drawText(text_object)

def create_label_pdf(output_path):
    custom_width = 102 * mm  # 自定义宽度
    custom_height = 35 * mm   # 自定义高度
    c = canvas.Canvas(output_path, pagesize=(custom_width, custom_height))

    # Draw the border
    # 将边框调整到适应新的纸张大小
    c.rect(1 * mm, 1 * mm, 100 * mm, 33 * mm)  # (x, y, width, height)

    # Text and positions
    c.setFont("Helvetica-Bold", 16)  # 调整字体大小以适应小纸张
    draw_text_in_container(c, 2 * mm, 22 * mm, 90 * mm, 8 * mm, "Sorrento Cafe Sorrento Cafe Sorrento Cafe Sorrento Cafe Sorrento Cafe Sorrento Cafe  Sorrento Cafe Sorrento Cafe ")

    c.setFont("Helvetica", 6)
    c.drawString(2 * mm, 27 * mm, "ORDER NO.   750124")
    c.drawString(2 * mm, 25 * mm, "DATE          2024-10-24")

    # # Add logo placeholder (optional - you can load an actual image instead)
    # c.drawString(50 * mm, 24 * mm, "ARZ")  # Adjust位置以适应纸张
    # c.setFont("Helvetica", 5)
    # c.drawString(48 * mm, 22 * mm, "FOOD SERVICE")

    # Almond Meal text
    c.setFont("Helvetica-Bold", 12)
    draw_text_in_container(c, 2 * mm, 3* mm, 101 * mm, 8 * mm, "ALMOND **** MEAL **** 1kgALMOND **** MEAL **** 1kgALMOND **** MEAL **** 1kgALMOND **** MEAL **** 1kgALMOND **** MEAL **** 1kgALMOND **** MEAL **** 1kg",True)
    # c.drawString(2 * mm, 2 * mm, 102 * mm,"ALMOND **** MEAL **** 1kg")

    # Black box (Run, Sub, QTY)
    c.setFillColorRGB(0, 0, 0)  # Set color to black
    c.rect(70 * mm, 13 * mm, 30 * mm, 12 * mm, fill=1)  # Black box

    c.setFillColorRGB(1, 1, 1)  # Set color to white for text

    draw_text_in_container(c, 71* mm, 22* mm,  20* mm, 4 * mm, "RUN 9SUB A",True)
    draw_text_in_container(c, 71* mm, 18* mm,  20* mm, 4 * mm, "SUB RUN 2",True)

    c.setFont("Helvetica-Bold", 6)
    draw_text_in_container(c, 71* mm, 14* mm,  20* mm, 4 * mm, "QTY 1 unit",True)

    c.save()

# 调用函数并保存到当前路径
create_label_pdf("label_output.pdf")

