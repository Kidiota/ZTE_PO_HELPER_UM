import pdfplumber
import PyPDF2
import os
import shutil
import time
import xlsxwriter

# =========================
# 1. 读取 PDF 中的表格原始数据
# =========================
def get_raw_info():
    """
    从 input 文件夹中的所有 PDF 中提取表格行，返回结构：
    [
        [ header_rows_of_pdf1, item_rows_of_pdf1 ],
        [ header_rows_of_pdf2, item_rows_of_pdf2 ],
        ...
    ]
    这里实际上是拍扁成一维：raw[0], raw[1] 对应第1个PDF；raw[2], raw[3] 对应第2个PDF...
    """
    i = 0
    rawInfo = []
    print("开始提取数据")
    while i < len(filesName):
        fixedFileFullName = os.path.join("input", filesName[i])

        # ------- 第一次：horizontal_strategy='lines' -------
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "lines",
            }
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        # 一堆空行过滤逻辑，保持你原来的写法
                        if row != ['', '', '', '', '', '', '', '', '', '', '', '']:
                            if row != ['', '', '', '', '', '', '', '', '', '', '']:
                                if row != ['', '', '', '', '', '', '', '', '', '']:
                                    if row != ['', '', '', '', '', '', '', '', '']:
                                        if row != ['', '', '', '', '', '', '', '']:
                                            if row != ['', '', '', '', '', '', '']:
                                                if row != ['', '', '', '', '', '']:
                                                    if row != ['', '', '', '', '']:
                                                        if row != ['', '', '', '']:
                                                            if row != ['', '', '']:
                                                                if row != ['', '']:
                                                                    if row != ['']:
                                                                        infoInLine = [row]
                                                                        oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]

        # ------- 第二次：horizontal_strategy='text' -------
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
            }
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        if row != ['', '', '', '', '', '', '', '', '', '', '', '']:
                            if row != ['', '', '', '', '', '', '', '', '', '', '']:
                                if row != ['', '', '', '', '', '', '', '', '', '']:
                                    if row != ['', '', '', '', '', '', '', '', '']:
                                        if row != ['', '', '', '', '', '', '', '']:
                                            if row != ['', '', '', '', '', '', '']:
                                                if row != ['', '', '', '', '', '']:
                                                    if row != ['', '', '', '', '']:
                                                        if row != ['', '', '', '']:
                                                            if row != ['', '', '']:
                                                                if row != ['', '']:
                                                                    if row != ['']:
                                                                        infoInLine = [row]
                                                                        oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]

        i = i + 1
        print('\r' + '已提取' + str(i) + '/' + str(len(filesName)), end='', flush=True)
    print('\n提取完毕')
    return rawInfo


# =========================
# 2. 解析 raw，拆成 PO + 子项目
# =========================
def get_info(raw):
    """
    raw: get_raw_info() 的返回值
    返回 allData 结构：
    [
        [PONUM, DATE, OurRef, [
            [itemNum, material, description, quantity, UOM],
            ...
        ]],
        ...
    ]
    """
    i = len(raw)
    j = 0
    allData = []

    while j < i:
        # ---------- 解析 PO 头部 ----------
        if len(raw[j][0]) == 3:
            lineOfInfo = raw[j][0][2].split("\n")
            PONUM = lineOfInfo[0]
            DATE = lineOfInfo[1]
            if lineOfInfo[4] == "()":
                OurRef = lineOfInfo[5]
            else:
                OurRef = lineOfInfo[4]
            data = [PONUM, DATE, OurRef]

        if len(raw[j][0]) > 3:
            f = len(raw[j][0]) - 4
            lineOfInfo = raw[j][0][2].split("\n")
            PONUM = lineOfInfo[0]
            DATE = lineOfInfo[1]
            if lineOfInfo[3] == "()":
                OurRef = lineOfInfo[4]
            else:
                OurRef = lineOfInfo[3]
            while f > 0:
                c = 1
                lineOfInfo = raw[j][0][2 + c].split("\n")
                PONUM = PONUM + lineOfInfo[0]
                DATE = DATE + lineOfInfo[1]
                if lineOfInfo[3] == "()":
                    OurRef = OurRef + lineOfInfo[4]
                else:
                    OurRef = OurRef + lineOfInfo[3]
                f = f - 1
                c = c + 1
            data = [PONUM, DATE, OurRef]

        # ---------- 解析子项目 ----------
        subItems = []
        item = 0

        # raw[j+1] 是这一份 PO 对应的明细行
        while item < len(raw[j + 1]):
            row = raw[j + 1][item]
            if not row:
                item += 1
                continue

            # ====== 特判：这一行所有东西都挤在第 0 列 ======
            # 形如 "00950 4507971 Transportation Service 211 Trip"
            if (
                isinstance(row[0], str)
                and " " in row[0]
                and all((c is None or c == "") for c in row[1:])
            ):
                parts = row[0].split()
                # 至少要有：Item(5位) + Material(7位) + 描述... + 数量 + 单位
                if (
                    len(parts) >= 5
                    and parts[0].isdigit()
                    and parts[1].isdigit()
                    and parts[-2].isdigit()
                ):
                    itemNum_merged = parts[0]
                    material_merged = parts[1]
                    desc_merged = " ".join(parts[2:-2])
                    qty_merged = parts[-2]
                    uom_merged = parts[-1]
                    # 重写成标准格式，后面统一处理
                    row = [itemNum_merged, material_merged,
                           desc_merged, qty_merged, uom_merged]

            # 只处理以 "00" 开头的 item 行
            if not (isinstance(row[0], str) and row[0].startswith("00")):
                item += 1
                continue

            itemNum = row[0]
            material = row[1] if len(row) > 1 else ""
            description = ""
            quantity = ""
            UOM = ""

            # === 描述：本行从第3列开始，跳过空格，遇到纯数字就停止 ===
            k = 2
            while k < len(row):
                cell = row[k]
                if cell is None or cell == "":
                    k += 1
                    continue
                # 遇到纯数字，认为后面是数量/UOM
                if isinstance(cell, str) and cell.isdigit():
                    break
                description += cell
                k += 1

            # === 描述续行：如果下一行第一列为空，且不是公司信息/页眉，就拼接 ===
            if item + 1 < len(raw[j + 1]):
                nxt = raw[j + 1][item + 1]
                if nxt and len(nxt) > 0 and nxt[0] == "":
                    joined = "".join(c for c in nxt if c)
                    bad_keywords = [
                        "U Mobile Sdn Bhd",   # 带空格
                        "UMobile Sdn Bhd",    # 不带空格
                        "U Mobile",
                        "UMobile",
                        "Mobile Sdn Bhd",
                        "KUALA LU",
                        "Malaysia",
                        "JALAN TUN RAZA",
                        "JALAN TUN RAZAK",
                        "PAGE",
                        "LEVEL 18",
                        "SUITE 18-",
                        "G TOWER",
                        "199101013657",
                    ]
                    if not any(kwd in joined for kwd in bad_keywords):
                        # 只拼接前几列文字，避免把金额之类拼进去
                        for col in nxt[1:5]:
                            if col and (not isinstance(col, str) or not col.isdigit()):
                                description += col

            # === 数量：从右往左找第一个纯数字 ===
            qty_idx = None
            for idx in range(len(row) - 1, -1, -1):
                col = row[idx]
                if isinstance(col, str) and col.isdigit() and len(col) <= 4:
                    quantity = col
                    qty_idx = idx
                    break

            # === 单位：数量右边第一个非空非数字 ===
            if qty_idx is not None:
                for u in range(qty_idx + 1, len(row)):
                    cell = row[u]
                    if cell and (not isinstance(cell, str) or not cell.isdigit()):
                        UOM = cell
                        break

            subItems.append([itemNum, material, description, quantity, UOM])
            item += 1

        data.append(subItems)
        allData.append(data)
        j += 2  # 每个 PO 占用 raw[j], raw[j+1]

    return allData



# =========================
# 3. 主流程：读取 -> 解析 -> 导出 Excel
# =========================

folder_path = "input"
# 只拿 PDF 文件，防止夹带其它文件时报错
filesName = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

# 原始表格数据
raw_output = get_raw_info()

# 解析后的 PO + 子项目结构
output = get_info(raw_output)

# 如果想看看结构可以打开这行：
print(output)

print("开始生成.xlsx文件")
# 用当前时间做文件名
xlsxName = time.strftime('%Y%m%d%H%M%S', time.localtime()) + ".xlsx"

# 生成空文件
workbook = xlsxwriter.Workbook(xlsxName)
# 生成空工作表
worksheet = workbook.add_worksheet('PO Info')
worksheet.set_column('A:A', 10)
worksheet.set_column('B:B', 10)
worksheet.set_column('C:C', 12)
worksheet.set_column('D:D', 10)
worksheet.set_column('E:E', 10)
worksheet.set_column('F:F', 40)
worksheet.set_column('G:G', 10)
worksheet.set_column('H:H', 10)

# 向工作表输入表头
worksheet.write(0, 0, "PO Number")
worksheet.write(0, 1, "Date")
worksheet.write(0, 2, "Our Reference")
worksheet.write(0, 3, "Item")
worksheet.write(0, 4, "Material")
worksheet.write(0, 5, "Description")
worksheet.write(0, 6, "Quantity")
worksheet.write(0, 7, "UOM")

# 填数据
i = 1  # 当前 Excel 行
a = 0  # 第几个 PO

while a < len(output):
    # 写 PO 头部（这一行只写 A-C 列）
    worksheet.write(i, 0, output[a][0])
    worksheet.write(i, 1, output[a][1])
    worksheet.write(i, 2, output[a][2])

    q = 0              # 子项目索引
    c = i              # 当前写入行（随着子项目增加）

    while q < len(output[a][3]):
        item_row = output[a][3][q]
        worksheet.write(c, 3, item_row[0])  # Item
        worksheet.write(c, 4, item_row[1])  # Material
        worksheet.write(c, 5, item_row[2])  # Description
        worksheet.write(c, 6, item_row[3])  # Quantity
        worksheet.write(c, 7, item_row[4])  # UOM
        c += 1
        q += 1

    # 下一个 PO 的头部从子项目之后的下一行开始
    i = c + 1
    a += 1

workbook.close()
print("生成完毕，文件名为：" + xlsxName)
