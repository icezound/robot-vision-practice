# -*- coding: utf-8 -*-
"""Fill the practical training report template and build the final DOCX."""
import os
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TPL = r"D:\matlab chuchun\w\_template.docx"
ASSETS = r"D:\matlab chuchun\w\_assets"
OUT = r"D:\matlab chuchun\w\无锡太湖学院实训报告-机器人视觉技术与应用实践-完成版.docx"

SONG = "宋体"
HEI = "黑体"
TNR = "Times New Roman"


# ----------------------------------------------------------------- helpers
def set_run(run, size=12, bold=False, east=SONG, west=TNR):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), west)
    rf.set(qn("w:hAnsi"), west)
    rf.set(qn("w:eastAsia"), east)
    rf.set(qn("w:cs"), west)
    rf.set(qn("w:hint"), "eastAsia")


def fmt_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent_chars=None,
             line=1.2, before=0, after=0, keep_next=False, snap=False):
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = keep_next
    pPr = p._p.get_or_add_pPr()
    if not snap and pPr.find(qn("w:snapToGrid")) is None:
        el = OxmlElement("w:snapToGrid")
        el.set(qn("w:val"), "0")
        pPr.append(el)
    if indent_chars:
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(indent_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(indent_chars * 12 * 20)))
    return p


def body(cell, text, indent=2.0):
    """Regular body paragraph, 宋体小四, first line indented by 2 characters."""
    p = cell.add_paragraph()
    fmt_para(p, indent_chars=indent)
    set_run(p.add_run(text), size=12, east=SONG)
    return p


def h1(cell, text):
    p = cell.add_paragraph()
    return h1_para(p, text)


def h1_para(p, text):
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
    set_run(p.add_run(text), size=14, bold=True, east=HEI)
    return p


def h2(cell, text):
    p = cell.add_paragraph()
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
    set_run(p.add_run(text), size=12, bold=True, east=HEI)
    return p


def caption(cell, text, keep_next=False):
    """Add a caption. When the previous block is a figure block, the caption is
    appended after a line break so figure and caption never separate."""
    from docx.text.paragraph import Paragraph

    kids = [c for c in cell._element.iterchildren()
            if c.tag in (qn("w:p"), qn("w:tbl"))]
    if kids and kids[-1].tag == qn("w:tbl"):
        target = None
        for p in kids[-1].findall(".//" + qn("w:p")):
            if p.findall(".//" + qn("w:drawing")):
                target = p
        if target is not None:
            br = OxmlElement("w:r")
            br.append(OxmlElement("w:br"))
            target.append(br)
            par = Paragraph(target, cell)
            set_run(par.add_run("  " + text), size=10.5, east=SONG)
            return par
    p = cell.add_paragraph()
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.2, keep_next=keep_next)
    set_run(p.add_run(text), size=10.5, east=SONG)
    return p


def keep_block(container, width_cm=15.6):
    """A borderless single-cell table whose row cannot split across pages."""
    t = container.add_table(rows=1, cols=1)
    # python-docx appends a blank paragraph after a nested table; drop it so the
    # figure blocks do not add an extra empty line each time
    tail = list(container._element.iterchildren())[-1]
    if (tail.tag == qn("w:p") and not "".join(tail.itertext()).strip()
            and not tail.findall(".//" + qn("w:drawing"))):
        container._element.remove(tail)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    c = t.rows[0].cells[0]
    c.width = Cm(width_cm)
    t.columns[0].width = Cm(width_cm)
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", 0), ("left", 0), ("bottom", 0), ("right", 0)):
        e = OxmlElement("w:" + side)
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
    p = c.paragraphs[0]
    return c, p


def figure(container, path, width_cm=15.0):
    """Figure with its caption inside one unbreakable block."""
    c, p = keep_block(container)
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(os.path.join(ASSETS, path), width=Cm(width_cm))
    return p


def _insert_ordered(parent, element, successors):
    """Insert element before the first successor tag present (OOXML order)."""
    for tag in successors:
        found = parent.find(qn(tag))
        if found is not None:
            found.addprevious(element)
            return
    parent.append(element)


def three_line_table(table, color="000000", top_sz="12", mid_sz="6"):
    """Turn a grid table into a three-line (booktabs style) table."""
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "bottom"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), top_sz)
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), color)
        borders.append(e)
    for edge in ("left", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    _insert_ordered(tblPr, borders,
                    ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"))

    # rule under the header row
    for c in table.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcBorders")):
            tcPr.remove(old)
        tb = OxmlElement("w:tcBorders")
        e = OxmlElement("w:bottom")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), mid_sz)
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), color)
        tb.append(e)
        _insert_ordered(tcPr, tb, ("w:shd", "w:noWrap", "w:tcMar", "w:vAlign"))


# ----------------------------------------------------------------- document
doc = docx.Document(TPL)
table = doc.tables[0]
cell = table.rows[0].cells[0]

paras = cell.paragraphs
first = paras[0]
for p in paras[1:]:
    p._p.getparent().remove(p._p)
for t in cell.tables:
    t._tbl.getparent().remove(t._tbl)
for r in list(first.runs):
    r._r.getparent().remove(r._r)
first.style = doc.styles["Normal"]

# move the cover section break onto the last cover line so the trailing empty
# paragraphs no longer push the cover onto a second, blank page
body_el = doc.element.body
children = list(body_el.iterchildren())
sect_p = None
for el in children:
    if el.tag == qn("w:p") and el.find(".//" + qn("w:sectPr")) is not None:
        sect_p = el
        break
if sect_p is not None:
    idx = children.index(sect_p)
    target = None
    for j in range(idx - 1, -1, -1):
        if children[j].tag == qn("w:p") and "".join(children[j].itertext()).strip():
            target = children[j]
            break
    if target is not None:
        sectPr = sect_p.find(".//" + qn("w:sectPr"))
        target.find(qn("w:pPr")).append(sectPr)
        for j in range(children.index(target) + 1, idx + 1):
            body_el.remove(children[j])

# --- 一、实训目的及要求 -------------------------------------------------
h1_para(first, "一、实训目的及要求")
h2(cell, "（一）实训目的")
body(cell, "本次实训以机器人视觉技术的工程应用为背景，围绕工业机器视觉软件 HALCON、"
           "开源计算机视觉库 OpenCV 与深度学习框架 PyTorch 三条技术路线组织内容，"
           "通过上机操作完成图像读取、图像预处理、边缘提取、区域分析与开发环境验证等任务，"
           "熟悉机器视觉系统从图像输入到特征输出的基本流程。")
body(cell, "通过本次实训，应能够独立完成机器视觉软件的安装与开发环境配置，"
           "理解图像滤波、灰度变换、二值化、边缘检测和 Blob 分析等基本算法的原理与作用，"
           "能够根据检测目标选择合适的算子与参数，并对处理结果作出正确判断，"
           "为后续学习机器人视觉定位、尺寸测量与目标识别打下基础。")
h2(cell, "（二）实训要求")
body(cell, "1．掌握 HALCON（HDevelop）的安装与基本操作，能够使用 read_image、mean_image、"
           "sobel_amp、threshold、gray_histo、connection、select_shape、area_center 等算子"
           "完成一幅图像的读取、降噪、边缘提取与区域特征提取，并读懂变量窗口中图像变量与"
           "控制变量的含义。")
body(cell, "2．掌握 Python 与 OpenCV 的开发环境配置方法，能够使用 cv2.imread、cv2.cvtColor、"
           "cv2.blur、cv2.GaussianBlur、cv2.medianBlur、cv2.threshold、cv2.Canny、"
           "cv2.findContours 等函数完成图像的灰度化、滤波、灰度变换、二值化、边缘检测与轮廓绘制，"
           "并能对照说明处理前后图像发生的变化。")
body(cell, "3．了解深度学习的基本概念，掌握 PyTorch 的安装方法与运行环境的检测方法，"
           "能够通过 torch.__version__、torch.version.cuda、torch.cuda.is_available() 等接口"
           "确认框架版本以及 GPU 加速是否可用。")
body(cell, "4．实训过程中按要求记录每一步的操作与结果，保存原图与处理后图像，"
           "并在实训报告中以对比图的形式加以说明。报告要求文字通顺、无错别字，"
           "图表编号规范，量和单位符合国家标准。")
body(cell, "上述三个内容的操作过程与运行结果都需要在报告中如实记录："
           "软件安装与调试的关键界面以及程序代码以插图的形式给出，"
           "图像处理前后的对比结果统一放在第四部分说明。"
           "实训过程中既要保证程序能够正常运行，"
           "也要能够说明每一步处理的作用以及参数变化对结果的影响，"
           "养成边操作、边记录、边分析的习惯。")

# --- 二、实训设备与仪器 -------------------------------------------------
h1(cell, "二、实训设备与仪器")
body(cell, "本次实训在机房计算机上完成，全部处理对象为同一幅实训图像（文件名为 "
           "80955976215C036F5273844FFB2C87AD.png，画面为两名卡通人物），"
           "所用硬件与软件的名称及版本信息如表2-1所示。")
rows = [
    ("类别", "名称", "版本／说明", "用途"),
    ("硬件", "计算机", "PC 机，Windows 操作系统", "运行 HALCON 与 Python 程序"),
    ("软件", "MVTec HALCON", "实训机房安装版本，HDevelop 开发环境",
     "图像读取、滤波、边缘提取与 Blob 分析"),
    ("软件", "Python", "PyCharm 创建的虚拟环境 .venv", "编写并运行图像处理程序"),
    ("软件", "OpenCV（cv2）", "随 Python 环境安装", "灰度化、滤波、二值化、边缘检测与轮廓提取"),
    ("软件", "NumPy", "随 Python 环境安装", "图像矩阵运算、对数变换与幂次变换"),
    ("软件", "PyTorch", "2.14.0+cpu", "深度学习框架版本与环境检测"),
    ("软件", "CUDA", "torch.version.cuda 返回 None，未启用", "GPU 加速（本次实训以 CPU 模式运行）"),
    ("软件", "PyCharm", "实训机房安装版本", "代码编辑、调试与运行"),
    ("软件", "Microsoft Word", "2010 及以上", "实训报告撰写与排版"),
]
wrap, cap_p = keep_block(cell, 15.6)
fmt_para(cap_p, align=WD_ALIGN_PARAGRAPH.CENTER)
set_run(cap_p.add_run("表2-1 实训所用软硬件及版本信息"), size=10.5, east=SONG)
tbl = wrap.add_table(rows=len(rows), cols=4)
try:
    tbl.style = doc.styles["Table Normal"]
except KeyError:
    tbl.style = doc.styles["Table Grid"]
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.autofit = False
widths = [Cm(1.8), Cm(3.4), Cm(4.6), Cm(5.2)]
for ri, row in enumerate(rows):
    for ci, text in enumerate(row):
        c = tbl.cell(ri, ci)
        c.width = widths[ci]
        p = c.paragraphs[0]
        fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT if ci == 3 else WD_ALIGN_PARAGRAPH.CENTER)
        set_run(p.add_run(text), size=10.5, bold=(ri == 0), east=SONG)
three_line_table(tbl)
trPr = tbl.rows[0]._tr.get_or_add_trPr()
trPr.append(OxmlElement("w:tblHeader"))

# --- 三、实训内容与步骤 -------------------------------------------------
h1(cell, "三、实训内容与步骤")
body(cell, "本次实训共安排三个内容：内容一为 HALCON 的认识与应用，内容二为图像处理技术 "
           "OpenCV+Python，内容三为深度学习初认识，即 PyTorch 的认识与应用。"
           "三个内容使用同一幅实训图像，便于比较不同软件的处理思路与结果。"
           "按照要求，项目结果的对比截图统一放在第四部分，本部分主要说明所使用的软件、"
           "环境配置与调试过程以及程序的项目内容与要求。")

h2(cell, "3.1 内容一：HALCON 的认识与应用")
body(cell, "（1）软件介绍。HALCON 是德国 MVTec 公司开发的机器视觉算法包，"
           "其集成开发环境 HDevelop 把程序编辑、图形显示与变量观察集中在同一个界面中，"
           "算子以“算子名（输入参数，输出参数，控制参数）”的形式书写，"
           "执行后可以立即在图形窗口中查看结果，适合快速验证图像处理方案。")
body(cell, "（2）安装与调试过程。实训使用机房已安装好的 HALCON，启动 HDevelop 后新建程序文件，"
           "在程序窗口中按顺序输入算子并逐行执行。检查程序是否正常的方法是："
           "每执行一条算子，观察变量窗口中新出现的图像变量，"
           "并在变量窗口的缩略图上停留鼠标以查看该变量的类型与尺寸信息；"
           "如果算子输入输出参数个数或类型不匹配，程序编辑器会在出错行给出标记，"
           "需要根据提示修改参数后重新执行。需要注意的是，读取图像时要给出完整的绝对路径，"
           "并把路径中的分隔符写成斜杠或双反斜杠。")
body(cell, "（3）项目内容与要求。课堂给出的项目要求是：读取实训图像后先做均值滤波抑制噪声，"
           "再用 Sobel 算子提取边缘幅值，对幅值图像作阈值分割得到边缘区域，"
           "然后用 Blob 分析提取各连通区域的面积与中心坐标，"
           "同时统计图像的灰度直方图。据此编写的程序共分五步，如图3-1所示。")
figure(cell, "f3_1.png", 13.5)
caption(cell, "图3-1 HALCON 程序代码")
body(cell, "图3-1中的程序依次使用了八个算子。read_image 读取图像并存入图像变量 Image；"
           "mean_image 用 3×3 的掩膜对图像作均值滤波，输出噪声被抑制的 ImageFiltered；"
           "sobel_amp 以 'sum_abs' 方式按 3×3 邻域计算一阶梯度的绝对值之和，"
           "得到边缘幅值图像 EdgeAmplitude；threshold 以 20 为下限、255 为上限"
           "对幅值图像作阈值分割，得到边缘区域 EdgeBinary；gray_histo 统计图像的灰度分布，"
           "输出绝对频数 AbsoluteHisto 与相对频数 RelativeHisto；"
           "connection 按连通关系把 EdgeBinary 分割为互不相连的 ConnectedRegions；"
           "select_shape 以面积（area）为特征、下限 100、上限 99999 筛选区域，"
           "得到 ValidRegions；area_center 计算 ValidRegions 中每个区域的面积 Area "
           "与中心坐标 Row、Column。")

h2(cell, "3.2 内容二：图像处理技术 OpenCV+Python")
body(cell, "（1）软件介绍。OpenCV 是开源的计算机视觉库，提供大量以函数形式封装的图像处理算法；"
           "NumPy 用于把图像当作矩阵进行运算。二者配合 Python 使用，"
           "可以在几行代码内完成图像的读取、变换与显示。")
body(cell, "（2）安装与调试过程。在 PyCharm 中新建工程并创建虚拟环境 .venv，"
           "在终端中执行 pip install opencv-python numpy 完成安装；"
           "随后在程序中写入 import cv2 与 import numpy as np，"
           "如果运行时不出现模块找不到的错误，说明安装成功。"
           "调试时每完成一步处理就用 cv2.imshow 单独开窗显示结果，"
           "并用 cv2.waitKey(0) 与 cv2.destroyAllWindows() 控制窗口的停留与关闭，"
           "便于逐步骤观察图像的变化。")
body(cell, "（3）项目内容与要求。程序要求完成五步处理：读取原图并转换为灰度图；"
           "图像滤波，分别用均值滤波、高斯滤波与中值滤波比较不同滤波方法的效果；"
           "对数变换与伽马（指数）变换，观察灰度映射对图像明暗的影响；"
           "二值化并用 Canny 算子作边缘检测；最后进行图像分割并绘制轮廓。"
           "程序代码如图3-2所示。")
figure(cell, "f3_2.png", 14.0)
caption(cell, "图3-2 OpenCV 图像处理程序代码")
body(cell, "图3-2中的程序与上述五个步骤对应。读取图像后先用 cv2.cvtColor 完成灰度转换，"
           "再用 cv2.blur、cv2.GaussianBlur 和 cv2.medianBlur 分别实现均值滤波、高斯滤波与中值滤波；"
           "对数变换按 np.uint8(25*np.log(1+gray)) 计算，"
           "伽马变换按 np.uint8(255*(gray/255)**gamma_val) 计算，其中 gamma_val 取 1.2；"
           "二值化使用 cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)，"
           "边缘检测使用 cv2.Canny(gray, 50, 150)；"
           "最后用 cv2.findContours 以 RETR_TREE 方式提取轮廓，用于后续绘制。")

h2(cell, "3.3 内容三：深度学习初认识——PyTorch 的认识与应用")
body(cell, "（1）软件介绍。PyTorch 是基于张量运算与自动求导机制的深度学习框架，"
           "既可以把它当作支持 GPU 加速的 NumPy 使用，也可以用它搭建和训练神经网络。"
           "搭建深度学习环境时，需要先确认框架版本以及 GPU 加速是否可用，"
           "本次实训即围绕这一环节展开。")
body(cell, "（2）安装与调试过程。在 PyCharm 工程使用的虚拟环境中执行 pip install torch，"
           "安装完成后编写一段环境检测程序，通过框架提供的接口读取版本信息与计算设备信息，"
           "并在终端中查看输出结果，判断当前环境能否使用 GPU 加速。")
body(cell, "（3）项目内容与要求。程序要求输出 PyTorch 的版本号、CUDA 版本号以及 CUDA 是否可用；"
           "如果显卡可用，则进一步打印显卡名称与显卡数量，否则提示当前使用 CPU 模式。"
           "程序代码如图3-3所示。")
figure(cell, "f3_3.png", 13.5)
caption(cell, "图3-3 PyTorch 环境检测程序代码")
body(cell, "图3-3中的程序先导入 torch，然后依次使用 torch.__version__ 读取框架版本，"
           "使用 torch.version.cuda 读取所绑定的 CUDA 版本，"
           "使用 torch.cuda.is_available() 判断 GPU 是否可用；"
           "当判断结果为真时调用 torch.cuda.get_device_name(0) 与 torch.cuda.device_count() "
           "输出显卡名称与数量，否则输出提示信息并使用 CPU 完成计算。")

# --- 四、实训结果 -------------------------------------------------------
h1(cell, "四、实训结果")
body(cell, "三个内容的运行结果按顺序整理如下。为了便于对照，"
           "下面给出的每幅图都把原图与处理后的图像并列显示，"
           "并说明处理后图像发生的变化。")

h2(cell, "4.1 内容一：HALCON 的图像识别结果")
body(cell, "HALCON 程序执行完成后，图形窗口与变量窗口中的结果如图4-1至图4-4所示。")
figure(cell, "f4_1.png", 13.5)
caption(cell, "图4-1 原图与均值滤波后边缘幅值图像的对比")
body(cell, "由图4-1可见，边缘幅值图像以黑色为背景，原图中人物的轮廓、发丝、衣物边缘"
           "以及两处文字笔画的位置都出现了明亮的线条。sobel_amp 计算的是相邻像素的灰度差，"
           "灰度变化越剧烈的位置幅值越大，因此亮线集中出现在原图中颜色突变的位置；"
           "头发内部的色块边界同样被提取出来。画面背景因灰度变化平缓而保持黑色，"
           "说明 3×3 均值滤波在抑制噪声的同时保留了主要边缘。")
figure(cell, "f4_2.png", 13.5)
caption(cell, "图4-2 原图与边缘二值化区域的对比")
body(cell, "由图4-2可见，对边缘幅值图像以 20 为下限作阈值分割后，边缘被表示为区域。"
           "图中红色部分即为 EdgeBinary，其形状与图4-1中的亮线一致，"
           "但只保留了幅值大于阈值的像素。把鼠标停留在变量窗口的 EdgeBinary 缩略图上，"
           "可以读出该区域的面积为 44286 像素、中心点为 (242.83, 259.365)，"
           "说明阈值分割把分散的边缘像素连成了一片可供后续统计的区域。"
           "阈值下限的取值直接影响结果：取值过大会丢失细节，取值过小则会引入噪声。")
figure(cell, "f4_3.png", 13.5)
caption(cell, "图4-3 连通区域与面积筛选后有效区域的对比")
body(cell, "由图4-3可见，connection 算子按照 8 连通关系把 EdgeBinary 划分成了多个"
           "互不相连的区域，图中不同颜色代表不同的 ConnectedRegions，"
           "主体轮廓集中在一个面积最大的区域中。"
           "在此基础上用 select_shape 以面积 100 至 99999 为条件筛选，得到 ValidRegions。"
           "两幅图对比后可以看出，筛选只去除了少量面积很小的噪声区域，"
           "人物的主要轮廓区域全部保留，说明所设定的面积下限是合适的。")
figure(cell, "f4_4.png", 14.0)
caption(cell, "图4-4 灰度直方图与控制变量数据")
body(cell, "图4-4的控制变量窗口中列出了程序计算得到的两类数据。"
           "AbsoluteHisto 与 RelativeHisto 分别是各灰度级的绝对频数与相对频数，"
           "反映图像灰度的分布情况，其中相对频数的最大值约为 0.0007，"
           "说明图像中每种灰度的像素所占比例都很小、灰度分布较为分散。"
           "Area、Row、Column 三个数组分别记录了筛选后各有效区域的面积及其中心的"
           "行、列坐标，例如第一个区域的面积为 909 像素，中心列坐标为 62.90、行坐标为 80.29；"
           "面积数组中最大值为 33746 像素，对应人物整体的外轮廓区域，"
           "它也是 Blob 分析中用于定位目标的主要依据。")

h2(cell, "4.2 内容二：OpenCV+Python 的图像处理结果")
body(cell, "OpenCV 程序按照灰度化、滤波、灰度变换、二值化与边缘检测、轮廓绘制的顺序执行，"
           "各步骤的处理效果如图4-5至图4-9所示。")
figure(cell, "f4_5.png", 13.5)
caption(cell, "图4-5 原图与灰度图像的对比")
body(cell, "由图4-5可见，原图为彩色图像，经过 cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) 转换后"
           "变为单通道的灰度图像。灰度值按 0.299R+0.587G+0.114B 的加权公式计算，"
           "原图中浅紫色的头发、蓝紫色的衣物和红色的描边分别被转换成深浅不同的灰色，"
           "画面的明暗层次与原图保持一致，但不再含有颜色信息。"
           "灰度化把需要处理的数据量减少到原来的三分之一，"
           "同时保留了后续处理所需的亮度信息。")
figure(cell, "f4_6.png", 13.5)
caption(cell, "图4-6 灰度图像与对数变换后图像的对比")
body(cell, "由图4-6可见，经过对数变换后图像整体明显变暗，"
           "原图中占面积最大的亮色背景被压缩到接近零的低灰度区，"
           "而原本灰度较低的笔画、描边和文字则被拉开层次，在暗色背景上反而更加清晰。"
           "这是因为对数变换对低灰度区间有放大作用、对高灰度区间有压缩作用，"
           "适合处理整体偏亮、细节集中在中低灰度区的图像，"
           "可以在保留暗部细节的同时降低背景的影响。")
figure(cell, "f4_7.png", 13.5)
caption(cell, "图4-7 灰度图像与二值化图像的对比")
body(cell, "由图4-7可见，经过 cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY) 处理后，"
           "图像只保留了黑白两种取值：灰度大于阈值的像素被置为 255，其余被置为 0。"
           "原图中较亮的背景与人物身体连成白色区域，"
           "较暗的文字、描边与阴影则成为黑色线条，"
           "对比度低、灰度接近阈值的部分出现了锯齿状的边界。"
           "二值化把灰度图像转换为只有两类取值的图像，"
           "使目标的形状被突出出来，为后续的轮廓提取做好了准备；"
           "阈值取 127 是一个折中值，阈值偏大或偏小都会使黑白区域的比例明显改变。")
figure(cell, "f4_8.png", 13.5)
caption(cell, "图4-8 灰度图像与 Canny 边缘检测结果的对比")
body(cell, "由图4-8可见，cv2.Canny(gray, 50, 150) 的处理结果是一幅黑底白线的图像，"
           "人物的轮廓、发丝、五官、相机外形和文字笔画都被提取为宽度约一个像素的细线，"
           "背景区域不再有任何响应。Canny 算子采用双阈值判别，"
           "梯度大于高阈值的像素被判为强边缘，"
           "介于低阈值与高阈值之间的像素只有与强边缘相连时才被保留，"
           "因此边缘定位准确、连续性较好。"
           "与灰度图像相比，处理结果把“灰度是多少”变成了“灰度在哪里发生变化”，"
           "这正是后续区域分割与轮廓提取所需要的信息。")
figure(cell, "f4_9.png", 13.5)
caption(cell, "图4-9 原图与轮廓检测结果的对比")
body(cell, "由图4-9可见，在二值图像上使用 cv2.findContours 提取出各条边界轮廓后，"
           "再用 cv2.drawContours 以红色、线宽 2 把轮廓绘制到原图上，"
           "人物的身体轮廓、头发与衣物的分界、相机的形状以及两处文字都被红色线条勾出，"
           "图像最外层的边框也被当作一条轮廓绘制了出来。"
           "轮廓绘制把二值图像中包含的位置信息重新叠加到原图上，"
           "可以直观地检验分割结果与目标是否一致，"
           "也是后续计算目标外接矩形、面积和周长等几何特征的基础。")

h2(cell, "4.3 内容三：PyTorch 环境检测结果")
body(cell, "环境检测程序在 PyCharm 中运行，输出结果如图4-10所示。")
figure(cell, "f4_10.png", 15.0)
caption(cell, "图4-10 PyTorch 环境检测程序的运行输出")
body(cell, "由图4-10可见，终端依次输出了 Pytorch 版本为 2.14.0+cpu，"
           "torch.version.cuda 的返回值为 None，"
           "torch.cuda.is_available() 的返回值为 False，"
           "程序因此进入 else 分支并打印“当前使用CPU模式”。"
           "这说明当前环境中安装的是 CPU 版本的 PyTorch，"
           "没有可用的 CUDA 加速设备，程序能够正常运行但只能使用 CPU 完成计算。"
           "如果需要使用 GPU 加速，应安装与显卡驱动相匹配的 CUDA 版本，"
           "并重新安装对应的 GPU 版本 PyTorch。")

# --- 五、实训总结 -------------------------------------------------------
h1(cell, "五、实训总结")
body(cell, "本次实训把同一幅图像放到三种不同的工具中处理，"
           "让我比较直观地看到了机器视觉中“读图、处理、分析”这条主线。"
           "在 HALCON 中，我学会了用算子搭出一条完整的处理链："
           "读取图像后用均值滤波抑制噪声，用 sobel_amp 提取边缘幅值，"
           "用 threshold 把幅值图像变成区域，再用 connection、select_shape 和 area_center "
           "把区域拆开、筛选并量化。"
           "变量窗口里那一串面积与中心坐标让我第一次意识到，"
           "视觉系统最终交给机器人的并不是图像，而是一组可以参与运算的数字。")
body(cell, "在 OpenCV 与 Python 部分，我体会到开源库的灵活与“看得见过程”的好处。"
           "灰度化、滤波、对数变换、二值化、边缘检测和轮廓绘制每一步都能单独开窗显示，"
           "参数一变结果立刻不同。"
           "比如二值化阈值取 127 时黑白区域的比例已经比较合理，"
           "阈值再高一点人物的浅色身体就会与背景连成一片，"
           "这也让我理解了为什么实际项目中阈值常常需要结合直方图来确定。"
           "通过 PyTorch 的环境检测，我学会了先确认框架版本与计算设备再开始写代码，"
           "torch.cuda.is_available() 返回 False 时程序会自动切换到 CPU 模式，"
           "这种“先检查环境、再运行任务”的习惯对以后做项目很有帮助。")
body(cell, "回顾整个过程，我也发现了自己的不足。"
           "一是对算子和函数的参数意义理解得还不够透彻，"
           "不少参数是照着例子填的，比如 Sobel 邻域取 3、Canny 的高低阈值取 50 和 150，"
           "当时并没有认真比较过不同取值的效果；"
           "二是处理过程中缺少系统的记录，截图的时间顺序与程序的执行顺序不完全对应，"
           "整理报告时不得不反复核对；"
           "三是对结果的评价还停留在“看起来对不对”的层面，"
           "没有进一步统计面积、中心坐标等特征来判断结果的稳定性。")
body(cell, "针对这些不足，我打算从三个方面改进："
           "第一，对滤波、二值化、边缘检测中的关键参数做对比实验，"
           "把不同参数下的结果整理成表格，总结参数的取值范围与影响规律；"
           "第二，养成按步骤截图的习惯，为每一步结果标注处理名称与参数，"
           "让实验记录可以直接支撑报告的编写；"
           "第三，在完成基本流程之后，尝试把 HALCON 与 OpenCV 的处理结果放在一起定量比较，"
           "并把区域特征用于简单的目标定位，"
           "让视觉处理真正与机器人的运动控制衔接起来。")

# --- 页面设置与页码 -----------------------------------------------------
section = doc.sections[1]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
pgMar = section._sectPr.find(qn("w:pgMar"))
pgMar.set(qn("w:gutter"), "283")   # 装订线 0.5 cm

footer = section.footer
txbx = next(footer._element.iter(qn("w:txbxContent")), None)
if txbx is not None:
    fp = txbx.findall(qn("w:p"))[0]
    for child in list(fp):
        if child.tag != qn("w:pPr"):
            fp.remove(child)
    pPr = fp.find(qn("w:pPr"))
    if pPr is not None:
        for el in pPr.findall(qn("w:rPr")):
            pPr.remove(el)

    def add_text(txt):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), TNR)
        rf.set(qn("w:hAnsi"), TNR)
        rf.set(qn("w:eastAsia"), SONG)
        rf.set(qn("w:hint"), "eastAsia")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "21")
        rPr.append(rf)
        rPr.append(sz)
        r.append(rPr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = txt
        r.append(t)
        fp.append(r)

    def add_field():
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), " PAGE ")
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), TNR)
        rf.set(qn("w:hAnsi"), TNR)
        rf.set(qn("w:eastAsia"), SONG)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "21")
        rPr.append(rf)
        rPr.append(sz)
        r.append(rPr)
        t = OxmlElement("w:t")
        t.text = "1"
        r.append(t)
        fld.append(r)
        fp.append(fld)

    add_text("第 ")
    add_field()
    add_text(" 页")

# keep the paragraph that follows the report table on the last page instead of
# letting it spill onto a trailing blank page
from docx.text.paragraph import Paragraph

kids = list(doc.element.body.iterchildren())
pos = kids.index(table._tbl)
if pos + 1 < len(kids) and kids[pos + 1].tag == qn("w:p") and len(kids[pos + 1]) <= 1:
    tail_p = Paragraph(kids[pos + 1], doc.element.body)
    fmt_para(tail_p, align=WD_ALIGN_PARAGRAPH.LEFT)
    tail_p.paragraph_format.line_spacing = Pt(1)

doc.save(OUT)
print("saved:", OUT)
