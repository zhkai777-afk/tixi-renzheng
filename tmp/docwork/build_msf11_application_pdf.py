from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


OUT = "output/pdf/MSF11-01-认证申请表-演示预览版.pdf"

FONT = "MSF11CJK"
FONT_PATH = "/System/Library/Fonts/STHeiti Light.ttc"
pdfmetrics.registerFont(TTFont(FONT, FONT_PATH, subfontIndex=0))

PAGE_W, PAGE_H = A4
LEFT = RIGHT = 17 * mm
TOP = 22 * mm
BOTTOM = 18 * mm
CONTENT_W = PAGE_W - LEFT - RIGHT

BLUE = colors.HexColor("#245a8d")
BLUE_DARK = colors.HexColor("#173b60")
LIGHT_BLUE = colors.HexColor("#eef5fb")
BG = colors.HexColor("#f7f9fb")
BORDER = colors.HexColor("#d9e1ea")
HEADER_BG = colors.HexColor("#f1f5f9")
TEXT = colors.HexColor("#1f2937")
MUTED = colors.HexColor("#64748b")

styles = getSampleStyleSheet()
styles.add(
    ParagraphStyle(
        name="CN",
        fontName=FONT,
        fontSize=9.2,
        leading=13,
        textColor=TEXT,
        spaceAfter=3,
    )
)
styles.add(
    ParagraphStyle(
        name="TitleCN",
        fontName=FONT,
        fontSize=20,
        leading=24,
        alignment=TA_CENTER,
        textColor=BLUE_DARK,
        spaceAfter=8,
    )
)
styles.add(
    ParagraphStyle(
        name="Subtle",
        fontName=FONT,
        fontSize=8.3,
        leading=11,
        textColor=MUTED,
    )
)
styles.add(
    ParagraphStyle(
        name="Section",
        fontName=FONT,
        fontSize=12,
        leading=16,
        textColor=BLUE_DARK,
        spaceBefore=10,
        spaceAfter=6,
    )
)
styles.add(
    ParagraphStyle(
        name="Cell",
        fontName=FONT,
        fontSize=8.7,
        leading=12,
        textColor=TEXT,
    )
)
styles.add(
    ParagraphStyle(
        name="CellSmall",
        fontName=FONT,
        fontSize=8.1,
        leading=11,
        textColor=TEXT,
    )
)
styles.add(
    ParagraphStyle(
        name="Field",
        fontName=FONT,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#334155"),
    )
)
styles.add(
    ParagraphStyle(
        name="Note",
        fontName=FONT,
        fontSize=8,
        leading=11,
        textColor=MUTED,
        leftIndent=8,
    )
)


def p(text, style="CN"):
    return Paragraph(text, styles[style])


def section(title):
    return [
        Spacer(1, 3),
        Table(
            [[p(title, "Section")]],
            colWidths=[CONTENT_W],
            style=[
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT_BLUE),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#c9dff2")),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ],
        ),
        Spacer(1, 6),
    ]


def field_table(rows, widths=None, small=False):
    if widths is None:
        widths = [30 * mm, 55 * mm, 30 * mm, CONTENT_W - 115 * mm]
    data = []
    for row in rows:
        data.append([p(cell, "CellSmall" if small else "Cell") for cell in row])
    style = TableStyle(
        [
            ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]
    )
    for c in range(len(rows[0])):
        if c % 2 == 0:
            style.add("BACKGROUND", (c, 0), (c, -1), HEADER_BG)
            style.add("TEXTCOLOR", (c, 0), (c, -1), colors.HexColor("#334155"))
    return Table(data, colWidths=widths, style=style)


def checklist(items, cols=2):
    rows = []
    per_row = cols
    for i in range(0, len(items), per_row):
        row = items[i : i + per_row]
        while len(row) < per_row:
            row.append("")
        rows.append([p(("□ " + x) if x else "", "CellSmall") for x in row])
    return Table(
        rows,
        colWidths=[CONTENT_W / cols] * cols,
        style=[
            ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ],
    )


class DemoDoc(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(
            filename,
            pagesize=A4,
            leftMargin=LEFT,
            rightMargin=RIGHT,
            topMargin=TOP,
            bottomMargin=BOTTOM,
            title="MSF11-01 认证申请表 - 演示预览版",
            author="CQC Demo",
        )
        frame = Frame(LEFT, BOTTOM, CONTENT_W, PAGE_H - TOP - BOTTOM, id="normal")
        self.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=draw_page)])


def draw_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG)
    canvas.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
    canvas.setFillColor(colors.white)
    canvas.roundRect(LEFT - 5, BOTTOM - 6, CONTENT_W + 10, PAGE_H - TOP - BOTTOM + 12, 5, stroke=0, fill=1)
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.roundRect(LEFT - 5, BOTTOM - 6, CONTENT_W + 10, PAGE_H - TOP - BOTTOM + 12, 5, stroke=1, fill=0)
    canvas.setFillColor(BLUE_DARK)
    canvas.setFont(FONT, 8.5)
    canvas.drawString(LEFT, PAGE_H - 13 * mm, "MSF11-01 认证申请表")
    canvas.setFillColor(MUTED)
    canvas.drawRightString(PAGE_W - RIGHT, PAGE_H - 13 * mm, "演示预览版 | 2025-12-31 修订 | 2026-01-01 实施")
    canvas.setStrokeColor(BORDER)
    canvas.line(LEFT, PAGE_H - 16 * mm, PAGE_W - RIGHT, PAGE_H - 16 * mm)
    canvas.setFont(FONT, 8)
    canvas.setFillColor(MUTED)
    canvas.drawString(LEFT, 10 * mm, "文件编号：MSF11-01    版本：2015版")
    canvas.drawRightString(PAGE_W - RIGHT, 10 * mm, f"第 {doc.page} 页")
    canvas.restoreState()


story = []
story.append(p("认证申请表", "TitleCN"))
story.append(
    Table(
        [[p("文件编号：MSF11-01", "Subtle"), p("适用体系：Q / E / S / F", "Subtle"), p("状态：演示预览", "Subtle")]],
        colWidths=[CONTENT_W * 0.34, CONTENT_W * 0.33, CONTENT_W * 0.33],
        style=[
            ("BACKGROUND", (0, 0), (-1, -1), HEADER_BG),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ],
    )
)
story.append(Spacer(1, 9))

story += section("一、客户基本信息")
story.append(
    field_table(
        [
            ["客户名称", "根据登录人所属公司自动带入", "工作语言", "□ 汉语    □ 其他："],
            ["注册地址", "自动带入", "邮编", "自动带入"],
            ["运营地址", "自动带入", "邮编", "自动带入"],
            ["联系人", "可选择", "所在部门", "选择后自动带入"],
            ["职务", "选择后自动带入", "电话", "选择后自动带入，可修改，必填"],
            ["E-mail", "选择后自动带入，可修改，必填", "传真", "选择后自动带入，可修改，非必填"],
        ]
    )
)
story.append(Spacer(1, 8))
story.append(
    field_table(
        [
            ["申请认证所涉及的产品/服务/活动", "填写认证范围内的主要产品、服务或活动"],
            ["作息时间（管理人员）", "上午：              下午："],
            ["工作性质", "□ 常年工作    □ 季节性工作，季节："],
        ],
        widths=[48 * mm, CONTENT_W - 48 * mm],
    )
)

story += section("二、管理体系范围内人员与过程")
story.append(
    field_table(
        [
            ["人员总数", "共      人", "全职/兼职", "全职      人；兼职      人，兼职每日工作      小时"],
            ["外部场所服务人员", "共      人，从事：", "适用说明", "OHSMS 适用"],
            ["承包商/分包方人员", "共      人，从事：", "适用说明", "OHSMS 适用"],
            ["生产/服务班次", "共      班次，每班起止时间：", "控制方法", "□ 相同    □ 不同，请说明："],
            ["生产线情况", "共      条生产线；相同生产线      条", "涉及人数", "      人"],
            ["重复过程", "□ 有，过程：", "班组情况", "      个班组，每组      人"],
            ["临时非熟练人员", "□ 有，共      人", "从事工作", ""],
        ],
        small=True,
    )
)
story.append(Spacer(1, 8))
story.append(
    field_table(
        [
            ["食品包装形式", "是否采用影响食品安全的包装形式，如玻璃、真空包装、气调包装等：□ 否/不适用    □ 是，请说明："],
            ["外部仓库或加工步骤", "加工场所之外是否包括仓库或其他加工步骤：□ 否    □ 是，请说明："],
        ],
        widths=[48 * mm, CONTENT_W - 48 * mm],
    )
)

story += section("三、组织关系、外包与历史认证")
story.append(
    field_table(
        [
            ["上级/集团组织", "组织名称：                与本组织关系："],
            ["相关职能", "该组织负责的认证相关职能（采购、供应商批准、质量保证等）："],
            ["外包过程", "外包（外委）过程或业务及承担方："],
            ["医疗器械灭菌", "采用灭菌技术：        是否外包：□ 否 □ 是    是否通过 ISO13485：□ 否 □ 是"],
            ["其他机构认证", "认证机构名称：                    认证标准："],
            ["证书信息", "证书有效期：                    最近一次审核日期："],
            ["暂停/撤销说明", "如证书被暂停或撤销，请说明时间和原因："],
            ["体系运行", "管理体系开始运行时间：            最近一次内审时间："],
            ["咨询机构", "如管理体系在咨询机构帮助下建立，请填写咨询机构名称："],
            ["处罚/事故", "申请前一年内是否被处罚或发生相关事故/事件：□ 否    □ 是，请简述："],
        ],
        widths=[42 * mm, CONTENT_W - 42 * mm],
        small=True,
    )
)

story += section("四、认证需求")
story.append(p("申请认证标准", "CN"))
story.append(
    checklist(
        [
            "GB/T 19001-2016 / ISO 9001:2015",
            "GB/T 50430-2017",
            "GB/T 24001-2016 / ISO 14001:2015",
            "GB/T 45001-2020 / ISO 45001:2018",
            "GB/T 22000-____ / ISO 22000:____",
            "HACCP：____（其他技术规范：CAC《食品卫生通则》）",
            "乳制品 GMP",
            "GB/T 42061-____ / ISO 13485：____",
            "其他：",
            "组织 QMS 不适用 ISO 9001:2015 的要求，请注明条款号或具体要求：",
        ],
        cols=2,
    )
)
story.append(Spacer(1, 8))
story.append(
    field_table(
        [
            ["认可机构", "□ CNAS    □ DAKKS    □ UKAS"],
            ["认证类型", "□ 初次认证    □ 再认证    □ 认证转换"],
            ["认证转换信息", "认证机构、认证领域、认证标准、证书编号、首次发证日期、本周期发证日期、有效期、最近一次现场审核开始/结束时间等"],
            ["纸质证书", "□ 否    □ 是（默认电子证书；纸质证书另付费 50 元/页）"],
            ["希望现场审核时间", "      年      月"],
            ["多体系同时审核", "□ 否    □ 是"],
            ["其他要求", "其他对认证及审核的要求："],
        ],
        widths=[42 * mm, CONTENT_W - 42 * mm],
    )
)
story.append(Spacer(1, 6))
story.append(p("注：中心已通过多家认可机构多个认证领域的认可，详细认可范围可登录 CQC 官网查询或咨询项目管理人员。", "Note"))

story += section("五、组织管理体系整合程度（适用于多体系认证申请）")
story.append(p("如申请两个或两个以上体系，请根据组织实际情况勾选，并确定最终整合程度：______%。", "CN"))
integrated_rows = [
    [p("整合程度低（0%-40%）", "Cell"), p("整合程度中（40%-80%）", "Cell"), p("整合程度高（80%-100%）", "Cell")],
    [
        p("□ 分别建立管理体系<br/>□ 策划机制各不相同<br/>□ 管理评审各自进行<br/>□ 法律要求监视不一致<br/>□ 有不同的管理体系文件", "CellSmall"),
        p("□ 一定程度上建立整合管理体系<br/>□ 一个体系协调员和不同管理者代表<br/>□ 管理体系文件部分整合<br/>□ 文件和记录协调控制<br/>□ 管理评审一起进行", "CellSmall"),
        p("□ 一套整合文件及适度融合作业文件<br/>□ 结合经营战略和计划的管理评审<br/>□ 内部审核一体化<br/>□ 方针和目标一体化<br/>□ 体系过程、改进机制和职责一体化", "CellSmall"),
    ],
]
story.append(
    Table(
        integrated_rows,
        colWidths=[CONTENT_W / 3] * 3,
        style=[
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ],
    )
)

story += section("六、应附文件和资料")
story.append(p("通用申请资料", "CN"))
story.append(
    checklist(
        [
            "法律地位证明文件复印件；多场所组织应附每个场所证明文件及法律/合同安排",
            "相关资质文件复印件（适用时）",
            "认证场所清单（适用于多场所，如分公司、厂、办、处、所、站、项目部等）",
            "组织机构图；多场所组织需注明中心职能",
            "认证范围内所有生产/服务流程；多场所组织需注明流程适用场所",
            "管理体系过程信息（过程清单、过程图、工作流、乌龟图等）",
            "产品或服务质量标准清单（QMS、FSMS、HACCP、乳制品相关认证适用）",
            "有效版本的管理体系文件；多场所组织的体系文件构成信息",
            "文件、管理体系要素和职责整合的信息（适用于整合管理体系）",
            "原认证证书、初次/再认证/监督审核报告及不符合复印件（适用于认证转换）",
        ],
        cols=1,
    )
)
story.append(Spacer(1, 8))
story.append(p("申请 EMS、OHSMS 认证另需提供", "CN"))
story.append(
    checklist(
        [
            "重要环境因素",
            "主要危险源、OHS 风险及主要危险材料",
            "环评及三同时证明文件、排污许可证复印件（适用时）",
            "安评、职评及三同时证明文件复印件（适用时）",
            "近一年污染物排放监测报告/作业场所危害因素检测报告（适用时）",
            "消防法规符合性证明文件（适用于 OHSMS）",
            "需要应对的风险和机遇",
            "合规义务（适用法律法规及其他要求）清单",
        ],
        cols=2,
    )
)
story.append(Spacer(1, 8))
story.append(p("申请 FSMS、HACCP、乳制品 GMP 认证另需提供", "CN"))
story.append(
    checklist(
        [
            "食品安全管理体系/HACCP 体系文件（产品描述、工艺流程图、危害分析、HACCP 计划等）",
            "厂区位置图、平面图；加工车间平面图",
            "食品添加剂使用情况说明（名称、用量、适用产品及限量标准等）",
            "食品认证申请组织自我声明",
            "产品符合卫生安全要求的检测报告",
            "生鲜乳来源、最大收奶半径及生鲜乳日供应量（仅适用于乳制品 GMP）",
        ],
        cols=1,
    )
)

story += section("七、声明与签署")
story.append(p("声明：我方确认以上提供的信息（包括资料）均属实。", "CN"))
story.append(
    Table(
        [
            [p("客户代表签字：", "Cell"), p("", "Cell"), p("公章：", "Cell")],
            [p("日期：        年      月      日", "Cell"), p("", "Cell"), p("", "Cell")],
        ],
        colWidths=[CONTENT_W * 0.34, CONTENT_W * 0.28, CONTENT_W * 0.38],
        rowHeights=[22 * mm, 16 * mm],
        style=[
            ("GRID", (0, 0), (-1, -1), 0.45, BORDER),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fbfdff")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
        ],
    )
)
story.append(Spacer(1, 8))
story.append(
    p(
        "如需进一步了解 CQC 简介、认证业务、收费标准、权利义务、申诉投诉渠道、获证方公告方式、体系变更通报要求及 IQNet 认证合作简介，请查阅 CQC 官网公开文件目录或向相关人员索取。",
        "Note",
    )
)
story.append(Spacer(1, 6))
story.append(p("注解：", "CN"))
for note in [
    "重复过程是指认证范围内较高比例人员从事某项重复活动/工作，如保洁、运送、销售、呼叫中心等。",
    "管理体系过程信息可按组织自己的方式提供；外部场所服务、承包方人员如涉及多个场所，应分别说明人数和服务。",
]:
    story.append(p(note, "Note"))


if __name__ == "__main__":
    doc = DemoDoc(OUT)
    doc.build(story)
    print(OUT)
