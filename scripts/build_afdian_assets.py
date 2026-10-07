"""Build the local Afdian review assets and buyer ZIP; never uploads or publishes."""
from pathlib import Path
import hashlib
import json
import shutil
import zipfile
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
SALE = ROOT / 'sales/爱发电'
BUYER = SALE / '交付内容'
IMAGES = SALE / '预览图'
FONT = next((p for p in [
    Path('/System/Library/Fonts/Hiragino Sans GB.ttc'),
    Path('/System/Library/Fonts/STHeiti Medium.ttc'),
    Path('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'),
] if p.exists()), None)
if FONT is None:
    raise SystemExit('Chinese font missing; install a CJK font or adjust FONT candidates.')

def font(size):
    return ImageFont.truetype(str(FONT), size)

def text(draw, xy, content, size, fill):
    draw.text(xy, content, font=font(size), fill=fill)

def fit_check(draw, content, size, width):
    if draw.textbbox((0, 0), content, font=font(size))[2] > width:
        raise ValueError(f'Text exceeds available width: {content}')

def cover():
    image = Image.new('RGB', (1600, 1000), '#F4F6F2')
    d = ImageDraw.Draw(image)
    d.rounded_rectangle((70, 58, 520, 120), radius=20, fill='#163D34')
    text(d, (98, 72), '电脑表格 / v0.3 试用版', 31, '#FFFFFF')
    text(d, (82, 175), '一款商品，三套方案', 74, '#172F29')
    text(d, (84, 285), '把成本与退货损耗算在一起', 52, '#395B50')
    cards = [
        (82, '01', '先归集账单', '总额换算为每单金额'),
        (566, '02', '再比较场景', '售价、运费、广告、退货'),
        (1050, '03', '看订单贡献', '两类订单与整批核对'),
    ]
    for x, number, title, line in cards:
        d.rounded_rectangle((x, 402, x + 450, 660), radius=26, fill='white')
        text(d, (x + 28, 424), number, 32, '#87A295')
        text(d, (x + 28, 490), title, 40, '#172F29')
        fit_check(d, line, 28, 394)
        text(d, (x + 28, 567), line, 28, '#567166')
    text(d, (84, 727), '¥19.90', 78, '#163D34')
    text(d, (435, 749), '一次购买本版本文件', 35, '#395B50')
    text(d, (84, 861), '含XLSX、PDF说明、虚构案例与60秒演示', 29, '#395B50')
    text(d, (84, 921), '贡献需另算固定成本和税 · 案例全部虚构 · 尚未真人验证', 25, '#64776D')
    image.save(IMAGES / '01-封面.png')

def flow():
    image = Image.new('RGB', (1600, 1400), '#F4F6F2')
    d = ImageDraw.Draw(image)
    text(d, (82, 74), '先整理，再比较', 68, '#172F29')
    text(d, (84, 180), '需要同款商品、售后已结束的一批订单与费用依据', 34, '#567166')
    steps = [
        ('1  分好两类订单', '正常完成 + 退货；取消、仅退款、换货另算。'),
        ('2  归集收入和费用', '录入账单总额；未知先查账，确定没有才填0。'),
        ('3  只粘贴数值到场景', '归集改变后重新粘贴；方案不会自动刷新。'),
        ('4  比较输入假设下的贡献', '改单价时核对费用；实际销量与退货率仍需验证。'),
    ]
    for i, (title, line) in enumerate(steps):
        y = 282 + i * 175
        d.rounded_rectangle((82, y, 1518, y + 151), radius=24, fill='white')
        text(d, (116, y + 23), title, 39, '#172F29')
        fit_check(d, line, 30, 1365)
        text(d, (116, y + 89), line, 30, '#567166')
    d.rounded_rectangle((82, 1030, 1518, 1336), radius=24, fill='#163D34')
    text(d, (118, 1060), '购买前，先确认使用条件', 42, '#FFFFFF')
    lines = [
        '电脑XLSX；已做macOS WPS部分实测，其他版本尚未实测。',
        '输入金额和分类由你确认；不自动导入账单或归因广告。',
        '固定成本、税、自有劳动与现金流另算；不是店铺净利润。',
        'AI虚构示例用于说明口径，不代表真实经营收益。',
    ]
    for i, line in enumerate(lines):
        fit_check(d, line, 28, 1364)
        text(d, (118, 1130 + i * 48), line, 28, '#DCEAE1')
    image.save(IMAGES / '02-填写流程与边界.png')

def square_cover():
    image = Image.new('RGB', (1200, 1200), '#F4F6F2')
    d = ImageDraw.Draw(image)
    d.rounded_rectangle((72, 65, 590, 135), radius=20, fill='#163D34')
    text(d, (97, 83), '电脑表格 / v0.3 试用版', 35, '#FFFFFF')
    text(d, (72, 219), '一款商品', 90, '#172F29')
    text(d, (72, 346), '三套方案', 90, '#172F29')
    text(d, (77, 489), '成本 · 运费 · 广告 · 退货损耗', 39, '#395B50')
    d.rounded_rectangle((72, 602, 1128, 823), radius=26, fill='white')
    text(d, (103, 639), '先归集账单，再比较订单贡献', 43, '#172F29')
    text(d, (103, 725), '含XLSX、PDF、虚构案例与演示', 34, '#567166')
    text(d, (72, 874), '¥19.90', 86, '#163D34')
    text(d, (77, 1000), '一次购买本版本文件', 36, '#395B50')
    text(d, (77, 1094), '需另算固定成本和税 · AI虚构案例 · 尚未真人验证', 25, '#64776D')
    image.save(IMAGES / '04-商品正方形封面.png')

def build_package():
    BUYER.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ROOT / 'templates/单品利润与退货测算样品.xlsx', BUYER / '单品利润与退货测算样品.xlsx')
    shutil.copy2(ROOT / 'docs/两页使用说明.pdf', BUYER / '两页使用说明.pdf')
    start = (ROOT / 'docs/开始试用.md').read_text()
    start = start.replace('本地试用样品，尚未发布收费商品。', '电脑填写试用版。')
    start = start[:start.index('本项目包含')] + '交付包文件顺序见「先读我.md」，使用与售后范围见「使用与售后说明.md」。所有案例和示例均为AI虚构。\n'
    (BUYER / '开始试用.md').write_text(start)
    shutil.copy2(SALE / '一分钟图文演示.mp4', BUYER / '一分钟图文演示.mp4')
    case = BUYER / '虚构填写案例'
    case.mkdir(exist_ok=True)
    for filename in ['填写结果.xlsx', '独立虚构订单.csv', '虚构批次及假设.md']:
        shutil.copy2(ROOT / 'examples/merchant-v0.3' / filename, case / filename)
    zip_path = SALE / '单品成本与退货贡献测算_v0.3_交付包.zip'
    payload = sorted(p for p in BUYER.rglob('*') if p.is_file())
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in payload:
            z.write(p, str(Path('单品成本与退货贡献测算_v0.3') / p.relative_to(BUYER)))
    with zipfile.ZipFile(zip_path) as z:
        if z.testzip() is not None:
            raise RuntimeError('ZIP failed integrity check')
        if len(z.namelist()) != len(payload):
            raise RuntimeError('ZIP file count mismatch')
    manifest = {
        'version': 'v0.3', 'draft_price_yuan': '19.90',
        'published': False, 'platform_draft_saved': False,
        'delivery_link_configured': False, 'file_count': len(payload),
        'zip_sha256': hashlib.sha256(zip_path.read_bytes()).hexdigest(),
        'files': {str(p.relative_to(BUYER)): hashlib.sha256(p.read_bytes()).hexdigest() for p in payload},
    }
    state_path = SALE / '内部记录/平台状态.json'
    if state_path.exists():
        state = json.loads(state_path.read_text())
        for key in ['published', 'platform_draft_saved', 'delivery_link_configured']:
            manifest[key] = state[key]
        manifest['platform_state_checked_on'] = state['checked_on']
    (SALE / '内部记录/交付校验.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({'zip': str(zip_path), 'file_count': len(payload), 'zip_bytes': zip_path.stat().st_size}, ensure_ascii=False))

if __name__ == '__main__':
    IMAGES.mkdir(parents=True, exist_ok=True)
    cover()
    flow()
    square_cover()
    build_package()
