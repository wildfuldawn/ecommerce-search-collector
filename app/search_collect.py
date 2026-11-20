"""搜索列表采集入口：按目标数量翻页，绝不请求商品详情。"""
import argparse
import html
import json
import math
import os
from pathlib import Path
import random
import re
import time
from datetime import datetime
from io import BytesIO
from urllib.parse import urlparse


def collect(fetch, identify, target, max_pages=100, pause=lambda: None):
    """按唯一 ID 计数；中断和下一页失败仍返回已成功获取的部分。"""
    rows, seen, pages, reason = [], set(), 0, "达到页数上限"
    for page in range(1, max_pages + 1):
        try:
            items = fetch(page)
        except KeyboardInterrupt:
            reason = "用户中断，已保留此前数据"
            break
        except (Exception, SystemExit) as error:
            reason = f"第 {page} 页获取失败：{error}"
            break
        pages = page
        added = 0
        for item in items:
            key = str(identify(item) or "").strip()
            if not key or key in seen:
                continue
            seen.add(key)
            rows.append(item)
            added += 1
            if len(rows) >= target:
                break
        print(f"第 {page} 页新增 {added} 条，已获取 {len(rows)}/{target} 条")
        if len(rows) >= target:
            reason = "已达到目标数量"
            break
        if not items or not added:
            reason = "搜索结果为空或本页无新增商品"
            break
        estimate = math.ceil((target - len(rows)) / (len(rows) / pages))
        print(f"按当前去重结果估算还需约 {estimate} 页；实际按新增数量继续补足")
        if page < max_pages:
            try:
                pause()
            except KeyboardInterrupt:
                reason = "用户中断，已保留此前数据"
                break
    return rows, {"target": target, "actual": len(rows), "pages": pages,
                  "complete": len(rows) >= target, "reason": reason}


def text(value):
    if value is None:
        return ""
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False)
    return html.unescape(re.sub(r"<[^>]*>", "", str(value))).strip()


def image_url(value, jd=False):
    if isinstance(value, list):
        value = value[0] if value else ""
    if isinstance(value, dict):
        value = value.get("url") or value.get("src") or ""
    value = str(value or "").strip()
    if value.startswith("//"):
        return "https:" + value
    if jd and value and not value.startswith(("http://", "https://")):
        return "https://img14.360buyimg.com/n1/" + value.lstrip("/")
    return value if urlparse(value).scheme in ("http", "https") else ""


def normalize(platform, item):
    if platform == "taobao":
        shop = item.get("shopInfo") or {}
        price = item.get("priceShow") or {}
        key = item.get("item_id")
        values = [key, item.get("title"), price.get("price") or item.get("price"),
                  shop.get("title"), item.get("realSales"), item.get("procity"),
                  item.get("auctionURL"), image_url(item.get("pic_path") or item.get("picUrl") or item.get("pic"))]
    elif platform == "jingdong":
        key = item.get("skuId") or item.get("wareId")
        values = [key, item.get("wareName"), item.get("jdPrice") or item.get("price"),
                  item.get("shopName"), item.get("commentNum"), item.get("area"),
                  f"https://item.jd.com/{key}.html", image_url(item.get("imageurl") or item.get("imageUrl") or item.get("imgPath"), True)]
    else:
        key = item.get("itemId")
        values = [key, item.get("title"), item.get("price"), item.get("seller"),
                  item.get("wantCnt"), item.get("area"),
                  f"https://www.goofish.com/item?id={key}", image_url(item.get("picUrl"))]
    return dict(zip(["商品ID", "商品标题", "列表价格", "店铺或卖家", "销量/评论/想要", "地区", "商品链接", "图片链接"], map(text, values)))


def export(platform, keyword, products, status, output, download=True, session=None):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill, Alignment
    from PIL import Image
    import requests
    folder = Path(output)
    folder.mkdir(parents=True, exist_ok=True)
    stem = f"{platform}_search_{datetime.now():%Y%m%d_%H%M%S_%f}"
    images = folder / (stem + "_images")
    records = []
    wb = Workbook()
    ws = wb.active
    ws.title = "搜索商品"
    headers = ["商品ID", "商品标题", "列表价格", "店铺或卖家", "销量/评论/想要", "地区", "商品链接", "图片链接", "本地图片", "图片状态", "图片"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="204E76")
        cell.font = Font(color="FFFFFF", bold=True)
    client = session or requests.Session()
    for index, product in enumerate(products, 2):
        record = normalize(platform, product)
        record.update({"本地图片": "", "图片状态": "未提供图片"})
        local = None
        if record["图片链接"]:
            record["图片状态"] = "仅保留链接"
            if download:
                try:
                    response = client.get(record["图片链接"], timeout=20)
                    response.raise_for_status()
                    with Image.open(BytesIO(response.content)) as im:
                        im.thumbnail((800, 800))
                        images.mkdir(exist_ok=True)
                        local = images / f"{index-1:05d}.jpg"
                        im.convert("RGB").save(local, quality=88)
                    record["本地图片"] = str(local.relative_to(folder))
                    record["图片状态"] = "已下载并嵌入"
                except Exception as error:
                    record["图片状态"] = f"下载失败：{type(error).__name__}"
        ws.append([record.get(h, "") for h in headers])
        # 所有来源文本均作为文本保存，避免标题被解释为 Excel 公式。
        for cell in ws[index]:
            cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if local:
            picture = XLImage(str(local))
            scale = min(110 / picture.width, 110 / picture.height)
            picture.width, picture.height = round(picture.width * scale), round(picture.height * scale)
            ws.add_image(picture, f"K{index}")
            ws.row_dimensions[index].height = 86
        records.append(record)
    from openpyxl.utils import get_column_letter
    for col, width in enumerate([22, 52, 16, 25, 22, 18, 45, 45, 38, 28, 18], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(folder / (stem + ".xlsx"))
    (folder / (stem + ".json")).write_text(json.dumps(
        {"platform": platform, "keyword": keyword, **status, "products": records},
        ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已保存 {len(records)}/{status['target']} 条：{status['reason']}")
    print(f"Excel：{folder / (stem + '.xlsx')}")
    return folder / (stem + ".xlsx")


def xianyu_search_fetch(browser, keyword):
    """服务器会话过期时，在同一浏览器中手动登录并重试当前页一次。"""
    relogin_attempted = False

    def fetch(page):
        nonlocal relogin_attempted
        items, problem = browser.search(keyword, page)
        login_required = any(marker in (problem or "").upper() for marker in (
            "SESSION_EXPIRED", "SESSION_EMPTY", "NOT_LOGIN", "登录态已失效",
            "要求登录", "请先登录", "用户未登录",
        ))
        if not items and not login_required:
            login_required = browser.has_login_prompt()
        if login_required and not relogin_attempted:
            relogin_attempted = True
            print("闲鱼登录已失效，正在打开登录页面；请手动登录后回到终端按 Enter。")
            if not browser.ensure_login(force=True):
                raise RuntimeError("闲鱼登录未完成；已停止采集")
            items, problem = browser.search(keyword, page)
        if problem:
            raise RuntimeError(problem)
        return items

    return fetch


def run(platform, module):
    parser = argparse.ArgumentParser(description="仅采集搜索列表；按数量自动翻页")
    parser.add_argument("--keyword", default=module.SEARCH)
    parser.add_argument("--limit", type=int, default=int(os.environ.get("EC_LIMIT", "10")))
    parser.add_argument("--max-pages", type=int, default=100, help="自动翻页安全上限")
    parser.add_argument("--no-images", action="store_true")
    parser.add_argument("--relogin", action="store_true", help="打开当前平台浏览器重新登录")
    args = parser.parse_args()
    if args.limit < 1 or args.max_pages < 1 or not args.keyword.strip():
        parser.error("关键词不能为空，数量与页数上限必须大于 0")
    keyword = args.keyword.strip()
    module.SEARCH = keyword
    pause = lambda: time.sleep(random.uniform(3, 6))
    output = os.environ.get("EC_OUTPUT") or str(Path(__file__).resolve().parents[1] / "output")
    try:
        from browser_search import BrowserSearch
        with BrowserSearch(platform, module, keyword, force_login=args.relogin) as search:
            rows, status = collect(search.fetch, search.identify, args.limit, args.max_pages, pause)
        export(platform, keyword, rows, status, output, not args.no_images)
        return 0 if status["complete"] else 2
    except Exception as error:
        print(f"采集未能启动：{error}")
        return 1
