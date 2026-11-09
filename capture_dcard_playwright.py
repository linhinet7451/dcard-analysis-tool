"""用 Playwright 在自己的電腦上擷取 Dcard _api 回應。

為什麼需要它：Dcard 前方有 Cloudflare。已實測 plain requests、curl_cffi
與真實瀏覽器在被擋的 IP（資料中心／雲端）上都會 403，TLS 偽裝不是主因，
IP 信譽才是。所以最穩的自動化方式，是在自己的電腦、用自己的瀏覽器
profile 真的去逛 Dcard，讓瀏覽器自己發出 _api 請求，再把回應存下來，
之後用 dcard_desktop.py 的「匯入 HAR 檔案…」或直接讀這些 JSON 來分析。

2.0 版起，主程式已把 Playwright 直接整合進「自動（建議）」與
「Playwright 自動瀏覽」模式（見 dcard_desktop.py 的 playwright_collect_boards），
按「開始蒐集」就會自動用，不必另外跑本腳本；本腳本保留為獨立工具，方便
在終端機單獨擷取、除錯。

用法（在自己的電腦上，不要在雲端 VM 上跑，一定會被擋）：
    pip install playwright
    playwright install chromium
    python capture_dcard_playwright.py --forum food --posts 10
    python capture_dcard_playwright.py --forum relationship --posts 20 --headed

輸出：./har_captures/<timestamp>_<編號>.json，每個檔案是一次 _api 回應的 JSON。
Playwright 未安裝時，本檔案只會印出安裝提示就結束，不會印 traceback；
--help 在未安裝 Playwright 時也可以正常顯示。
"""
import argparse
import json
import sys
import time
from pathlib import Path

DEFAULT_OUT_DIR = 'har_captures'
DEFAULT_USER_DATA_DIR = 'dcard_profile'


def is_dcard_api_url(url):
    """是否為 Dcard API 回應（_api 或 service/api/v2）。純函式，可單獨測試。"""
    url = url or ''
    return 'dcard.tw' in url and ('/_api/' in url or '/service/api/v2/' in url)


def capture_filename(timestamp, index):
    """擷取檔名：<timestamp>_<3 位編號>.json。"""
    return f'{timestamp}_{index:03d}.json'


def build_parser():
    p = argparse.ArgumentParser(
        description='用 Playwright 擷取 Dcard _api 回應（在自己的電腦上執行）')
    p.add_argument('--forum', default='relationship',
                   help='看板 alias，預設 relationship（例如 food、travel）')
    p.add_argument('--posts', type=int, default=10,
                   help='目標瀏覽篇數（用於捲動次數），預設 10')
    p.add_argument('--out-dir', default=DEFAULT_OUT_DIR,
                   help='JSON 回應儲存目錄，預設 ./har_captures')
    p.add_argument('--user-data-dir', default=DEFAULT_USER_DATA_DIR,
                   help='Playwright 持續性瀏覽器 profile 目錄，預設 ./dcard_profile')
    p.add_argument('--headed', action='store_true',
                   help='顯示瀏覽器視窗執行（預設為無頭模式）')
    p.add_argument('--timeout', type=int, default=30,
                   help='頁面載入逾時秒數，預設 30')
    return p


def main(argv=None):
    # --help 必須在未安裝 Playwright 時也能顯示：先解析參數，再 import。
    args = build_parser().parse_args(argv)
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print('尚未安裝 Playwright，請先執行：')
        print('  pip install playwright')
        print('  playwright install chromium')
        print('然後再執行本腳本（請在自己的電腦上執行，不要在雲端 VM 上跑）。')
        return 2

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    saved = []

    def _on_response(response):
        if not is_dcard_api_url(response.url):
            return
        try:
            ctype = (response.headers.get('content-type') or '').lower()
            if 'json' not in ctype:
                return
            data = response.json()
        except Exception:
            return
        index = len(saved) + 1
        path = out_dir / capture_filename(timestamp, index)
        try:
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                            encoding='utf-8')
            saved.append(str(path))
        except OSError as exc:
            print(f'寫入失敗 {path}：{exc}', file=sys.stderr)

    forum_url = f'https://www.dcard.tw/f/{args.forum}'
    print(f'看板：{args.forum}（{forum_url}）')
    print(f'輸出目錄：{out_dir}，瀏覽器 profile：{args.user_data_dir}')
    print('提示：若 Dcard 要求登入，請在開啟的瀏覽器中登入後再重新執行。')
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            args.user_data_dir,
            headless=not args.headed,
            args=['--disable-blink-features=AutomationControlled'],
        )
        page = context.new_page()
        page.on('response', _on_response)
        try:
            page.goto(forum_url, timeout=args.timeout * 1000,
                      wait_until='domcontentloaded')
            page.wait_for_timeout(3000)
            # 捲動看板頁，觸發更多 _api/posts 請求
            steps = max(1, min(args.posts, 50))
            for _ in range(steps):
                page.mouse.wheel(0, 1200)
                page.wait_for_timeout(1500)
            # 點開前幾篇文章，觸發 /posts/{id} 與留言請求
            links = page.locator('a[href*="/p/"]').all()[: max(1, min(args.posts, 10))]
            for link in links:
                try:
                    link.click(timeout=5000)
                    page.wait_for_timeout(2500)
                    page.go_back(timeout=args.timeout * 1000)
                    page.wait_for_timeout(1500)
                except Exception:
                    continue
        except Exception as exc:
            print(f'瀏覽過程發生錯誤：{exc}', file=sys.stderr)
        finally:
            context.close()

    print(f'完成：擷取到 {len(saved)} 個 Dcard _api JSON 回應，存於 {out_dir}/')
    for path in saved:
        print(f'  {path}')
    if not saved:
        print('沒有擷取到任何 _api 回應：可能被 Cloudflare 阻擋，'
              '請改用 --headed 在自己的電腦上執行，或改用 Chrome 手動匯出 HAR。')
    else:
        print('下一步：在 dcard_desktop.py 按「匯入 HAR 檔案…」，'
              '或把這些 JSON 當作資料來源分析。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
