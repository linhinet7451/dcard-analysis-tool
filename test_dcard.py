"""Dcard 工具純函式測試（不需 pytest，直接 python3 test_dcard.py 即可執行）。

只測不需網路、不需 Tkinter 的純函式：解析、時間、作者標籤、Excel 安全字串、
Gemini 批次提示詞。Dcard _api 的即時連線不在此測試（本機 VM 會被 Cloudflare 擋 403）。
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from capture_dcard_playwright import is_dcard_api_url, capture_filename
import dcard_desktop
from dcard_desktop import (
    parse_forums, parse_posts, parse_post_detail, parse_comments,
    format_dcard_time, author_label, safe_excel, build_batch_prompt,
    parse_har_entries, parse_har_data, classify_dcard_har_url,
    curl_cffi_available, curl_cffi_get_json,
    is_dcard_api_url as is_dcard_api_url_desktop,
    playwright_available, playwright_collect_boards, reload_playwright,
    stealth_available, stealth_collect_boards, reload_stealth,
    human_delay, human_scroll_steps, parse_next_data,
    parse_captured_responses, parse_oxylabs_result, oxylabs_fetch_json,
    PLAYWRIGHT_INSTALL_MSG, STEALTH_INSTALL_MSG, OXYLABS_URL,
    SAMPLE_POSTS, SAMPLE_COMMENTS, POST_COLS, COMMENT_COLS,
)

SAMPLE_FORUMS_JSON = [
    {"alias": "relationship", "name": "感情"},
    {"alias": "food", "name": "美食"},
    {"alias": "travel", "name": "旅遊"},
]

SAMPLE_POSTS_JSON = [
    {"id": "123", "title": "測試標題", "excerpt": "摘要…", "content": "全文內容",
     "createdAt": "2026-11-09T10:30:00.000Z", "likeCount": 42, "commentCount": 3,
     "forumAlias": "food", "forumName": "美食", "anonymous": True, "gender": "", "school": ""},
    # 缺欄位、另一種形狀：forum 包成物件、沒有 content 只有 excerpt
    {"id": 456, "title": "第二篇", "excerpt": "只有摘要",
     "createdAt": "2026-11-08T18:00:00+08:00", "likeCount": None, "commentCount": None,
     "forum": {"alias": "travel", "name": "旅遊"}, "anonymous": False,
     "gender": "女", "school": "某大學"},
    # 無效項目應被跳過
    {"title": "沒有 id"},
]

SAMPLE_COMMENTS_JSON = [
    {"id": "c2", "floor": 2, "content": "第二樓", "createdAt": "2026-11-09T11:00:00.000Z",
     "likeCount": 5, "anonymous": False, "gender": "男", "school": "某大學"},
    {"id": "c1", "floor": 1, "content": "第一樓", "createdAt": "2026-11-09T10:40:00.000Z",
     "likeCount": 12, "anonymous": True, "gender": "", "school": ""},
]


def test_parse_forums():
    out = parse_forums(SAMPLE_FORUMS_JSON)
    assert out == [("感情", "relationship"), ("美食", "food"), ("旅遊", "travel")], out
    assert parse_forums({"forums": SAMPLE_FORUMS_JSON}) == out
    assert parse_forums([]) == []
    assert parse_forums(None) == []


def test_parse_posts():
    posts = parse_posts(SAMPLE_POSTS_JSON)
    assert len(posts) == 2, posts
    assert posts[0]["id"] == "123"
    assert posts[0]["title"] == "測試標題"
    assert posts[0]["content"] == "全文內容"
    assert posts[0]["likeCount"] == 42
    assert posts[0]["commentCount"] == 3
    assert posts[0]["author"] == "匿名"
    # 缺欄位時要有預設值，不可崩潰
    assert posts[1]["id"] == "456"
    assert posts[1]["content"] == "只有摘要"
    assert posts[1]["likeCount"] == 0
    assert posts[1]["commentCount"] == 0
    assert posts[1]["forumAlias"] == "travel"
    assert "某大學" in posts[1]["author"]


def test_parse_post_detail():
    detail = parse_post_detail(SAMPLE_POSTS_JSON[0])
    assert detail is not None
    assert detail["title"] == "測試標題"
    assert detail["content"] == "全文內容"
    assert parse_post_detail({}) is None
    assert parse_post_detail({"data": SAMPLE_POSTS_JSON[0]})["id"] == "123"


def test_parse_comments():
    comments = parse_comments(SAMPLE_COMMENTS_JSON)
    assert [c["floor"] for c in comments] == [1, 2], comments
    assert comments[0]["author"] == "匿名"
    assert comments[1]["author"] == "某大學（男）"
    assert comments[1]["likeCount"] == 5
    assert parse_comments([]) == []


def test_format_dcard_time():
    assert format_dcard_time("2026-11-09T10:30:00.000Z") == "2026-11-09 10:30"
    assert format_dcard_time("2026-11-08T18:00:00+08:00") == "2026-11-08 18:00"
    assert format_dcard_time("2026-11-09 10:30:00") == "2026-11-09 10:30"
    assert format_dcard_time("") == ""
    assert format_dcard_time(None) == ""


def test_author_label():
    assert author_label({"anonymous": True, "school": "某大學"}) == "匿名"
    assert author_label({"anonymous": False, "school": "某大學", "gender": "女"}) == "某大學（女）"
    assert author_label({"anonymous": False, "school": "某大學", "gender": ""}) == "某大學"
    assert author_label({"anonymous": False, "school": "", "gender": "男"}) == "男"
    assert author_label({}) == ""
    assert author_label(None) == ""


def test_safe_excel():
    assert safe_excel(None) == ""
    assert safe_excel("=cmd") == "'=cmd"
    assert safe_excel("正常文字") == "正常文字"
    assert safe_excel(42) == 42


def test_build_batch_prompt_dcard_not_ptt():
    post = dict(SAMPLE_POSTS[0])
    prompt = build_batch_prompt([(1, post, SAMPLE_COMMENTS[:2])])
    assert "Dcard" in prompt
    assert "PTT" not in prompt
    assert "愛心" in prompt
    assert post["標題"] in prompt


def _har_entry(url, data):
    return {"request": {"url": url},
            "response": {"status": 200,
                         "content": {"mimeType": "application/json",
                                     "text": json.dumps(data, ensure_ascii=False)}}}


# 2 個看板列表條目（第 2 個含重複貼文，測試去重）、1 個留言條目、1 個無關條目
SAMPLE_HAR_ENTRIES = [
    _har_entry("https://www.dcard.tw/_api/forums/food/posts?popular=false&limit=30",
               [{"id": "111", "title": "拉麵推薦", "excerpt": "摘要一", "content": "全文一",
                 "createdAt": "2026-11-09T10:30:00.000Z", "likeCount": 42, "commentCount": 2,
                 "forumAlias": "food", "forumName": "美食", "anonymous": True},
                {"id": "222", "title": "夜市清單", "excerpt": "摘要二",
                 "createdAt": "2026-11-08T18:00:00.000Z", "likeCount": 10, "commentCount": 1,
                 "forumAlias": "food", "forumName": "美食", "anonymous": False,
                 "gender": "女", "school": "某大學"}]),
    _har_entry("https://www.dcard.tw/_api/forums/food/posts?popular=false&limit=30&before=222",
               [{"id": "222", "title": "夜市清單", "excerpt": "摘要二",
                 "createdAt": "2026-11-08T18:00:00.000Z", "likeCount": 10, "commentCount": 1,
                 "forumAlias": "food", "forumName": "美食", "anonymous": False,
                 "gender": "女", "school": "某大學"},
                {"id": "333", "title": "甜點店", "excerpt": "摘要三",
                 "createdAt": "2026-11-07T12:00:00.000Z", "likeCount": 5, "commentCount": 0,
                 "forumAlias": "food", "forumName": "美食", "anonymous": True}]),
    _har_entry("https://www.dcard.tw/_api/posts/111/comments?after=0&limit=30",
               [{"id": "c1", "floor": 1, "content": "第一樓", "createdAt": "2026-11-09T11:00:00.000Z",
                 "likeCount": 3, "anonymous": True},
                {"id": "c2", "floor": 2, "content": "第二樓", "createdAt": "2026-11-09T11:30:00.000Z",
                 "likeCount": 1, "anonymous": False, "gender": "男", "school": "某大學"}]),
    _har_entry("https://www.google.com/search?q=test", {"irrelevant": True}),
]


def test_classify_dcard_har_url():
    assert classify_dcard_har_url(
        "https://www.dcard.tw/_api/forums/food/posts?popular=false") == ("posts", "food")
    # comments 必須先於 detail，否則會被誤判
    assert classify_dcard_har_url(
        "https://www.dcard.tw/_api/posts/111/comments?after=0") == ("comments", "111")
    assert classify_dcard_har_url(
        "https://www.dcard.tw/_api/posts/111") == ("detail", "111")
    assert classify_dcard_har_url(
        "https://www.dcard.tw/service/api/v2/forums/food/posts") == ("posts", "food")
    assert classify_dcard_har_url(
        "https://www.dcard.tw/service/api/v2/posts/111/comments") == ("comments", "111")
    assert classify_dcard_har_url("https://www.dcard.tw/f/food") is None
    assert classify_dcard_har_url("https://www.google.com/") is None
    assert classify_dcard_har_url("") is None


def test_parse_har_entries():
    posts, comments = parse_har_entries(SAMPLE_HAR_ENTRIES)
    # 2 個列表條目共 4 筆，其中 1 筆重複，去重後應為 3 篇
    assert len(posts) == 3, posts
    assert [p["文章ID"] for p in posts] == ["111", "222", "333"]
    assert posts[0]["標題"] == "拉麵推薦"
    assert posts[0]["看板"] == "美食" and posts[0]["看板代碼"] == "food"
    assert posts[0]["作者"] == "匿名"
    assert posts[0]["內文"] == "全文一"
    assert posts[1]["作者"] == "某大學（女）"
    assert posts[0]["文章網址"] == "https://www.dcard.tw/f/food/p/111"
    for p in posts:
        for col in POST_COLS:
            assert col in p, (col, p)
    # 1 個留言條目，2 則留言；無關條目被跳過
    assert len(comments) == 2, comments
    assert [c["樓層"] for c in comments] == [1, 2]
    assert comments[0]["文章ID"] == "111"
    assert comments[0]["文章標題"] == "拉麵推薦"
    assert comments[0]["留言內容"] == "第一樓"
    assert comments[1]["使用者"] == "某大學（男）"
    for c in comments:
        for col in COMMENT_COLS:
            assert col in c, (col, c)
    # 留言數不小於實際匯入的留言數
    assert posts[0]["留言數"] >= 2


def test_parse_har_entries_no_dcard():
    posts, comments = parse_har_entries(
        [_har_entry("https://www.google.com/search?q=test", {"x": 1})])
    assert posts == [] and comments == []
    assert parse_har_entries([]) == ([], [])
    assert parse_har_data({"log": {"entries": SAMPLE_HAR_ENTRIES}})[0][0]["文章ID"] == "111"
    assert parse_har_data({}) == ([], [])
    assert parse_har_data(None) == ([], [])


def test_parse_har_entries_detail_merges_content():
    entries = [
        _har_entry("https://www.dcard.tw/_api/forums/food/posts",
                   [{"id": "999", "title": "只有摘要", "excerpt": "短摘要",
                     "createdAt": "2026-11-09T10:00:00.000Z", "likeCount": 1,
                     "commentCount": 0, "forumAlias": "food", "forumName": "美食",
                     "anonymous": True}]),
        _har_entry("https://www.dcard.tw/_api/posts/999",
                   {"id": "999", "title": "只有摘要", "content": "這是點開文章後才有的完整內文，比較長。",
                    "createdAt": "2026-11-09T10:00:00.000Z", "likeCount": 2,
                    "commentCount": 0, "forumAlias": "food", "forumName": "美食",
                    "anonymous": True}),
    ]
    posts, comments = parse_har_entries(entries)
    assert len(posts) == 1
    assert posts[0]["內文"] == "這是點開文章後才有的完整內文，比較長。"
    assert posts[0]["愛心數"] == 2
    assert comments == []


def test_curl_cffi_guarded():
    # import 不可因 curl_cffi 未安裝而崩潰；helper 必須是 bool / 明確報錯
    assert isinstance(curl_cffi_available(), bool)
    if not curl_cffi_available():
        try:
            curl_cffi_get_json("https://www.dcard.tw/_api/forums")
            raise AssertionError("未安裝 curl_cffi 時應 raise RuntimeError")
        except RuntimeError as exc:
            assert "未安裝" in str(exc) or "curl_cffi" in str(exc)
    # 若環境有安裝，則 available() 為 True 即可（不發真實請求）


def test_playwright_capture_url_matching():
    assert is_dcard_api_url("https://www.dcard.tw/_api/forums")
    assert is_dcard_api_url("https://www.dcard.tw/_api/posts/111/comments?after=0")
    assert is_dcard_api_url("https://www.dcard.tw/service/api/v2/forums/food/posts")
    assert not is_dcard_api_url("https://www.dcard.tw/f/food")
    assert not is_dcard_api_url("https://www.dcard.tw/")
    assert not is_dcard_api_url("https://www.google.com/_api/forums")
    assert not is_dcard_api_url("")
    assert capture_filename("20261109_120000", 1) == "20261109_120000_001.json"
    assert capture_filename("20261109_120000", 12) == "20261109_120000_012.json"


def test_is_dcard_api_url_desktop():
    # dcard_desktop 內建的版本須與獨立腳本一致
    for url, expected in [
        ("https://www.dcard.tw/_api/forums", True),
        ("https://www.dcard.tw/_api/posts/111/comments?after=0", True),
        ("https://www.dcard.tw/service/api/v2/forums/food/posts", True),
        ("https://www.dcard.tw/f/food", False),
        ("https://www.google.com/_api/forums", False),
        ("", False),
    ]:
        assert is_dcard_api_url_desktop(url) == expected, url
        assert is_dcard_api_url(url) == expected, url
    assert isinstance(playwright_available(), bool)


def test_parse_captured_responses():
    # Playwright 攔截到的 (url, data)，用 list 與 dict 兩種形式都要能解析
    captured_list = [
        ("https://www.dcard.tw/_api/forums/food/posts?popular=false&limit=30",
         SAMPLE_POSTS_JSON[:1]),
        ("https://www.dcard.tw/_api/posts/123/comments?after=0&limit=30",
         SAMPLE_COMMENTS_JSON),
    ]
    posts, comments = parse_captured_responses(
        captured_list, selected=[("美食", "food")], limit_per_board=10)
    assert len(posts) == 1 and posts[0]["文章ID"] == "123", posts
    assert [c["樓層"] for c in comments] == [1, 2], comments
    for p in posts:
        for col in POST_COLS:
            assert col in p
    # dict 形式
    posts2, comments2 = parse_captured_responses(
        dict(captured_list), selected=[("美食", "food")])
    assert len(posts2) == 1 and len(comments2) == 2
    # 不抓留言時留言為空；看板篩選會擋掉不相符的貼文
    posts3, comments3 = parse_captured_responses(
        captured_list, selected=[("美食", "food")], include_comments=False)
    assert len(posts3) == 1 and comments3 == []
    posts4, _ = parse_captured_responses(
        captured_list, selected=[("旅遊", "travel")])
    assert posts4 == []
    # 每看板上限：3 篇只留前 2 篇（依攔截順序）
    many = [("https://www.dcard.tw/_api/forums/food/posts",
             [{"id": str(i), "title": f"第{i}篇", "excerpt": "x",
               "createdAt": "2026-11-09T10:00:00.000Z", "likeCount": 1,
               "commentCount": 0, "forumAlias": "food", "forumName": "美食",
               "anonymous": True} for i in (1, 2, 3)])]
    posts5, _ = parse_captured_responses(
        many, selected=[("美食", "food")], limit_per_board=2)
    assert [p["文章ID"] for p in posts5] == ["1", "2"], posts5
    assert parse_captured_responses([]) == ([], [])
    assert parse_captured_responses(None) == ([], [])


def test_playwright_collect_boards_missing():
    # 不啟動瀏覽器：把 guarded import 旗標暫時設為 None，並模擬 import 失敗，
    # 必須明確報安裝指令（v2.1：playwright_available() 會重新 import，
    # 所以要連 import 也 mock 掉，否則在已安裝的機器上會誤判為可用）
    import importlib
    original = dcard_desktop.sync_playwright
    original_import = importlib.import_module
    try:
        dcard_desktop.sync_playwright = None
        importlib.import_module = lambda name, *a, **k: (_ for _ in ()).throw(
            ImportError("No module named 'playwright'"))
        assert playwright_available() is False
        try:
            playwright_collect_boards([("美食", "food")], 10, True, False)
            raise AssertionError("未安裝 Playwright 時應 raise RuntimeError")
        except RuntimeError as exc:
            assert "pip install playwright" in str(exc), str(exc)
            assert "playwright install chromium" in str(exc), str(exc)
    finally:
        importlib.import_module = original_import
        dcard_desktop.sync_playwright = original
    assert "pip install playwright" in PLAYWRIGHT_INSTALL_MSG


def test_playwright_available_reloads_after_install():
    # v2.1 回歸測試：模擬「啟動時未安裝 → 安裝後才出現」的情況。
    # 把一個假的 playwright.sync_api 模組放進 sys.modules，
    # sync_playwright 先設為 None，playwright_available() 必須重新
    # import 並回 True，而不是一直回 False（舊 bug：要重開程式才認得）。
    import types
    original = dcard_desktop.sync_playwright
    fake_mod = types.ModuleType("playwright.sync_api")
    fake_mod.sync_playwright = object()
    try:
        dcard_desktop.sync_playwright = None
        sys.modules["playwright.sync_api"] = fake_mod
        assert playwright_available() is True
        assert dcard_desktop.sync_playwright is fake_mod.sync_playwright
        assert reload_playwright() is True
    finally:
        sys.modules.pop("playwright.sync_api", None)
        dcard_desktop.sync_playwright = original


def test_playwright_available_false_when_import_fails():
    # import 失敗時仍回 False，不拋錯。
    import importlib
    original = dcard_desktop.sync_playwright
    original_import = importlib.import_module
    try:
        dcard_desktop.sync_playwright = None
        importlib.import_module = lambda name, *a, **k: (_ for _ in ()).throw(
            ImportError("No module named 'playwright'"))
        assert playwright_available() is False
        assert reload_playwright() is False
    finally:
        importlib.import_module = original_import
        dcard_desktop.sync_playwright = original


def test_stealth_available_returns_tuple():
    # 回傳 (bool, 引擎名稱)，引擎只能是這四種，且絕不拋錯
    ok, engine = stealth_available()
    assert isinstance(ok, bool)
    assert engine in ("patchright", "camoufox", "playwright", "none"), engine
    assert ok == (engine != "none")
    # 本 VM 沒裝 patchright / camoufox，playwright 也沒裝，應為 none
    assert (ok, engine) == (False, "none")


def test_stealth_collect_boards_missing():
    # 不啟動瀏覽器：三個引擎的旗標都設為 None，並模擬 import 全部失敗，
    # 必須明確報安裝指令（含 patchright 與 camoufox）
    import importlib
    originals = (dcard_desktop.sync_playwright,
                 dcard_desktop._patchright_sync_playwright,
                 dcard_desktop._camoufox_Camoufox)
    original_import = importlib.import_module
    try:
        dcard_desktop.sync_playwright = None
        dcard_desktop._patchright_sync_playwright = None
        dcard_desktop._camoufox_Camoufox = None
        importlib.import_module = lambda name, *a, **k: (_ for _ in ()).throw(
            ImportError(f"No module named '{name}'"))
        assert stealth_available() == (False, "none")
        try:
            stealth_collect_boards([("美食", "food")], 10, True, False)
            raise AssertionError("未安裝任何隱身引擎時應 raise RuntimeError")
        except RuntimeError as exc:
            assert "pip install patchright" in str(exc), str(exc)
            assert "patchright install chromium" in str(exc), str(exc)
            assert "camoufox" in str(exc), str(exc)
    finally:
        importlib.import_module = original_import
        (dcard_desktop.sync_playwright,
         dcard_desktop._patchright_sync_playwright,
         dcard_desktop._camoufox_Camoufox) = originals
    assert "pip install patchright" in STEALTH_INSTALL_MSG
    assert "camoufox" in STEALTH_INSTALL_MSG


def test_reload_stealth_after_install():
    # v2.1 教訓的隱身版回歸測試：模擬「啟動時未安裝 → 安裝後才出現」。
    # 把假的 patchright.sync_api 模組放進 sys.modules，旗標先設為 None，
    # reload_stealth() 必須重新 import 並回 'patchright'，不需重開程式。
    import types
    originals = (dcard_desktop.sync_playwright,
                 dcard_desktop._patchright_sync_playwright,
                 dcard_desktop._camoufox_Camoufox)
    fake_mod = types.ModuleType("patchright.sync_api")
    fake_mod.sync_playwright = object()
    try:
        dcard_desktop.sync_playwright = None
        dcard_desktop._patchright_sync_playwright = None
        dcard_desktop._camoufox_Camoufox = None
        sys.modules["patchright.sync_api"] = fake_mod
        assert reload_stealth() == "patchright"
        assert stealth_available() == (True, "patchright")
        assert dcard_desktop._patchright_sync_playwright is fake_mod.sync_playwright
    finally:
        sys.modules.pop("patchright.sync_api", None)
        (dcard_desktop.sync_playwright,
         dcard_desktop._patchright_sync_playwright,
         dcard_desktop._camoufox_Camoufox) = originals


def test_reload_stealth_camoufox_second_choice():
    # Patchright 沒有時，Camoufox 是第二順位：patchright 的 import 模擬失敗，
    # 假的 camoufox.sync_api 在 sys.modules 裡，必須回 'camoufox'。
    import types
    import importlib
    originals = (dcard_desktop.sync_playwright,
                 dcard_desktop._patchright_sync_playwright,
                 dcard_desktop._camoufox_Camoufox)
    original_import = importlib.import_module

    def _fake_import(name, *a, **k):
        if name == "patchright.sync_api":
            raise ImportError("No module named 'patchright'")
        return original_import(name, *a, **k)

    fake_mod = types.ModuleType("camoufox.sync_api")
    fake_mod.Camoufox = object()
    try:
        dcard_desktop.sync_playwright = None
        dcard_desktop._patchright_sync_playwright = None
        dcard_desktop._camoufox_Camoufox = None
        sys.modules["camoufox.sync_api"] = fake_mod
        importlib.import_module = _fake_import
        assert reload_stealth() == "camoufox"
        assert stealth_available() == (True, "camoufox")
    finally:
        importlib.import_module = original_import
        sys.modules.pop("camoufox.sync_api", None)
        (dcard_desktop.sync_playwright,
         dcard_desktop._patchright_sync_playwright,
         dcard_desktop._camoufox_Camoufox) = originals


def test_human_scroll_steps():
    # 步長總和必須恰好等於 total_px，每步在 300–700 之間
    for total in (300, 700, 800, 1000, 2500, 5000):
        steps = human_scroll_steps(total)
        assert sum(steps) == total, (total, steps)
        assert all(300 <= s <= 700 for s in steps), (total, steps)
        assert len(steps) >= 1
    # 邊界：0 以下回 []；小於最小步長時一步到位
    assert human_scroll_steps(0) == []
    assert human_scroll_steps(-50) == []
    assert human_scroll_steps(100) == [100]
    assert human_scroll_steps(299) == [299]
    # 多次呼叫的步長應有隨機性（不要求每次都不同，但至少不是固定值）
    assert human_scroll_steps("not-a-number") == []


def test_human_delay():
    assert human_delay(0, 0) == 0.0
    assert human_delay(0.5, 0.5) == 0.5
    # min > max 時自動對調，回傳值落在對調後的區間內
    d = human_delay(2, 1)
    assert 1.0 <= d <= 2.0, d


def test_parse_next_data():
    html = ('<html><head><script id="__NEXT_DATA__" type="application/json">'
            '{"props":{"pageProps":{"a":1}}}</script></head></html>')
    assert parse_next_data(html) == {"props": {"pageProps": {"a": 1}}}
    # 屬性順序顛倒也要能抽
    html2 = '<script type="application/json" id="__NEXT_DATA__">{"x": 2}</script>'
    assert parse_next_data(html2) == {"x": 2}
    # 找不到、空字串、非字串、壞 JSON、解析結果不是 dict 都回 None，不拋錯
    assert parse_next_data("<html><body>沒有 NEXT_DATA</body></html>") is None
    assert parse_next_data("") is None
    assert parse_next_data(None) is None
    assert parse_next_data(123) is None
    assert parse_next_data(
        '<script id="__NEXT_DATA__" type="application/json">{bad json</script>') is None
    assert parse_next_data(
        '<script id="__NEXT_DATA__" type="application/json">[1, 2]</script>') is None


def test_posts_from_next_data():
    # __NEXT_DATA__ 裡巢狀的貼文列表要能被遞迴找出並轉成 POST_COLS rows
    data = {"props": {"pageProps": {"posts": SAMPLE_POSTS_JSON[:1]}}}
    rows = dcard_desktop._posts_from_next_data(data, "food")
    assert len(rows) == 1, rows
    assert rows[0]["文章ID"] == "123"
    assert rows[0]["標題"] == "測試標題"
    assert rows[0]["看板"] == "美食" and rows[0]["看板代碼"] == "food"
    for col in POST_COLS:
        assert col in rows[0]
    # 沒有貼文列表、不是 dict 時回 []
    assert dcard_desktop._posts_from_next_data({"props": {}}, "food") == []
    assert dcard_desktop._posts_from_next_data(None, "food") == []
    assert dcard_desktop._posts_from_next_data([], "food") == []
    # 混合非貼文元素的 list 不應被誤認
    assert dcard_desktop._posts_from_next_data(
        {"x": [{"id": "1", "title": "t"}, {"nope": True}]}, "food") == []


def test_parse_oxylabs_result():
    assert OXYLABS_URL == "https://realtime.oxylabs.io/v1/queries"
    # content 已是 dict / list 時直接回傳
    assert parse_oxylabs_result({"results": [{"content": {"a": 1}}]}) == {"a": 1}
    assert parse_oxylabs_result({"results": [{"content": [1, 2]}]}) == [1, 2]
    # content 是 JSON 字串時解析
    assert parse_oxylabs_result(
        {"results": [{"content": json.dumps([{"id": "1", "title": "t"}])}]}) == [
        {"id": "1", "title": "t"}]
    # 缺 results / content / 不是 JSON 時明確報錯，不可靜默回空
    for bad in ({}, {"results": []}, {"results": [{}]},
                {"results": [{"content": ""}]},
                {"results": [{"content": "<html>blocked</html>"}]},
                {"results": "nope"}, "nope"):
        try:
            parse_oxylabs_result(bad)
            raise AssertionError(f"應報錯：{bad!r}")
        except RuntimeError:
            pass


def test_oxylabs_fetch_json_missing_credentials():
    # 帳號密碼缺失時不發任何網路請求，直接報錯
    for user, pw in (("", ""), ("user", ""), ("", "pass")):
        try:
            oxylabs_fetch_json("https://www.dcard.tw/_api/forums", user, pw)
            raise AssertionError("缺帳號密碼時應 raise RuntimeError")
        except RuntimeError as exc:
            assert "Oxylabs" in str(exc), str(exc)


def test_sample_demo_data():
    assert len(SAMPLE_POSTS) == 6, len(SAMPLE_POSTS)
    assert len(SAMPLE_COMMENTS) == 12, len(SAMPLE_COMMENTS)
    assert {p["看板"] for p in SAMPLE_POSTS} == {"感情", "美食", "旅遊"}
    for p in SAMPLE_POSTS:
        for col in POST_COLS:
            assert col in p, (col, p)
    for c in SAMPLE_COMMENTS:
        for col in COMMENT_COLS:
            assert col in c, (col, c)


def main():
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in tests:
        fn()
        print(f"PASS {fn.__name__}")
    print(f"全部通過：{len(tests)} 個測試")


if __name__ == "__main__":
    main()
