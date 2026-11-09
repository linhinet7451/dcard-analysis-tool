"""Dcard 工具純函式測試（不需 pytest，直接 python3 test_dcard.py 即可執行）。

只測不需網路、不需 Tkinter 的純函式：解析、時間、作者標籤、Excel 安全字串、
Gemini 批次提示詞。Dcard _api 的即時連線不在此測試（本機 VM 會被 Cloudflare 擋 403）。
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from capture_dcard_playwright import is_dcard_api_url, capture_filename
from dcard_desktop import (
    parse_forums, parse_posts, parse_post_detail, parse_comments,
    format_dcard_time, author_label, safe_excel, build_batch_prompt,
    parse_har_entries, parse_har_data, classify_dcard_har_url,
    curl_cffi_available, curl_cffi_get_json,
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
