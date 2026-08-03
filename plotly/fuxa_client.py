from __future__ import annotations

import json
from datetime import datetime
from typing import Any

import requests

# ============================================================
# FUXA接続設定
# ============================================================

# FUXAのHTTP API
FUXA_URL = "http://127.0.0.1:1881"

# HTTPリクエストのタイムアウト秒数
TIMEOUT_SECONDS = 10


# ============================================================
# タグ一覧取得
#
# FUXAのプロジェクト設定から、登録されているタグを取得する。
#
# daq_only=False:
#   すべてのタグを取得する。
#
# daq_only=True:
#   DAQ保存が有効なタグだけを取得する。
# ============================================================

def get_tags(daq_only: bool = False) -> list[dict]:
    response = requests.get(f"{FUXA_URL}/api/project", timeout=TIMEOUT_SECONDS)
    response.raise_for_status()
    project = response.json()
    tags = []

    # デバイスごとに登録されているタグを取得する
    for device in project["devices"].values():
        for tag in device.get("tags", {}).values():
            daq = tag.get("daq", {})

            # DAQ対象のみ取得する場合は無効タグを除外する
            if daq_only and not daq.get("enabled", False):
                continue

            tags.append({
                "id": tag["id"],
                "name": tag["name"],
                "device": device["name"],
                "type": tag["type"],
                "value": tag.get("value"),
                "timestamp": tag.get("timestamp"),
                "daq": daq,
            })

    return tags

# ============================================================
# タグ現在値取得
#
# 指定したタグIDの現在値をFUXAから取得する。
#
# 例:
# get_current_values(["t_2b7faf71-ae894970"])
# ============================================================

def get_current_values(tag_ids: list[str]) -> list[dict]:
    response = requests.get(
        f"{FUXA_URL}/api/getTagValue",
        params={"ids": json.dumps(tag_ids, separators=(",", ":"))},
        timeout=TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    return response.json()

# ============================================================
# タグ値書き込み
#
# 指定したタグIDへ値を書き込む。
#
# 例:
# set_tag_values({"t_2b7faf71-ae894970": 1})
# ============================================================

def set_tag_values(tag_values: dict[str, Any]) -> Any:
    response = requests.post(
        f"{FUXA_URL}/api/setTagValue",
        json={"tags": [{"id": tag_id, "value": value} for tag_id, value in tag_values.items()]},
        timeout=TIMEOUT_SECONDS,
    )
    response.raise_for_status()

    # レスポンス本文が空の場合はNoneを返す
    if not response.content:
        return None

    # JSONでないレスポンスの場合は文字列を返す
    try:
        return response.json()
    except requests.exceptions.JSONDecodeError:
        return response.text

# ============================================================
# DAQ履歴取得
#
# 指定したタグIDと時間範囲からDAQ履歴を取得する。
#
# 戻り値はタグIDをキーとした辞書にする。
#
# {
#     "タグID": [
#         {
#             "dt": Unixミリ秒,
#             "value": "値",
#         }
#     ]
# }
# ============================================================

def get_daq_history(tag_ids: list[str], start: datetime, end: datetime) -> dict[str, list[dict]]:
    if start >= end:
        raise ValueError("startはendより前にしてください。")

    query = {
        "sids": tag_ids,
        "from": int(start.timestamp() * 1000),
        "to": int(end.timestamp() * 1000),
    }
    response = requests.get(
        f"{FUXA_URL}/api/daq",
        params={"query": json.dumps(query, separators=(",", ":"))},
        timeout=TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    histories = response.json()

    # FUXAの戻り値はtag_idsと同じ順番の配列になっている
    return dict(zip(tag_ids, histories))

# ============================================================
# 単体確認
#
# TEST_WRITE_VALUEを手動で変更して実行する。
#
# 実行:
# python fuxa_client.py
#
# 確認内容:
# 1. 指定タグへ値を1回書き込む
# 2. 指定タグの現在値を1回読み取る
# ============================================================

if __name__ == "__main__":
    # ============================================================
    # 単体確認設定
    #
    # python fuxa_client.pyを実行すると、
    # 指定タグへ1回書き込み、その直後に1回読み取る。
    # ============================================================

    # 単体確認するタグID
    TEST_TAG_ID = "t_2b7faf71-ae894970"

    # 単体確認で書き込む値
    TEST_WRITE_VALUE = 1

    print(f"書込: tag={TEST_TAG_ID}, value={TEST_WRITE_VALUE}")
    write_result = set_tag_values({TEST_TAG_ID: TEST_WRITE_VALUE})
    print("書込結果:", json.dumps(write_result, ensure_ascii=False, default=str))
    read_result = get_current_values([TEST_TAG_ID])
    print("読取結果:", json.dumps(read_result, ensure_ascii=False, default=str))