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

def get_tags(
    daq_only: bool = False,
) -> list[dict]:
    project = _get(
        "/api/project"
    )

    tags = []

    # デバイスごとに登録されているタグを取得する
    for device in project["devices"].values():
        for tag in device.get("tags", {}).values():
            daq = tag.get("daq", {})

            # DAQ対象のみ取得する場合は無効タグを除外する
            if daq_only and not daq.get("enabled", False):
                continue

            tags.append(
                {
                    "id": tag["id"],
                    "name": tag["name"],
                    "device": device["name"],
                    "type": tag["type"],
                    "value": tag.get("value"),
                    "timestamp": tag.get("timestamp"),
                    "daq": daq,
                }
            )

    return tags


# ============================================================
# タグ現在値取得
#
# 指定したタグIDの現在値をFUXAから取得する。
#
# 例:
# get_current_values([
#     "t_2b7faf71-ae894970",
# ])
# ============================================================

def get_current_values(
    tag_ids: list[str],
) -> list[dict]:
    return _get(
        "/api/getTagValue",
        {
            "ids": json.dumps(
                tag_ids,
                separators=(",", ":"),
            ),
        },
    )


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

def get_daq_history(
    tag_ids: list[str],
    start: datetime,
    end: datetime,
) -> dict[str, list[dict]]:
    if start >= end:
        raise ValueError(
            "startはendより前にしてください。"
        )

    histories = _get(
        "/api/daq",
        {
            "query": json.dumps(
                {
                    "sids": tag_ids,
                    "from": _to_milliseconds(start),
                    "to": _to_milliseconds(end),
                },
                separators=(",", ":"),
            ),
        },
    )

    # FUXAの戻り値はtag_idsと同じ順番の配列になっている
    return dict(
        zip(
            tag_ids,
            histories,
        )
    )


# ============================================================
# FUXA GET API共通処理
#
# 指定されたAPIへGETリクエストを送信し、
# JSONレスポンスをPythonオブジェクトとして返す。
# ============================================================

def _get(
    path: str,
    params: dict | None = None,
) -> Any:
    response = requests.get(
        f"{FUXA_URL}{path}",
        params=params,
        timeout=TIMEOUT_SECONDS,
    )

    # HTTPエラーの場合は例外を発生させる
    response.raise_for_status()

    return response.json()


# ============================================================
# Unixミリ秒変換
#
# datetimeをFUXAのDAQ APIで使用するUnixミリ秒へ変換する。
# ============================================================

def _to_milliseconds(
    value: datetime,
) -> int:
    return int(
        value.timestamp() * 1000
    )