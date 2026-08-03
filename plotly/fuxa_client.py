from __future__ import annotations

import json
from datetime import datetime
from typing import Any

import requests

FUXA_URL = "http://127.0.0.1:1881"
TIMEOUT_SECONDS = 10

def get_tags(daq_only: bool = False) -> list[dict]:
    response = requests.get(f"{FUXA_URL}/api/project", timeout=TIMEOUT_SECONDS)
    response.raise_for_status()
    project = response.json()
    tags = []

    for device in project["devices"].values():
        for tag in device.get("tags", {}).values():
            daq = tag.get("daq", {})
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


def get_current_values(tag_ids: list[str]) -> list[dict]:
    response = requests.get(
        f"{FUXA_URL}/api/getTagValue",
        params={"ids": json.dumps(tag_ids, separators=(",", ":"))},
        timeout=TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    return response.json()


def set_tag_values(tag_values: dict[str, Any]) -> Any:
    response = requests.post(
        f"{FUXA_URL}/api/setTagValue",
        json={"tags": [{"id": tag_id, "value": value} for tag_id, value in tag_values.items()]},
        timeout=TIMEOUT_SECONDS,
    )
    response.raise_for_status()
    return response.json()


def get_daq_history(
    tag_ids: list[str],
    start: datetime,
    end: datetime,
) -> dict[str, list[dict]]:
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
    return dict(zip(tag_ids, histories))


if __name__ == "__main__":
    # 単体確認設定
    TEST_TAG_ID = "t_2b7faf71-ae894970"
    TEST_WRITE_VALUE = "1"

    write_result = set_tag_values({TEST_TAG_ID: TEST_WRITE_VALUE})
    read_result = get_current_values([TEST_TAG_ID])
    print("書込結果:", json.dumps(write_result, ensure_ascii=False, default=str))
    print("読取結果:", json.dumps(read_result, ensure_ascii=False, default=str))