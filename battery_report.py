# -*- coding: utf-8 -*-
"""
Windows Battery Report 深度分析工具
================================================

功能：
1. 读取 Windows powercfg /batteryreport 生成的 HTML。
2. 解析：
   - Installed batteries
   - Recent usage
   - Battery usage
   - Usage history
   - Battery capacity history
   - Battery life estimates
3. 生成中文 HTML 深度分析报告。
4. 重点优化图表：
   - 柱状图使用纯 CSS/HTML，便于直接调整。
   - 折线图使用纯 CSS/HTML 方案，无需额外依赖。
   - 长周期数据自动按月聚合，同时保留完整数据统计。
   - 容量对比采用“背景柱 + 当前容量柱”的层次设计。
   - 所有数值标签动态选择柱内/柱外位置，减少重叠。
   - 日期轴自动选择合适刻度。
   - 异常历史点不会把整张图的纵轴压扁。
   - 图表统一风格、留白和字体。
5. 支持两种运行方式：
   A. python battery_report_analyzer_optimized.py
      -> 自动调用 powercfg /batteryreport
   B. python battery_report_analyzer_optimized.py "battery-report.html"
      -> 直接读取已有报告，适合调试/重新出图。

依赖：
    pip install lxml

输出：
    %TEMP%\\BatteryReport\\笔记本电池深度报告.html
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import webbrowser
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import lxml.html
except ImportError as exc:
    print(f"缺少依赖库：{exc}")
    print("请运行：pip install lxml")
    input("按回车键退出...")
    raise SystemExit(1)


# ============================================================
# 全局样式
# ============================================================

COLORS = {
    "blue": "#2563EB",
    "blue_light": "#DBEAFE",
    "blue_mid": "#60A5FA",
    "green": "#10B981",
    "green_light": "#DCFCE7",
    "orange": "#F59E0B",
    "orange_dark": "#EA580C",
    "orange_light": "#FEF3C7",
    "red": "#EF4444",
    "red_light": "#FEE2E2",
    "slate": "#64748B",
    "slate_light": "#CBD5E1",
    "slate_lighter": "#F1F5F9",
    "text": "#0F172A",
    "muted": "#475569",
    "border": "#E2E8F0",
    "white": "#FFFFFF",
}


# ============================================================
# 基础解析
# ============================================================

def norm_text(node) -> str:
    if node is None:
        return ""
    return re.sub(r"\s+", " ", node.text_content()).strip()


def parse_mwh(text) -> Optional[float]:
    if text is None:
        return None
    s = str(text).replace(",", "").replace("mWh", "").strip()
    if not s or s in {"-", "—"}:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(match.group()) if match else None


def parse_int(text) -> Optional[int]:
    if text is None:
        return None
    match = re.search(r"-?\d+", str(text).replace(",", ""))
    return int(match.group()) if match else None


def parse_hms(text) -> Optional[int]:
    if not text:
        return None
    s = str(text).strip()
    if s in {"", "-", "—"}:
        return None

    m = re.fullmatch(r"(\d+):(\d{2}):(\d{2})", s)
    if m:
        h, minute, second = map(int, m.groups())
        return h * 3600 + minute * 60 + second

    m = re.fullmatch(r"(\d+):(\d{2})", s)
    if m:
        minute, second = map(int, m.groups())
        return minute * 60 + second

    return None


def parse_period(text) -> Tuple[Optional[dt.datetime], Optional[dt.datetime]]:
    if not text:
        return None, None
    dates = re.findall(r"(\d{4}-\d{2}-\d{2})", str(text))
    if not dates:
        return None, None
    try:
        start = dt.datetime.strptime(dates[0], "%Y-%m-%d")
        end = dt.datetime.strptime(dates[-1], "%Y-%m-%d")
    except ValueError:
        return None, None
    return start, end


def parse_datetime_text(text) -> Optional[dt.datetime]:
    if not text:
        return None
    s = re.sub(r"\s+", " ", str(text).strip())
    m = re.search(
        r"(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})",
        s,
    )
    if not m:
        return None
    try:
        return dt.datetime.strptime(
            f"{m.group(1)} {m.group(2)}",
            "%Y-%m-%d %H:%M:%S",
        )
    except ValueError:
        return None


def parse_time_from_row(row) -> Tuple[Optional[str], Optional[str]]:
    date_spans = row.xpath(".//span[contains(@class,'date')]")
    time_spans = row.xpath(".//span[contains(@class,'time')]")

    date_text = date_spans[0].text_content().strip() if date_spans else ""
    time_text = time_spans[0].text_content().strip() if time_spans else ""
    return date_text, time_text


def find_table_after_h2(root, keyword: str):
    keyword = keyword.lower()
    for h2 in root.xpath("//h2"):
        if keyword in h2.text_content().lower():
            node = h2.getnext()
            while node is not None:
                if node.tag == "table":
                    return node
                node = node.getnext()
    return None


# ============================================================
# 数据解析
# ============================================================

def parse_installed_batteries(table) -> List[dict]:
    if table is None:
        return []

    rows = table.xpath(".//tr")
    if not rows:
        return []

    header = rows[0].xpath("./td")
    count = max(len(header) - 1, 0)
    batteries = [{"slot": f"BATTERY {i + 1}"} for i in range(count)]

    key_map = {
        "NAME": "name",
        "MANUFACTURER": "manufacturer",
        "SERIAL NUMBER": "serial",
        "CHEMISTRY": "chemistry",
        "DESIGN CAPACITY": "design_capacity",
        "FULL CHARGE CAPACITY": "full_charge_capacity",
        "CYCLE COUNT": "cycle_count",
    }

    for row in rows[1:]:
        cells = row.xpath("./td")
        if len(cells) < 2:
            continue

        raw_key = norm_text(cells[0]).upper()
        if not raw_key:
            continue

        key = key_map.get(raw_key, raw_key.lower())
        for i in range(1, min(len(cells), count + 1)):
            batteries[i - 1][key] = norm_text(cells[i])

    for battery in batteries:
        battery["design_num"] = parse_mwh(battery.get("design_capacity"))
        battery["full_num"] = parse_mwh(battery.get("full_charge_capacity"))
        battery["cycle_num"] = parse_int(battery.get("cycle_count"))

        design = battery["design_num"]
        full = battery["full_num"]
        if design and full is not None and design > 0:
            battery["health"] = round(full / design * 100, 1)
            battery["loss"] = round(design - full, 1)
        else:
            battery["health"] = None
            battery["loss"] = None

    return batteries


def parse_recent_usage(table) -> List[dict]:
    if table is None:
        return []

    entries = []
    last_date = None

    for row in table.xpath(".//tr"):
        if not row.xpath('./td[contains(@class,"dateTime")]'):
            continue

        cells = row.xpath("./td")
        if len(cells) < 5:
            continue

        date_text, time_text = parse_time_from_row(row)
        if date_text:
            last_date = date_text

        full_dt = f"{last_date or ''} {time_text}".strip()
        percent_match = re.search(r"(\d+)\s*%", norm_text(cells[3]))

        entries.append(
            {
                "time_text": full_dt,
                "dt": parse_datetime_text(full_dt),
                "state": norm_text(cells[1]),
                "source": norm_text(cells[2]),
                "pct": int(percent_match.group(1)) if percent_match else None,
                "mwh": parse_mwh(norm_text(cells[4])),
            }
        )

    return entries


def parse_battery_usage(table) -> List[dict]:
    if table is None:
        return []

    entries = []
    last_date = None

    for row in table.xpath(".//tr"):
        if not row.xpath('./td[contains(@class,"dateTime")]'):
            continue

        cells = row.xpath("./td")
        if len(cells) < 5:
            continue

        date_text, time_text = parse_time_from_row(row)
        if date_text:
            last_date = date_text

        full_dt = f"{last_date or ''} {time_text}".strip()
        entries.append(
            {
                "time_text": full_dt,
                "dt": parse_datetime_text(full_dt),
                "state": norm_text(cells[1]),
                "duration_sec": parse_hms(norm_text(cells[2])),
                "percent": norm_text(cells[3]),
                "drain_mwh": parse_mwh(norm_text(cells[4])),
            }
        )

    return entries


def parse_usage_history(table) -> List[dict]:
    if table is None:
        return []

    entries = []
    for row in table.xpath(".//tr"):
        cells = row.xpath("./td")
        if len(cells) < 6:
            continue

        start, end = parse_period(norm_text(cells[0]))
        if start is None:
            continue

        entries.append(
            {
                "start": start,
                "end": end,
                "period_text": norm_text(cells[0]),
                "bat_active": parse_hms(norm_text(cells[1])) or 0,
                "bat_cs": parse_hms(norm_text(cells[2])) or 0,
                "ac_active": parse_hms(norm_text(cells[4])) or 0,
                "ac_cs": parse_hms(norm_text(cells[5])) or 0,
            }
        )

    return entries


def parse_capacity_history(table) -> List[dict]:
    if table is None:
        return []

    entries = []
    for row in table.xpath(".//tr"):
        cells = row.xpath("./td")
        if len(cells) < 3:
            continue

        start, end = parse_period(norm_text(cells[0]))
        if start is None:
            continue

        full = parse_mwh(norm_text(cells[1]))
        design = parse_mwh(norm_text(cells[2]))
        health = round(full / design * 100, 1) if full and design else None

        entries.append(
            {
                "start": start,
                "end": end,
                "full": full,
                "design": design,
                "health": health,
            }
        )

    return entries


def parse_life_estimates(table) -> List[dict]:
    if table is None:
        return []

    entries = []
    for row in table.xpath(".//tr"):
        cells = row.xpath("./td")
        if len(cells) < 6:
            continue

        start, end = parse_period(norm_text(cells[0]))
        if start is None:
            continue

        entries.append(
            {
                "start": start,
                "end": end,
                "fc_active": parse_hms(norm_text(cells[1])),
                "dc_active": parse_hms(norm_text(cells[4])),
            }
        )

    return entries


def parse_current_estimate(root) -> dict:
    for table in root.xpath("//table"):
        for row in table.xpath(".//tr"):
            cells = row.xpath("./td")
            if len(cells) >= 5 and "Since OS install" in norm_text(cells[0]):
                return {
                    "fc_active": parse_hms(norm_text(cells[1])),
                    "dc_active": parse_hms(norm_text(cells[4])),
                }
    return {}



def parse_system_info(root) -> dict:
    """读取 Windows battery-report 顶部系统信息。"""
    info = {}
    # 原始报告的第一张表即系统概况表。
    tables = root.xpath("//table")
    if not tables:
        return info
    for row in tables[0].xpath(".//tr"):
        cells = row.xpath("./td")
        if len(cells) < 2:
            continue
        key = norm_text(cells[0]).upper()
        value = norm_text(cells[1])
        mapping = {
            "COMPUTER NAME": "computer_name",
            "SYSTEM PRODUCT NAME": "system_product",
            "BIOS": "bios",
            "OS BUILD": "os_build",
            "PLATFORM ROLE": "platform_role",
            "CONNECTED STANDBY": "connected_standby",
            "REPORT TIME": "report_time",
        }
        if key in mapping:
            info[mapping[key]] = value
    return info


def _to_number(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _format_power_watts(mw: Optional[float]) -> Optional[str]:
    if mw is None or mw <= 0:
        return None
    watts = mw / 1000.0
    return f"{watts:.1f} W"


def _format_current_amps(mw: Optional[float], mv: Optional[float]) -> Optional[str]:
    if mw is None or mv is None or mw <= 0 or mv <= 0:
        return None
    amps = mw / mv
    if amps >= 1:
        return f"{amps:.2f} A"
    return f"{amps * 1000:.0f} mA"


def query_live_battery_status() -> dict:
    raw = None
    ps_command = (
        "Get-CimInstance Win32_Battery | "
        "Select-Object Name,DeviceID,BatteryStatus,EstimatedChargeRemaining,"
        "ChargeRate,DischargeRate,Voltage | ConvertTo-Json -Compress"
    )
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps_command,
            ],
            capture_output=True,
            text=True,
            timeout=8,
            shell=False,
        )
        if result.returncode == 0 and result.stdout.strip():
            raw = json.loads(result.stdout.strip())
    except (FileNotFoundError, subprocess.TimeoutExpired, json.JSONDecodeError, OSError):
        raw = None

    # 兼容没有 Get-CimInstance / ConvertTo-Json 的旧版 Windows。
    # WMIC 在多块电池场景下输出多组记录，逐组解析。
    rows = []
    if raw is not None:
        if isinstance(raw, dict):
            rows = [raw]
        elif isinstance(raw, list):
            rows = raw
    else:
        try:
            result = subprocess.run(
                [
                    "wmic",
                    "path",
                    "Win32_Battery",
                    "get",
                    "Name,DeviceID,BatteryStatus,EstimatedChargeRemaining,"
                    "ChargeRate,DischargeRate,Voltage",
                    "/value",
                ],
                capture_output=True,
                text=True,
                timeout=8,
                shell=False,
            )
            if result.returncode == 0:
                current = {}
                for line in result.stdout.splitlines() + [""]:
                    line = line.strip()
                    if line and "=" in line:
                        k, v = line.split("=", 1)
                        current[k.strip()] = v.strip()
                    elif not line and current:
                        rows.append(current)
                        current = {}
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            rows = []

    if not rows:
        return {}

    status_map = {
        1: "放电中",
        2: "接通 AC 电源",
        3: "已充满",
        4: "低电量",
        5: "极低电量",
        6: "充电中",
        7: "高电量充电中",
        8: "低电量充电中",
        9: "极低电量充电中",
        10: "状态未定义",
        11: "部分充电",
    }

    batteries_live = []
    for index, raw_row in enumerate(rows, 1):
        status_num = None
        try:
            status_num = int(float(raw_row.get("BatteryStatus")))
        except (TypeError, ValueError):
            pass

        charge_pct = _to_number(raw_row.get("EstimatedChargeRemaining"))
        charge_rate_mw = _to_number(raw_row.get("ChargeRate"))
        discharge_rate_mw = _to_number(raw_row.get("DischargeRate"))
        voltage_mv = _to_number(raw_row.get("Voltage"))

        status_label = status_map.get(status_num)
        detail_parts = []
        if voltage_mv:
            detail_parts.append(f"电池电压 {voltage_mv / 1000:.2f} V")
        if status_label:
            detail_parts.append("来自 Win32_Battery 实时状态")

        item = {
            "index": index,
            "name": (raw_row.get("Name") or "").strip() or None,
            "device_id": (raw_row.get("DeviceID") or "").strip() or None,
            "status_label": status_label,
            "status_num": status_num,
            "charge_percent_num": charge_pct,
            "charge_percent": f"{charge_pct:.0f}%" if charge_pct is not None else None,
            "charge_power": _format_power_watts(charge_rate_mw),
            "discharge_power": _format_power_watts(discharge_rate_mw),
            "charge_current": _format_current_amps(charge_rate_mw, voltage_mv),
            "discharge_current": _format_current_amps(discharge_rate_mw, voltage_mv),
            "detail": "；".join(detail_parts) if detail_parts else None,
        }
        batteries_live.append(item)

    # 兼容旧 HTML 字段：默认取第一块电池，同时保留全部电池列表。
    first = batteries_live[0]
    result = dict(first)
    result["batteries"] = batteries_live
    result["battery_count"] = len(batteries_live)
    return result

def infer_power_status_from_report(recent_usage: Sequence[dict]) -> dict:
    """仅根据 battery-report 的 Recent usage 推断状态；明确标注为历史记录推断。"""
    events = [e for e in recent_usage if e.get("dt")]
    if not events:
        return {}
    events.sort(key=lambda x: x["dt"])
    latest = events[-1]
    source = (latest.get("source") or "").strip().lower()
    latest_pct = latest.get("pct")

    if source == "battery":
        return {
            "label": "放电中（最近记录）",
            "detail": f"最近记录为电池供电，电量 {latest_pct}%" if latest_pct is not None else "最近记录为电池供电",
            "charge_percent": f"{latest_pct}%" if latest_pct is not None else None,
        }

    if source == "ac":
        previous = next((e for e in reversed(events[:-1]) if e.get("pct") is not None), None)
        if previous and latest_pct is not None and latest_pct > previous.get("pct", latest_pct):
            detail = f"最近记录显示电量由 {previous['pct']}% 升至 {latest_pct}%；battery-report 无法提供此刻充电电流/功率"
            return {
                "label": "充电中（根据最近记录推断）",
                "detail": detail,
                "charge_percent": f"{latest_pct}%",
            }
        return {
            "label": "接通 AC 电源",
            "detail": f"最近记录为 AC 供电，电量 {latest_pct}%" if latest_pct is not None else "最近记录为 AC 供电",
            "charge_percent": f"{latest_pct}%" if latest_pct is not None else None,
        }

    return {
        "label": latest.get("state") or "未知",
        "detail": "根据 battery-report 最近一条记录显示",
        "charge_percent": f"{latest_pct}%" if latest_pct is not None else None,
    }


def parse_report(path: str) -> dict:
    doc = lxml.html.parse(path)
    root = doc.getroot()

    data = {
        "system_info": parse_system_info(root),
        "batteries": parse_installed_batteries(
            find_table_after_h2(root, "Installed batteries")
        ),
        "recent_usage": parse_recent_usage(
            find_table_after_h2(root, "Recent usage")
        ),
        "battery_usage": parse_battery_usage(
            find_table_after_h2(root, "Battery usage")
        ),
        "usage_history": parse_usage_history(
            find_table_after_h2(root, "Usage history")
        ),
        "capacity_history": parse_capacity_history(
            find_table_after_h2(root, "Battery capacity history")
        ),
        "life_estimates": parse_life_estimates(
            find_table_after_h2(root, "Battery life estimates")
        ),
        "current_estimate": parse_current_estimate(root),
        "derived_power_status": {},
        "live_battery_status": {},
    }

    data["derived_power_status"] = infer_power_status_from_report(data["recent_usage"])
    return data


# ============================================================
# 输入文件
# ============================================================

def generate_battery_report() -> str:
    temp_dir = Path(tempfile.gettempdir())
    path = temp_dir / "battery_report_raw.html"

    if path.exists():
        try:
            path.unlink()
        except OSError:
            pass

    try:
        result = subprocess.run(
            ["powercfg", "/batteryreport", "/output", str(path)],
            capture_output=True,
            text=True,
            timeout=60,
            shell=False,
        )
    except FileNotFoundError:
        raise RuntimeError("找不到 powercfg，请在 Windows 上运行本程序。")
    except subprocess.TimeoutExpired:
        raise RuntimeError("powercfg /batteryreport 超时。")

    if result.returncode != 0 and not path.exists():
        detail = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"生成 battery-report 失败：{detail or '未知错误'}")

    if not path.exists():
        raise RuntimeError("battery-report.html 未生成。")

    print(f"✓ 已生成原始报告：{path}")
    return str(path)


def prepare_report_source(input_path: Optional[str]) -> str:
    if input_path:
        path = Path(input_path).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"找不到输入文件：{path}")
        print(f"✓ 使用已有 battery-report：{path}")
        return str(path)

    return generate_battery_report()


# ============================================================
# 图表工具
# ============================================================

def health_color(health: Optional[float]) -> str:
    if health is None:
        return COLORS["slate"]
    if health >= 80:
        return COLORS["green"]
    if health >= 60:
        return COLORS["orange"]
    return COLORS["red"]


def health_label(health: Optional[float]) -> str:
    if health is None:
        return "未知"
    if health >= 80:
        return "健康"
    if health >= 60:
        return "衰减"
    return "严重衰减"


def month_bucket(date: dt.datetime) -> dt.datetime:
    return dt.datetime(date.year, date.month, 1)


def aggregate_usage_monthly(history: Sequence[dict]) -> List[dict]:
    """
    把长期周/不规则区间记录聚合到月份。
    这里使用区间与月份的重叠天数按比例分摊，避免一个跨月区间全部记到某个月。
    """
    buckets = defaultdict(lambda: {"bat": 0.0, "ac": 0.0, "days": 0.0})

    for item in history:
        start = item.get("start")
        end = item.get("end")
        if not start:
            continue

        if end is None or end <= start:
            end = start + dt.timedelta(days=1)

        total_seconds = max((end - start).total_seconds(), 86400)
        cursor = start

        while cursor < end:
            next_month = (
                cursor.replace(day=28) + dt.timedelta(days=4)
            ).replace(day=1)
            segment_end = min(next_month, end)
            segment_seconds = max(
                (segment_end - cursor).total_seconds(),
                0,
            )
            ratio = segment_seconds / total_seconds

            key = dt.datetime(cursor.year, cursor.month, 1)
            buckets[key]["bat"] += (item["bat_active"] / 3600.0) * ratio
            buckets[key]["ac"] += (item["ac_active"] / 3600.0) * ratio
            buckets[key]["days"] += segment_seconds / 86400.0

            cursor = segment_end

    result = []
    for key in sorted(buckets):
        result.append(
            {
                "date": key,
                "bat": buckets[key]["bat"],
                "ac": buckets[key]["ac"],
            }
        )
    return result


# ============================================================
# 图 1：电池健康度
# ============================================================

def chart_health_compare(batteries: Sequence[dict], out_dir: Path):
    """纯 CSS/HTML 横向健康度条。"""
    valid = [b for b in batteries if b.get("health") is not None]
    if not valid:
        return None

    # 不排序：严格沿用 battery-card 的原始顺序。
    rows = []
    for b in valid:
        name = html.escape(str(b.get("name") or b.get("slot", "电池")))
        health = float(b["health"])
        color = health_color(health)
        rows.append(f"""
        <div class="cssbar-row">
          <div class="cssbar-label">{name}</div>
          <div class="health-track" aria-label="{name} 健康度 {health:.1f}%">
            <div class="health-zone zone-bad"></div>
            <div class="health-zone zone-warn"></div>
            <div class="health-zone zone-good"></div>
            <div class="health-fill" style="width:{min(max(health,0),100):.1f}%;background:{color}"></div>
            <span class="health-value" style="left:{min(max(health,0),100):.1f}%">{health:.1f}%</span>
          </div>
        </div>
        """)

    return "电池健康度对比", f"""
    <div class="css-chart health-chart">
      <div class="health-axis-wrap">
        <div class="cssbar-axis">
          <span>0</span><span>20</span><span>40</span><span>60</span><span>80</span><span>100</span>
        </div>
        <div class="health-zone-caption">
          <span class="bad">严重衰减</span><span class="warn">衰减</span><span class="good">健康</span>
        </div>
      </div>
      {''.join(rows)}
      <div class="health-reference"><span></span>80% 健康参考线</div>
      <div class="cssbar-xlabel">健康度 (%)</div>
    </div>
    """, "css"


def chart_capacity_compare(batteries: Sequence[dict], out_dir: Path):
    """纯 CSS/HTML 容量横向条：设计容量为轨道，当前满充容量为彩色填充。"""
    valid = [
        b for b in batteries
        if b.get("design_num") and b.get("full_num")
    ]
    if not valid:
        return None

    max_design = max(float(b["design_num"]) for b in valid)
    rows = []
    for b in valid:  # 不排序，严格跟随 battery-card
        name = html.escape(str(b.get("name") or b.get("slot", "电池")))
        design = float(b["design_num"])
        full = float(b["full_num"])
        health = float(b.get("health") or 0)
        loss = max(0.0, 100.0 - health)
        pct_of_design = min(max(full / design * 100.0, 0.0), 100.0)
        color = health_color(health)
        rows.append(f"""
        <div class="cap-row">
          <div class="cssbar-label">{name}</div>
          <div class="cap-main">
            <div class="cap-track" aria-label="{name} 当前满充容量">
              <div class="cap-fill" style="width:{pct_of_design:.1f}%;background:{color}">
                <span class="cap-fill-value">{full/1000:.1f} Ah</span>
              </div>
            </div>
            <div class="cap-meta">
              <span>设计 {design/1000:.1f} Ah</span>
              <b style="color:{color}">{health:.1f}%</b>
              <span>损耗 {loss:.1f}%</span>
            </div>
          </div>
        </div>
        """)

    return "各电池容量对比", f"""
    <div class="css-chart capacity-chart">
      <div class="capacity-head">
        <span>彩色 = 当前满充容量</span>
        <span>浅灰 = 设计容量</span>
      </div>
      {''.join(rows)}
      <div class="capacity-axis">
        <span>0 Ah</span>
        <span>{max_design/4000:.0f} Ah</span>
        <span>{max_design/2000:.0f} Ah</span>
        <span>{max_design/1333.3333:.0f} Ah</span>
        <span>{max_design/1000:.0f} Ah</span>
      </div>
      <div class="cssbar-xlabel">容量（Ah）</div>
    </div>
    """, "css"


def _css_line_chart(
    title: str,
    dates: Sequence[dt.datetime],
    series: List[dict],
    y_label: str,
    fill_between: Optional[Tuple[int, int]] = None,
    annotations: Optional[List[dict]] = None,
    health_badge: Optional[Tuple[str, str]] = None,
    y_format: str = "{:,.0f}",
) -> Tuple[str, str, str]:
    """纯 CSS/HTML 折线图（inline SVG + viewBox），无需 Matplotlib，自适应宽度。"""
    n = len(dates)
    plot_w = 1000
    plot_h = 340
    pad_l, pad_r, pad_t, pad_b = 60, 24, 24, 56
    inner_w = plot_w - pad_l - pad_r
    inner_h = plot_h - pad_t - pad_b

    xs = [pad_l + (i / max(n - 1, 1)) * inner_w for i in range(n)]

    all_vals = []
    for s in series:
        all_vals.extend(s["values"])
    y_max = max(all_vals) * 1.1
    y_min = min(all_vals)
    if y_min > 0:
        y_min = max(0, y_min - (y_max - y_min) * 0.15)
    y_range = y_max - y_min if y_max > y_min else 1

    def to_y(val):
        return pad_t + inner_h - ((val - y_min) / y_range) * inner_h

    # Y-axis ticks + grid
    n_ticks = 5
    grid_parts = []
    y_axis_parts = []
    for i in range(n_ticks + 1):
        val = y_min + (y_max - y_min) * i / n_ticks
        py = to_y(val)
        grid_parts.append(
            f'<line x1="{pad_l}" y1="{py:.1f}" x2="{plot_w - pad_r}" y2="{py:.1f}" '
            f'stroke="#f1f5f9" stroke-width="1"/>'
        )
        y_axis_parts.append(
            f'<text x="{pad_l - 8}" y="{py + 3.5:.1f}" '
            f'text-anchor="end" font-size="10" fill="#94a3b8">'
            f'{y_format.format(val)}</text>'
        )

    # Fill area between two series
    fill_svg = ""
    if fill_between is not None:
        ti, bi = fill_between
        top_ys = [to_y(v) for v in series[ti]["values"]]
        bot_ys = [to_y(v) for v in series[bi]["values"]]
        pts = []
        for x, y in zip(xs, top_ys):
            pts.append(f"{x:.1f},{y:.1f}")
        for x, y in zip(reversed(xs), reversed(bot_ys)):
            pts.append(f"{x:.1f},{y:.1f}")
        fill_color = series[ti].get("fill_color", COLORS["red_light"])
        fill_opacity = series[ti].get("fill_opacity", 0.4)
        fill_svg = (
            f'<polygon points="{" ".join(pts)}" '
            f'fill="{fill_color}" opacity="{fill_opacity}"/>'
        )

    # Lines + markers
    line_parts = []
    dot_parts = []
    for s in series:
        vals = s["values"]
        ys = [to_y(v) for v in vals]
        color = s["color"]
        lw = s.get("line_width", 2)
        dashed = s.get("dashed", False)
        ms = s.get("marker_size", 0)

        pts_str = " ".join(f"{x:.1f},{y:.1f}" for x, y in zip(xs, ys))
        dash_attr = ' stroke-dasharray="6,4"' if dashed else ""
        line_parts.append(
            f'<polyline points="{pts_str}" fill="none" '
            f'stroke="{color}" stroke-width="{lw}" stroke-linejoin="round" '
            f'stroke-linecap="round"{dash_attr}/>'
        )

        if ms > 0:
            for x, y in zip(xs, ys):
                dot_parts.append(
                    f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{ms / 2:.1f}" '
                    f'fill="{color}" stroke="white" stroke-width="1.5"/>'
                )

    # Latest-point highlight + annotation
    ann_parts = []
    if annotations:
        for ann in annotations:
            idx = ann["series_index"]
            x = xs[-1]
            y = to_y(series[idx]["values"][-1])
            color = ann.get("color", COLORS["blue"])
            text = ann["text"]
            ann_parts.append(
                f'<circle cx="{x:.1f}" cy="{y:.1f}" r="6" '
                f'fill="{color}" stroke="white" stroke-width="2.5"/>'
            )
            text_y = y + 18
            ann_parts.append(
                f'<text x="{x:.1f}" y="{text_y:.1f}" '
                f'text-anchor="end" font-size="11" font-weight="700" '
                f'fill="#dc2626">{html.escape(text)}</text>'
            )

    # X-axis labels: auto-distribute based on chart width to avoid overlap
    fmt = "%Y-%m"
    label_y = plot_h - pad_b + 14
    x_parts = []

    if dates:
        # Each rotated label needs ~35px of horizontal space to avoid overlap
        min_label_px = 35
        max_labels = max(2, int(inner_w / min_label_px))
        # Evenly sample data point indices across the full range
        num_ticks = min(max_labels, n)
        for k in range(num_ticks):
            idx = round(k * (n - 1) / max(num_ticks - 1, 1))
            x = xs[idx]
            x_parts.append(
                f'<text x="{x:.1f}" y="{label_y:.1f}" '
                f'transform="rotate(-45 {x:.1f} {label_y:.1f})" '
                f'text-anchor="end" font-size="9" fill="#94a3b8">'
                f'{dates[idx].strftime(fmt)}</text>'
            )

    # Legend (inside SVG, top area)
    leg_x = pad_l
    leg_y = 14
    leg_items = []
    cur_x = leg_x
    for s in series:
        if not s.get("label"):
            continue
        if s.get("dashed"):
            icon = (
                f'<line x1="0" y1="6" x2="20" y2="6" '
                f'stroke="{s["color"]}" stroke-width="2" '
                f'stroke-dasharray="6,4"/>'
            )
        else:
            icon = (
                f'<rect x="0" y="{6 - s.get("line_width", 2) / 2:.1f}" '
                f'width="20" height="{s.get("line_width", 2)}" '
                f'fill="{s["color"]}" rx="2"/>'
            )
        leg_items.append(
            f'<g transform="translate({cur_x:.0f},{leg_y})">'
            f'{icon}'
            f'<text x="26" y="9.5" font-size="10" fill="#64748b">'
            f'{html.escape(s["label"])}</text></g>'
        )
        cur_x += 130

    # Health badge (bottom-right)
    badge_svg = ""
    if health_badge:
        badge_text, badge_color = health_badge
        bw, bh = 140, 22
        bx = plot_w - pad_r - bw
        by = plot_h - pad_b - bh - 6
        badge_svg = (
            f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bw}" height="{bh}" '
            f'rx="8" fill="white" stroke="{badge_color}" stroke-width="1"/>'
            f'<text x="{bx + bw / 2:.1f}" y="{by + 15:.1f}" '
            f'text-anchor="middle" font-size="11" font-weight="700" '
            f'fill="{badge_color}">{html.escape(badge_text)}</text>'
        )

    svg = f"""<svg viewBox="0 0 {plot_w} {plot_h}" preserveAspectRatio="xMidYMid meet" style="width:100%;height:auto;display:block;font-family:inherit;">
  {''.join(grid_parts)}
  {fill_svg}
  {''.join(line_parts)}
  {''.join(dot_parts)}
  {''.join(ann_parts)}
  {badge_svg}
  {''.join(y_axis_parts)}
  {''.join(x_parts)}
  {''.join(leg_items)}
</svg>"""

    chart_html = f"""
    <div class="css-chart line-chart">
      {svg}
      <div class="cssbar-xlabel">{html.escape(y_label)}</div>
    </div>
    """

    return title, chart_html, "css"


def chart_capacity_history(history: Sequence[dict], out_dir: Path):
    """纯 CSS 折线图：电池容量变化历史。"""
    valid = [
        h for h in history
        if h.get("start") and h.get("full") and h.get("design") and h["full"] > 0
    ]
    if len(valid) < 2:
        return None

    valid = sorted(valid, key=lambda x: x["start"])
    dates = [h["start"] for h in valid]
    fulls = [h["full"] for h in valid]
    designs = [h["design"] for h in valid]

    # 中值趋势线
    window = 9
    trend = []
    for i in range(len(fulls)):
        left = max(0, i - window // 2)
        right = min(len(fulls), i + window // 2 + 1)
        values = sorted(fulls[left:right])
        mid = len(values) // 2
        trend.append(values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2)

    current = fulls[-1]
    current_health = current / designs[-1] * 100 if designs[-1] else None

    series = [
        {"values": designs, "color": COLORS["slate_light"], "label": "设计容量", "dashed": True, "line_width": 2},
        {"values": fulls, "color": COLORS["blue_mid"], "label": "历史记录", "line_width": 2, "marker_size": 4, "fill_color": COLORS["red_light"], "fill_opacity": 0.42},
        {"values": trend, "color": COLORS["blue"], "label": "中值趋势", "line_width": 3},
    ]

    annotations = [{"series_index": 2, "text": f"最新 {current:,.0f} mWh", "color": COLORS["blue"]}]
    health_badge = None
    if current_health is not None:
        health_badge = (f"最新健康度 {current_health:.1f}%", health_color(current_health))

    return _css_line_chart(
        "电池容量变化历史", dates, series, "容量 (mWh)",
        fill_between=(0, 1), annotations=annotations, health_badge=health_badge,
    )


# ============================================================
# 图 4：续航估算历史
# ============================================================

def chart_life_estimates(history: Sequence[dict], out_dir: Path):
    """纯 CSS 折线图：续航估算历史。"""
    valid = [
        h for h in history
        if h.get("start") and h.get("fc_active") and h.get("dc_active")
    ]
    if len(valid) < 2:
        return None

    valid = sorted(valid, key=lambda x: x["start"])
    dates = [h["start"] for h in valid]
    fc = [h["fc_active"] / 3600 for h in valid]
    dc = [h["dc_active"] / 3600 for h in valid]

    series = [
        {"values": dc, "color": COLORS["blue_mid"], "label": "按设计容量估算", "line_width": 2, "marker_size": 4, "fill_color": COLORS["orange_light"], "fill_opacity": 0.38},
        {"values": fc, "color": COLORS["blue"], "label": "按当前满充容量估算", "line_width": 3, "marker_size": 4},
    ]

    annotations = [{"series_index": 1, "text": f"{fc[-1]:.2f} h", "color": COLORS["blue"]}]

    return _css_line_chart(
        "续航估算历史", dates, series, "续航时间 (小时)",
        fill_between=(0, 1), annotations=annotations,
        y_format="{:.1f}",
    )


# ============================================================
# 图 5：使用时长历史
# ============================================================

def chart_usage_history(history: Sequence[dict], out_dir: Path):
    """纯 CSS 月度堆叠柱图。"""
    aggregated = aggregate_usage_monthly(history)
    aggregated = [x for x in aggregated if x["bat"] > 0 or x["ac"] > 0]
    if not aggregated:
        return None

    dates = [x["date"] for x in aggregated]
    bat = [x["bat"] for x in aggregated]
    ac = [x["ac"] for x in aggregated]
    totals = [a+b for a,b in zip(bat, ac)]
    max_total = max(totals) if totals else 1

    # 柱图自适应容器宽度，月份多时自动缩窄。
    bars = []
    for d, bv, av, tv in zip(dates, bat, ac, totals):
        bat_h = bv / max_total * 100 if max_total else 0
        ac_h = av / max_total * 100 if max_total else 0
        bars.append(f"""
        <div class="month-bar-item" title="{d.strftime('%Y-%m')}：电池 {bv:.1f}h，AC {av:.1f}h">
          <div class="month-bar-stack">
            <div class="month-seg bat" style="height:{bat_h:.2f}%"></div>
            <div class="month-seg ac" style="height:{ac_h:.2f}%"></div>
          </div>
          <div class="month-label">{d.strftime('%y-%m')}</div>
        </div>
        """)

    total_bat = sum(bat)
    total_ac = sum(ac)
    total = total_bat + total_ac
    ratio = total_bat / total * 100 if total else 0
    return "月度使用时长历史（AC vs 电池）", f"""
    <div class="css-chart usage-chart">
      <div class="usage-summary">
        <span><i class="legend-dot bat"></i>电池 {total_bat:,.0f} h</span>
        <span><i class="legend-dot ac"></i>AC {total_ac:,.0f} h</span>
        <b>电池供电占比 {ratio:.1f}%</b>
      </div>
      <div class="month-chart-area">
        {''.join(bars)}
      </div>
      <div class="cssbar-xlabel">月份</div>
    </div>
    """, "css"


def chart_recent_activity(
    recent_usage: Sequence[dict],
    battery_usage: Sequence[dict],
    out_dir: Path,
):
    """纯 CSS 活动分析：放电事件采用表格式横向条，彻底避免文字覆盖。"""
    events = []
    for r in recent_usage:
        event_dt = r.get("dt") or parse_datetime_text(r.get("time_text"))
        if event_dt is None:
            continue
        events.append({"dt": event_dt, "state": r.get("state", ""), "source": r.get("source", "")})
    if len(events) < 2:
        return None
    events.sort(key=lambda e: e["dt"])

    MAX_REASONABLE = 4 * 3600
    state_duration = defaultdict(float)
    source_duration = defaultdict(float)
    for cur, nxt in zip(events, events[1:]):
        gap = (nxt["dt"] - cur["dt"]).total_seconds()
        if 0 < gap <= MAX_REASONABLE:
            state_duration[cur["state"] or "未知"] += gap
            source_duration[cur["source"] or "未知"] += gap

    valid_events = [e for e in battery_usage if e.get("duration_sec") and e.get("drain_mwh")]
    if not valid_events and not state_duration:
        return None

    max_minutes = max((float(e["duration_sec"])/60 for e in valid_events), default=1)
    event_rows = []
    total_duration = 0.0
    total_drain = 0.0
    powers = []
    for i, event in enumerate(valid_events):
        event_dt = event.get("dt") or parse_datetime_text(event.get("time_text"))
        label = event_dt.strftime("%m-%d %H:%M") if event_dt else f"事件 {i+1}"
        sec = float(event["duration_sec"])
        drain = float(event["drain_mwh"])
        minutes = sec / 60.0
        power = drain / (sec / 3600.0) / 1000.0 if sec > 0 else 0.0
        total_duration += minutes
        total_drain += drain
        powers.append(power)
        width = minutes / max_minutes * 100
        event_rows.append(f"""
        <div class="event-row">
          <div class="event-time">{html.escape(label)}</div>
          <div class="event-track"><div class="event-fill" style="width:{width:.1f}%"></div></div>
          <div class="event-minutes">{minutes:.1f} 分钟</div>
          <div class="event-detail">{drain:,.0f} mWh · {power:.1f} W</div>
        </div>
        """)

    avg_power = sum(powers)/len(powers) if powers else 0.0

    state_map = {"Active":"活动使用", "Suspended":"睡眠/挂起", "ConnectedStandby":"连接待机", "Report generated":"报告生成"}
    state_colors = {"Active":"#3b82f6", "Suspended":"#f59e0b", "ConnectedStandby":"#10b981", "Report generated":"#94a3b8"}
    state_items = [(state_map.get(k,k), v/3600.0, state_colors.get(k,"#64748b")) for k,v in state_duration.items() if v>0]
    state_total = sum(v for _,v,_ in state_items) or 1
    state_html = ''.join(
        f'<div class="simple-stat-row"><span><i class="legend-dot" style="background:{c}"></i>{html.escape(label)}</span><b>{hours:.1f} h</b><em>{hours/state_total*100:.1f}%</em></div>'
        for label,hours,c in state_items
    )

    source_map = {"AC":"AC 供电", "Battery":"电池供电", "未知":"未知"}
    source_colors = {"AC":"#f97316", "Battery":"#2563eb", "未知":"#94a3b8"}
    source_items = [(source_map.get(k,k), v, source_colors.get(k,"#94a3b8")) for k,v in source_duration.items() if v>0]
    source_total = sum(v for _,v,_ in source_items) or 1
    source_segments = ''.join(
        f'<div class="source-seg" style="width:{v/source_total*100:.2f}%;background:{c}" title="{html.escape(label)} {v/source_total*100:.1f}%"></div>'
        for label,v,c in source_items
    )
    source_legend = ''.join(
        f'<span><i class="legend-dot" style="background:{c}"></i>{html.escape(label)} {v/source_total*100:.1f}%</span>'
        for label,v,c in source_items
    )

    source_range = ""
    if events:
        first_dt = events[0]["dt"]
        last_dt = events[-1]["dt"]
        if first_dt.strftime("%Y-%m-%d") == last_dt.strftime("%Y-%m-%d"):
            source_range = f"{first_dt.strftime('%m-%d %H:%M')}~{last_dt.strftime('%H:%M')}"
        else:
            source_range = f"{first_dt.strftime('%m-%d %H:%M')}~{last_dt.strftime('%m-%d %H:%M')}"

    return "最近使用活动分析", f"""
    <div class="css-chart activity-chart">
      <div class="activity-topbar">
        <span>总时长 <b>{total_duration:.1f} 分钟</b></span>
        <span>总耗电 <b>{total_drain:,.0f} mWh</b></span>
        <span>平均功率 <b>{avg_power:.1f} W</b></span>
      </div>
      <div class="activity-section-title">最近 3 天放电事件</div>
      <div class="event-list">
        {''.join(event_rows) if event_rows else '<div class="empty-note">没有可用的放电事件</div>'}
      </div>
      <div class="activity-foot-grid">
        <div class="activity-subcard">
          <div class="activity-subtitle">最近状态持续时间（估算）</div>
          <div class="state-note">根据相邻 Recent usage 记录时间差估算；超过 4 小时的断档不计入。</div>
          {state_html or '<div class="empty-note">无可用数据</div>'}
        </div>
        <div class="activity-subcard">
          <div class="activity-subtitle">供电来源占比<span class="source-range-note">（{html.escape(source_range)}）</span></div>
          <div class="state-note">统计自 Recent usage 相邻记录间隔（≤4h）的供电来源。</div>
          <div class="source-track">{source_segments or '<div class="source-seg" style="width:100%;background:#cbd5e1"></div>'}</div>
          <div class="source-legend">{source_legend or '<span>无可用数据</span>'}</div>
        </div>
      </div>
    </div>
    """, "css"


def svg_ring(
    percent: float,
    size: int = 132,
    stroke: int = 12,
    color: str = COLORS["blue"],
    bg: str = COLORS["border"],
    label: str = "",
    sublabel: str = "",
) -> str:
    radius = (size - stroke) / 2
    circumference = 2 * 3.141592653589793 * radius
    pct = min(max(percent, 0), 100)
    offset = circumference * (1 - pct / 100)

    return f"""
    <svg width="{size}" height="{size}" viewBox="0 0 {size} {size}"
         aria-label="{html.escape(label)}">
      <circle cx="{size/2}" cy="{size/2}" r="{radius}"
              fill="none" stroke="{bg}" stroke-width="{stroke}"/>
      <circle cx="{size/2}" cy="{size/2}" r="{radius}"
              fill="none" stroke="{color}" stroke-width="{stroke}"
              stroke-dasharray="{circumference}"
              stroke-dashoffset="{offset}"
              stroke-linecap="round"
              transform="rotate(-90 {size/2} {size/2})"/>
      <text x="{size/2}" y="{size/2}"
            text-anchor="middle" dominant-baseline="middle"
            font-size="{size*0.21}" font-weight="700" fill="{color}">
        {html.escape(label)}
      </text>
      <text x="{size/2}" y="{size/2 + size*0.20}"
            text-anchor="middle"
            font-size="{size*0.10}" fill="{COLORS["slate"]}">
        {html.escape(sublabel)}
      </text>
    </svg>
    """


# ============================================================
# HTML 报告
# ============================================================

def generate_output_html(
    data: dict,
    chart_files: Sequence[tuple],
    report_dir: Path,
) -> str:
    batteries = data["batteries"]
    capacity_history = data["capacity_history"]
    current_est = data.get("current_estimate") or {}

    total_design = sum(b.get("design_num") or 0 for b in batteries)
    total_full = sum(b.get("full_num") or 0 for b in batteries)
    total_cycles = sum(b.get("cycle_num") or 0 for b in batteries)

    combined_health = (
        total_full / total_design * 100
        if total_design
        else None
    )

    first_valid = next(
        (h["full"] for h in capacity_history if h.get("full")),
        None,
    )
    last_valid = next(
        (h["full"] for h in reversed(capacity_history) if h.get("full")),
        None,
    )
    cap_decay = (
        (first_valid - last_valid) / first_valid * 100
        if first_valid and last_valid
        else None
    )

    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")

    combined_display = (
        f"{combined_health:.1f}<span class='stat-unit'>%</span>"
        if combined_health is not None
        else "-"
    )

    overview = f"""
    <div class="stat-grid">
      <div class="stat-card">
        <div class="stat-label">设计容量合计</div>
        <div class="stat-value">{total_design:,.0f}<span class="stat-unit">mWh</span></div>
      </div>
      <div class="stat-card">
        <div class="stat-label">当前满充容量</div>
        <div class="stat-value">{total_full:,.0f}<span class="stat-unit">mWh</span></div>
      </div>
      <div class="stat-card">
        <div class="stat-label">组合健康度</div>
        <div class="stat-value" style="color:{health_color(combined_health)}">{combined_display}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">累计循环次数</div>
        <div class="stat-value">{total_cycles}<span class="stat-unit">次</span></div>
      </div>
    </div>
    """

    battery_cards = []
    for battery in batteries:
        name = battery.get("name") or battery.get("slot", "电池")
        health = battery.get("health")
        hc = health_color(health)

        ring = svg_ring(
            health or 0,
            size=134,
            stroke=12,
            color=hc,
            label=f"{health:.1f}%" if health is not None else "--",
            sublabel=health_label(health),
        )

        details = [
            ("制造商", battery.get("manufacturer")),
            ("序列号", battery.get("serial")),
            ("化学类型", battery.get("chemistry")),
        ]
        if battery.get("design_num"):
            details.append(
                ("设计容量", f"{battery['design_num']:,.0f} mWh")
            )
        if battery.get("full_num"):
            details.append(
                ("满充容量", f"{battery['full_num']:,.0f} mWh")
            )
        if battery.get("loss") is not None:
            details.append(
                ("容量损耗", f"{battery['loss']:,.0f} mWh")
            )
        if battery.get("cycle_num") is not None:
            details.append(
                ("循环次数", f"{battery['cycle_num']} 次")
            )
        if battery.get("loss") is not None and battery.get("cycle_num"):
            details.append(
                (
                    "每周期损耗",
                    f"{battery['loss']/battery['cycle_num']:.2f} mWh/次",
                )
            )

        kv_html = "".join(
            f"""
            <div class="kv">
              <span class="kv-k">{html.escape(str(k))}</span>
              <span class="kv-v">{html.escape(str(v))}</span>
            </div>
            """
            for k, v in details
            if v not in (None, "")
        )

        battery_cards.append(
            f"""
            <div class="battery-card">
              <div class="battery-head">
                <div class="battery-title">
                  🔋 {html.escape(str(name))}
                </div>
                <span class="badge"
                      style="background:{hc}18;color:{hc};border:1px solid {hc}55">
                  {health_label(health)}
                </span>
              </div>

              <div class="battery-body">
                <div class="ring-wrap">{ring}</div>
                <div class="kv-grid">{kv_html}</div>
              </div>
            </div>
            """
        )

    system_info = data.get("system_info") or {}
    report_time = system_info.get("report_time") or now
    live = data.get("live_battery_status") or {}
    derived = data.get("derived_power_status") or {}

    # 当前电量：Windows Win32_Battery 通常将多块电池合并为一个实例，无法分别显示。
    # 当报告中有 2+ 块电池但实时数据只有 1 个实例时，标注为"当前总电量"。
    live_batteries = live.get("batteries") or []
    multi_battery = len(batteries) > 1
    charge_items = []
    if live_batteries:
        if len(live_batteries) == 1 and multi_battery:
            pct = live_batteries[0].get("charge_percent")
            if pct:
                charge_items.append(("当前总电量", pct))
        else:
            for item in live_batteries:
                name = item.get("name") or f"电池 {item.get('index', 0)}"
                pct = item.get("charge_percent")
                if pct:
                    charge_items.append((f"当前电量（{name}）", pct))
    else:
        derived_pct = derived.get("charge_percent") or live.get("charge_percent")
        if derived_pct:
            label = "当前总电量" if multi_battery else "当前电量"
            charge_items.append((label, derived_pct))

    system_items = [
        ("电脑名称", system_info.get("computer_name")),
        ("产品型号", system_info.get("system_product")),
        ("BIOS", system_info.get("bios")),
        ("Windows 版本", system_info.get("os_build")),
    ]
    system_items.extend(charge_items)
    system_items.extend([
        ("当前充电功率", live.get("charge_power")),
        ("当前充电电流", live.get("charge_current")),
        ("当前放电功率", live.get("discharge_power")),
        ("当前放电电流", live.get("discharge_current")),
        ("报告时间", report_time),
    ])
    system_items = [(k, v) for k, v in system_items if v]
    system_html = ""
    if system_items:
        system_html = f"""
        <div class="card system-card">
          <h2>💻 系统与报告信息</h2>
          <div class="system-grid">
            {''.join(f'<div class="system-item"><span>{html.escape(k)}</span><b>{html.escape(str(v))}</b></div>' for k, v in system_items)}
          </div>
        </div>
        """

    current_html = ""
    if current_est.get("fc_active") or current_est.get("dc_active"):
        fc_h = (current_est.get("fc_active") or 0) / 3600
        dc_h = (current_est.get("dc_active") or 0) / 3600
        loss_h = max(0, dc_h - fc_h)

        current_html = f"""
        <div class="card">
          <h2>⚡ 当前续航估算</h2>
          <p class="explain">
            基于 Windows battery report 中“Since OS install”记录计算。
          </p>
          <div class="est-grid">
            <div class="est-item">
              <div class="est-label">按设计容量估算</div>
              <div class="est-value est-blue">{dc_h:.2f}<span class="est-unit">小时</span></div>
            </div>
            <div class="est-item">
              <div class="est-label">按当前满充容量估算</div>
              <div class="est-value est-red">{fc_h:.2f}<span class="est-unit">小时</span></div>
            </div>
            <div class="est-item">
              <div class="est-label">老化对应损失</div>
              <div class="est-value est-orange">-{loss_h:.2f}<span class="est-unit">小时</span></div>
            </div>
          </div>
        </div>
        """

    avg_health_values = [
        b["health"] for b in batteries if b.get("health") is not None
    ]
    avg_health = (
        sum(avg_health_values) / len(avg_health_values)
        if avg_health_values
        else None
    )

    if avg_health is None:
        summary = "当前没有足够的健康度数据进行综合判断。"
    elif avg_health >= 85:
        summary = "电池组整体状态优秀，可以继续正常使用。"
    elif avg_health >= 70:
        summary = "电池组整体状态良好，但已经能看到一定容量衰减。"
    elif avg_health >= 50:
        summary = "电池组存在明显衰减，建议关注实际续航变化。"
    else:
        summary = "电池组衰减较明显，建议认真评估更换电池的必要性。"

    advice = []
    for battery in batteries:
        h = battery.get("health")
        name = battery.get("name") or battery.get("slot", "电池")
        if h is None:
            continue
        if h < 40:
            advice.append(
                f"{name} 健康度仅 {h:.1f}%，已经属于严重衰减。"
            )
        elif h < 60:
            advice.append(
                f"{name} 健康度为 {h:.1f}%，续航能力已经明显下降。"
            )
        elif h < 80:
            advice.append(
                f"{name} 健康度为 {h:.1f}%，低于 80% 参考线，建议观察续航。"
            )

        cycles = battery.get("cycle_num")
        if cycles is not None and cycles >= 800:
            advice.append(
                f"{name} 已记录 {cycles} 次循环，循环次数较高。"
            )

    if cap_decay is not None and cap_decay > 30:
        advice.append(
            f"容量历史首末记录相比，满充容量累计变化约 {cap_decay:.1f}%。"
        )

    if not advice:
        advice.append("目前没有发现需要立即处理的电池状态异常。")

    advice_html = "".join(f"<li>{html.escape(x)}</li>" for x in advice)

    chart_cards_parts = []
    for item in chart_files:
        title, markup, kind = item
        wrapper_class = "chart-css" if kind == "css" else "chart-svg"
        chart_cards_parts.append(f"""
        <div class="chart-card">
          <div class="chart-title">{html.escape(title)}</div>
          <div class="{wrapper_class}">{markup}</div>
        </div>
        """)
    chart_cards = "".join(chart_cards_parts)

    page = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>笔记本电池深度分析报告</title>
<style>
* {{ box-sizing: border-box; }}
html {{ scroll-behavior: smooth; }}
body {{
  margin: 0;
  padding: 28px 16px 44px;
  color: {COLORS["text"]};
  background:
    radial-gradient(circle at top left, #eff6ff 0, transparent 33%),
    linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%);
  font-family:
    -apple-system, BlinkMacSystemFont, "Segoe UI",
    "Microsoft YaHei", "PingFang SC", sans-serif;
}}
.container {{
  max-width: 1120px;
  margin: 0 auto;
}}
.header {{
  text-align: center;
  margin-bottom: 26px;
}}
.header h1 {{
  margin: 0 0 7px;
  font-size: 28px;
  line-height: 1.25;
  letter-spacing: -0.6px;
}}
.meta {{
  color: {COLORS["slate"]};
  font-size: 13px;
}}
.stat-grid {{
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 12px;
  margin-bottom: 18px;
}}
.stat-card {{
  background: white;
  border: 1px solid rgba(226,232,240,.7);
  border-radius: 15px;
  padding: 16px 18px;
  box-shadow: 0 6px 18px rgba(15,23,42,.055);
}}
.stat-label {{
  color: {COLORS["slate"]};
  font-size: 12px;
  margin-bottom: 5px;
}}
.stat-value {{
  font-size: 23px;
  font-weight: 750;
  letter-spacing: -.5px;
}}
.stat-unit {{
  margin-left: 4px;
  font-size: 12px;
  font-weight: 500;
  color: #94a3b8;
}}
.card, .battery-card {{
  background: white;
  border: 1px solid rgba(226,232,240,.72);
  border-radius: 17px;
  padding: 21px;
  margin-bottom: 18px;
  box-shadow: 0 7px 20px rgba(15,23,42,.045);
}}
.card h2 {{
  margin: 0 0 10px;
  padding-left: 10px;
  border-left: 3px solid {COLORS["blue"]};
  font-size: 16px;
}}
.explain {{
  margin: 0 0 15px;
  color: {COLORS["slate"]};
  font-size: 12px;
}}
.battery-head {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  margin-bottom: 18px;
}}
.battery-title {{
  font-size: 17px;
  font-weight: 700;
}}
.badge {{
  padding: 5px 11px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 650;
}}
.battery-body {{
  display: grid;
  grid-template-columns: 165px minmax(0, 1fr);
  gap: 24px;
  align-items: center;
}}
.ring-wrap {{
  display: flex;
  justify-content: center;
}}
.kv-grid {{
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 6px 24px;
}}
.kv {{
  display: flex;
  justify-content: space-between;
  gap: 14px;
  padding: 7px 0;
  border-bottom: 1px dashed #edf2f7;
  font-size: 13px;
}}
.kv-k {{
  color: {COLORS["slate"]};
}}
.kv-v {{
  color: {COLORS["text"]};
  font-weight: 550;
  text-align: right;
}}
.est-grid {{
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
}}
.est-item {{
  padding: 16px;
  border-radius: 12px;
  text-align: center;
  background: #f8fafc;
}}
.est-label {{
  margin-bottom: 8px;
  color: {COLORS["slate"]};
  font-size: 12px;
}}
.est-value {{
  font-size: 24px;
  font-weight: 750;
}}
.est-unit {{
  margin-left: 3px;
  font-size: 12px;
  font-weight: 500;
}}
.est-blue {{ color: {COLORS["blue"]}; }}
.est-red {{ color: #dc2626; }}
.est-orange {{ color: {COLORS["orange_dark"]}; }}
.chart-card {{
  padding: 17px 0 5px;
  margin-bottom: 23px;
  overflow: hidden;
}}
.chart-title {{
  padding-left: 10px;
  margin: 0 18px 8px;
  border-left: 3px solid {COLORS["blue"]};
  color: {COLORS["text"]};
  font-size: 15px;
  font-weight: 650;
}}
/* =========================================================
   纯 CSS 图表：柱状图/条形图统一组件
   ========================================================= */
.chart-css {{
  width: 100%;
  padding: 8px 18px 8px;
}}
.css-chart {{ width: 100%; }}
.cssbar-row,
.cap-row {{
  display: grid;
  grid-template-columns: 74px minmax(0, 1fr);
  align-items: center;
  gap: 12px;
  min-height: 48px;
  margin: 5px 0;
}}
.cssbar-label {{
  color: #1e293b;
  font-size: 12px;
  font-weight: 700;
  white-space: nowrap;
  text-align: right;
}}
.health-track {{
  position: relative;
  height: 28px;
  border-radius: 6px;
  overflow: visible;
  background: linear-gradient(to right,
    #fee2e2 0%, #fee2e2 60%,
    #fef3c7 60%, #fef3c7 80%,
    #dcfce7 80%, #dcfce7 100%);
  box-shadow: inset 0 0 0 1px rgba(148,163,184,.12);
}}
.health-track::after {{
  content: "";
  position: absolute;
  top: -4px;
  bottom: -4px;
  left: 80%;
  border-left: 1px dashed #64748b;
  opacity: .9;
  pointer-events: none;
}}
.health-fill {{
  position: absolute;
  left: 0;
  top: 4px;
  height: 20px;
  border-radius: 5px;
  min-width: 2px;
  box-shadow: 0 1px 2px rgba(15,23,42,.10);
}}
.health-fill::after {{
  content: '';
  position: absolute;
  right: -1px;
  top: 0;
  width: 2px;
  height: 20px;
  background: rgba(255,255,255,.75);
}}
.health-value {{
  position: absolute;
  top: 50%;
  transform: translate(-50%, -50%);
  font-size: 12px;
  font-weight: 750;
  color: #fff;
  white-space: nowrap;
  text-shadow: 0 1px 2px rgba(15,23,42,.24);
  pointer-events: none;
}}
.health-fill[style*="width:3" i] + .health-value,
.health-fill[style*="width:2" i] + .health-value,
.health-fill[style*="width:1" i] + .health-value {{ color: #334155; text-shadow: none; }}
.health-axis-wrap {{
  margin-left: 86px;
  position: relative;
  height: 34px;
}}
.cssbar-axis {{
  display: grid;
  grid-template-columns: repeat(6, 1fr);
  color: #94a3b8;
  font-size: 10px;
}}
.cssbar-axis span {{ text-align: left; }}
.cssbar-axis span:last-child {{ text-align: right; }}
.health-zone-caption {{
  position: absolute;
  left: 0;
  right: 0;
  top: 15px;
  display: grid;
  grid-template-columns: 60fr 20fr 20fr;
  font-size: 9px;
  font-weight: 600;
  pointer-events: none;
}}
.health-zone-caption span {{ text-align: center; }}
.health-zone-caption .bad {{ color: #b91c1c; }}
.health-zone-caption .warn {{ color: #b45309; }}
.health-zone-caption .good {{ color: #047857; }}
.health-reference {{
  margin-left: calc(74px + 12px + 80% * 0);
  text-align: right;
  color: #64748b;
  font-size: 9px;
  margin-top: 3px;
}}
.health-reference span {{
  display: inline-block;
  width: 16px;
  border-top: 1px dashed #64748b;
  vertical-align: middle;
  margin-right: 4px;
}}
.cssbar-xlabel {{
  margin-top: 5px;
  text-align: center;
  font-size: 10px;
  color: #64748b;
}}

.capacity-chart {{ padding-top: 2px; }}
.capacity-head {{
  display: flex;
  justify-content: flex-end;
  gap: 18px;
  margin-bottom: 5px;
  color: #64748b;
  font-size: 10px;
}}
.cap-main {{ min-width: 0; }}
.cap-track {{
  position: relative;
  height: 28px;
  background: #dbeafe;
  border-radius: 5px;
  overflow: hidden;
  box-shadow: inset 0 0 0 1px #bfdbfe;
}}
.cap-fill {{
  position: absolute;
  left: 0;
  top: 0;
  bottom: 0;
  min-width: 34px;
  border-radius: 5px;
  display: flex;
  align-items: center;
  justify-content: flex-end;
  padding-right: 8px;
  box-shadow: 0 1px 2px rgba(15,23,42,.10);
}}
.cap-fill-value {{
  color: white;
  font-size: 11px;
  font-weight: 750;
  white-space: nowrap;
  text-shadow: 0 1px 2px rgba(15,23,42,.22);
}}
.cap-meta {{
  margin-top: 3px;
  display: flex;
  justify-content: flex-end;
  gap: 10px;
  color: #64748b;
  font-size: 9.5px;
  white-space: nowrap;
}}
.capacity-axis {{
  margin-left: 86px;
  display: flex;
  justify-content: space-between;
  color: #94a3b8;
  font-size: 9px;
  margin-top: 8px;
}}

.usage-summary {{
  display: flex;
  align-items: center;
  gap: 16px;
  flex-wrap: wrap;
  color: #64748b;
  font-size: 10px;
  margin-bottom: 8px;
}}
.usage-summary b {{ margin-left: auto; color: #475569; }}
.legend-dot {{
  display: inline-block;
  width: 8px;
  height: 8px;
  border-radius: 2px;
  margin-right: 4px;
  vertical-align: -1px;
  background: #94a3b8;
}}
.legend-dot.bat {{ background: #2563eb; }}
.legend-dot.ac {{ background: #f97316; }}
.month-chart-area {{
  width: 100%;
  height: 260px;
  display: flex;
  align-items: stretch;
  gap: 4px;
  padding: 10px 6px 22px;
  border-bottom: 1px solid #e2e8f0;
  background: repeating-linear-gradient(
    to top,
    transparent 0,
    transparent 49px,
    #f1f5f9 50px
  );
}}
.month-bar-item {{
  flex: 1 1 0;
  min-width: 0;
  height: 100%;
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  align-items: center;
  position: relative;
}}
.month-bar-stack {{
  width: 100%;
  max-width: 16px;
  height: calc(100% - 16px);
  display: flex;
  flex-direction: column-reverse;
  justify-content: flex-start;
  background: transparent;
  border-radius: 3px 3px 0 0;
  overflow: hidden;
}}
.month-seg {{ width: 100%; min-height: 0; }}
.month-seg.bat {{ background: #2563eb; }}
.month-seg.ac {{ background: #f97316; }}
.month-label {{
  position: absolute;
  bottom: -4px;
  transform: rotate(-45deg);
  transform-origin: right top;
  color: #94a3b8;
  font-size: 8px;
  white-space: nowrap;
}}

.activity-topbar {{
  display: flex;
  gap: 18px;
  flex-wrap: wrap;
  padding: 7px 10px;
  margin-bottom: 12px;
  border-radius: 8px;
  background: #f8fafc;
  color: #64748b;
  font-size: 10px;
}}
.activity-topbar b {{ color: #334155; }}
.activity-section-title,
.activity-subtitle {{
  font-size: 12px;
  font-weight: 700;
  color: #0f172a;
  margin: 5px 0 9px;
}}
.source-range-note {{
  font-size: 9.5px;
  font-weight: 400;
  color: #94a3b8;
}}
.event-list {{ display: grid; gap: 9px; }}
.event-row {{
  display: grid;
  grid-template-columns: 92px minmax(100px, 1fr) 58px 150px;
  align-items: center;
  gap: 8px;
  min-height: 25px;
}}
.event-time {{ font-size: 10px; color: #475569; white-space: nowrap; }}
.event-track {{
  height: 16px;
  background: #eff6ff;
  border-radius: 4px;
  overflow: hidden;
}}
.event-fill {{
  height: 100%;
  background: #2563eb;
  border-radius: 4px;
}}
.event-minutes {{
  text-align: right;
  font-size: 10px;
  color: #1d4ed8;
  font-weight: 700;
  white-space: nowrap;
}}
.event-detail {{
  font-size: 9.5px;
  color: #64748b;
  white-space: nowrap;
}}
.activity-foot-grid {{
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px;
  margin-top: 18px;
}}
.activity-subcard {{
  min-width: 0;
  padding-top: 11px;
  border-top: 1px solid #e2e8f0;
}}
.state-note {{
  margin-bottom: 7px;
  color: #94a3b8;
  font-size: 8.5px;
  line-height: 1.5;
}}
.simple-stat-row {{
  display: grid;
  grid-template-columns: 1fr auto auto;
  gap: 9px;
  align-items: center;
  padding: 6px 0;
  border-bottom: 1px dashed #edf2f7;
  font-size: 10px;
}}
.simple-stat-row b {{ color: #334155; }}
.simple-stat-row em {{ color: #94a3b8; font-style: normal; }}
.source-track {{
  display: flex;
  height: 24px;
  border-radius: 5px;
  overflow: hidden;
  background: #f1f5f9;
  box-shadow: inset 0 0 0 1px #e2e8f0;
  margin: 9px 0;
}}
.source-seg {{ height: 100%; }}
.source-legend {{
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  color: #64748b;
  font-size: 9px;
}}
.empty-note {{ padding: 12px 0; color: #94a3b8; font-size: 10px; }}

@media (max-width: 700px) {{
  .event-row {{ grid-template-columns: 76px minmax(80px, 1fr) 52px; }}
  .event-detail {{ grid-column: 2 / -1; }}
  .activity-foot-grid {{ grid-template-columns: 1fr; }}
  .cap-meta {{ justify-content: flex-start; overflow: hidden; }}
  .system-item {{ flex-direction: column; gap: 3px; }}
  .system-item b {{ text-align: left; }}
}}

.chart-svg {{
  width: 100%;
  overflow-x: auto;
}}
.chart-svg svg {{
  display: block;
  width: 100%;
  height: auto;
  min-height: 80px;
}}
.system-grid {{
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 0 24px;
}}
.system-item {{
  display: flex;
  justify-content: space-between;
  gap: 15px;
  padding: 8px 0;
  border-bottom: 1px dashed #edf2f7;
  font-size: 12px;
}}
.system-item span {{ color: #64748b; }}
.system-item b {{ color: #334155; font-weight: 600; text-align: right; }}
.advice-card {{
  background: linear-gradient(135deg, #eff6ff, #e0ecff);
  border: 1px solid #dbeafe;
  border-radius: 17px;
  padding: 21px;
  box-shadow: 0 7px 20px rgba(15,23,42,.045);
}}
.advice-title {{
  color: #1e40af;
  font-size: 15px;
  font-weight: 750;
  margin-bottom: 9px;
}}
.advice-summary {{
  color: #1e3a8a;
  font-size: 15px;
  line-height: 1.65;
  font-weight: 650;
  margin-bottom: 10px;
}}
.advice-card ul {{
  margin: 0;
  padding-left: 20px;
  color: #334155;
  font-size: 13px;
  line-height: 1.9;
}}
.footer {{
  margin-top: 25px;
  text-align: center;
  color: #94a3b8;
  font-size: 11px;
  line-height: 1.8;
}}
@media (max-width: 800px) {{
  .stat-grid {{ grid-template-columns: repeat(2, 1fr); }}
  .battery-body {{ grid-template-columns: 1fr; }}
  .kv-grid {{ grid-template-columns: 1fr; }}
  .est-grid {{ grid-template-columns: 1fr; }}
  .system-grid {{ grid-template-columns: 1fr; }}
}}
@media (max-width: 520px) {{
  body {{ padding-left: 10px; padding-right: 10px; }}
  .stat-grid {{ grid-template-columns: 1fr 1fr; gap: 8px; }}
  .stat-card {{ padding: 13px; }}
  .stat-value {{ font-size: 19px; }}
  .card, .battery-card, .advice-card {{ padding: 15px; }}
}}
</style>
</head>
<body>
<div class="container">

  <div class="header">
    <h1>🔋 笔记本电池深度分析报告</h1>
    <div class="meta">
      生成于 {now} · 检测到 {len(batteries)} 块电池
    </div>
  </div>

  {system_html}
  {overview}
  {"".join(battery_cards)}
  {current_html}

  <div class="card">
    <h2>📈 可视化分析</h2>
    {"<p class='explain'>长期使用记录已按月度聚合展示；容量历史保留原始记录并增加趋势线。</p>" if chart_files else ""}
    {chart_cards if chart_cards else "<p class='explain'>暂无可视化数据。</p>"}
  </div>

  <div class="advice-card">
    <div class="advice-title">💡 综合评估与建议</div>
    <div class="advice-summary">{html.escape(summary)}</div>
    <ul>{advice_html}</ul>
  </div>

  <div class="footer">
    数据来源：Windows powercfg /batteryreport<br>
    健康度 = 满充容量 ÷ 设计容量 × 100%
  </div>
</div>
</body>
</html>
"""

    report_path = report_dir / "笔记本电池深度报告.html"
    report_path.write_text(page, encoding="utf-8")
    return str(report_path)


# ============================================================
# 主流程
# ============================================================

def create_output_dir() -> Path:
    report_dir = Path(tempfile.gettempdir()) / "BatteryReport"
    report_dir.mkdir(parents=True, exist_ok=True)

    for item in report_dir.iterdir():
        if item.is_file() and item.suffix.lower() in {".png", ".html"}:
            try:
                item.unlink()
            except OSError:
                pass

    return report_dir


def print_summary(data: dict) -> None:
    batteries = data["batteries"]

    print("\n检测结果")
    print("-" * 65)
    print(f"已安装电池：{len(batteries)} 块")
    print(f"最近使用：{len(data['recent_usage'])} 条")
    print(f"放电事件：{len(data['battery_usage'])} 条")
    print(f"使用历史：{len(data['usage_history'])} 条")
    print(f"容量历史：{len(data['capacity_history'])} 条")
    print(f"续航估算历史：{len(data['life_estimates'])} 条")
    print()

    for b in batteries:
        print(
            f"  - {b.get('name', b.get('slot'))}: "
            f"健康度 {b.get('health')}% | "
            f"设计 {b.get('design_capacity')} | "
            f"满充 {b.get('full_charge_capacity')} | "
            f"循环 {b.get('cycle_count')} 次"
        )


def build_report(input_path: Optional[str] = None) -> str:
    report_dir = create_output_dir()
    print(f"输出目录：{report_dir}")

    source = prepare_report_source(input_path)

    print("\n解析 HTML...")
    data = parse_report(source)
    if not input_path:
        data["live_battery_status"] = query_live_battery_status()
    print_summary(data)

    if not data["batteries"]:
        raise RuntimeError("未检测到 Installed batteries 数据。")

    print("\n生成图表...")
    chart_jobs = [
        ("chart_health_compare", chart_health_compare, (data["batteries"], report_dir)),
        ("chart_capacity_compare", chart_capacity_compare, (data["batteries"], report_dir)),
        ("chart_capacity_history", chart_capacity_history, (data["capacity_history"], report_dir)),
        ("chart_life_estimates", chart_life_estimates, (data["life_estimates"], report_dir)),
        ("chart_usage_history", chart_usage_history, (data["usage_history"], report_dir)),
        ("chart_recent_activity", chart_recent_activity, (
            data["recent_usage"],
            data["battery_usage"],
            report_dir,
        )),
    ]

    chart_files = []

    for name, func, args in chart_jobs:
        try:
            result = func(*args)
            if result:
                chart_files.append(result)
                print(f"  ✓ {result[0]}")
            else:
                print(f"  - {name}: 数据不足，跳过")
        except Exception as exc:
            print(f"  ✗ {name}: {exc}")

    report_path = generate_output_html(
        data,
        chart_files,
        report_dir,
    )

    print(f"\n✅ 报告已生成：{report_path}")

    # 使用已有文件调试时，不删除用户文件；
    # 自动生成的临时 battery report 则在成功解析后删除。
    if not input_path:
        try:
            Path(source).unlink()
        except OSError:
            pass

    try:
        open_report_in_existing_browser(report_path)
    except Exception:
        pass

    return report_path


def open_report_in_existing_browser(report_path: str) -> bool:
    """把文件交给 Windows Shell，由当前默认浏览器自己决定复用已有窗口/标签页。"""
    path = Path(report_path).resolve()
    try:
        if os.name == "nt":
            # os.startfile 不直接创建浏览器进程，而是交给 Windows 文件关联处理。
            os.startfile(str(path))
            return True
    except OSError:
        pass

    try:
        controller = webbrowser.get()
        return bool(controller.open_new_tab(path.as_uri()))
    except Exception:
        return bool(webbrowser.open(path.as_uri()))


def parse_args():
    parser = argparse.ArgumentParser(
        description="Windows battery-report 深度分析工具（优化版）"
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="已有 battery-report.html；不填写则自动调用 powercfg",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="生成后不自动打开 HTML",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    try:
        report_dir = create_output_dir()
        print(f"输出目录：{report_dir}")

        source = prepare_report_source(args.input)

        print("\n解析 HTML...")
        data = parse_report(source)
        if not args.input:
            data["live_battery_status"] = query_live_battery_status()
            if data["live_battery_status"]:
                print("✓ 已读取实时电池状态（Win32_Battery）")
            else:
                print("- 未读取到实时电池状态，将使用 battery-report 最近记录进行判断")
        print_summary(data)

        if not data["batteries"]:
            raise RuntimeError("未检测到 Installed batteries 数据。")

        print("\n生成图表...")
        chart_jobs = [
            ("chart_health_compare", chart_health_compare, (data["batteries"], report_dir)),
            ("chart_capacity_compare", chart_capacity_compare, (data["batteries"], report_dir)),
            ("chart_capacity_history", chart_capacity_history, (data["capacity_history"], report_dir)),
            ("chart_life_estimates", chart_life_estimates, (data["life_estimates"], report_dir)),
            ("chart_usage_history", chart_usage_history, (data["usage_history"], report_dir)),
            ("chart_recent_activity", chart_recent_activity, (
                data["recent_usage"],
                data["battery_usage"],
                report_dir,
            )),
        ]

        chart_files = []
        for name, func, args_tuple in chart_jobs:
            try:
                result = func(*args_tuple)
                if result:
                    chart_files.append(result)
                    print(f"  ✓ {result[0]}")
                else:
                    print(f"  - {name}: 数据不足，跳过")
            except Exception as exc:
                print(f"  ✗ {name}: {exc}")

        report_path = generate_output_html(
            data,
            chart_files,
            report_dir,
        )

        print("\n" + "=" * 65)
        print(f"✅ 分析完成：{report_path}")
        print("=" * 65)

        if not args.input:
            try:
                Path(source).unlink()
            except OSError:
                pass

        if not args.no_open:
            if not open_report_in_existing_browser(report_path):
                print("提示：无法自动打开报告，请手动打开 HTML 文件。")

        input("\n按回车键退出...")

    except KeyboardInterrupt:
        print("\n用户取消。")
    except Exception as exc:
        print(f"\n❌ 运行失败：{exc}")
        input("按回车键退出...")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
