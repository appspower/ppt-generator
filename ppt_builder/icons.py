"""아이콘 시스템 — Segoe MDL2 Assets + Segoe Fluent Icons 기반.

실제 벡터 아이콘을 폰트 글리프로 렌더. 크기 조절 자유, OS 의존성 없음 (Windows).
PwC Component Library의 선형 아이콘 스타일을 재현.

사용 예:
    from ppt_builder.icons import draw_icon, ICON_MAP
    draw_icon(canvas, "database", x=1.0, y=2.0, size=24, color="accent", region=r)
"""

from __future__ import annotations
from ppt_builder.primitives import Canvas, Region

# Segoe MDL2 Assets 아이콘 매핑
# https://learn.microsoft.com/en-us/windows/apps/design/style/segoe-ui-symbol-font
ICON_FONT = "Segoe MDL2 Assets"

ICON_MAP = {
    # 비즈니스/일반
    "database":     "\uE968",  # Database
    "cloud":        "\uE753",  # Cloud
    "globe":        "\uE774",  # Globe
    "lock":         "\uE72E",  # Lock
    "unlock":       "\uE785",  # Unlock
    "settings":     "\uE713",  # Settings (gear)
    "person":       "\uE77B",  # Contact
    "people":       "\uE716",  # People
    "building":     "\uE731",  # Library (building)
    "calendar":     "\uE787",  # Calendar
    "clock":        "\uE823",  # Clock
    "mail":         "\uE715",  # Mail
    "phone":        "\uE717",  # Phone
    "search":       "\uE721",  # Search
    "filter":       "\uE71C",  # Filter
    "sort":         "\uE8CB",  # Sort

    # 데이터/차트
    "chart_bar":    "\uE9D2",  # BarChartHorizontal
    "chart_line":   "\uE9D9",  # LineChart (reserved)
    "chart_pie":    "\uEB05",  # PieDouble
    "dashboard":    "\uF246",  # ViewDashboard
    "analytics":    "\uE9D2",  # Analytics

    # 상태/액션
    "check":        "\uE73E",  # CheckMark
    "check_circle": "\uE73E",  # CheckMark (circle variant via styling)
    "warning":      "\uE7BA",  # Warning
    "error":        "\uE783",  # Error
    "info":         "\uE946",  # Info
    "star":         "\uE734",  # FavoriteStar
    "star_fill":    "\uE735",  # FavoriteStarFill
    "flag":         "\uE7C1",  # Flag
    "pin":          "\uE718",  # Pin

    # 화살표/방향
    "arrow_right":  "\uE72A",  # Forward
    "arrow_left":   "\uE72B",  # Back
    "arrow_up":     "\uE74A",  # Up
    "arrow_down":   "\uE74B",  # Down
    "chevron_right":"\uE76C",  # ChevronRight
    "chevron_down": "\uE70D",  # ChevronDown

    # 기술/시스템
    "code":         "\uE943",  # Code
    "server":       "\uE839",  # HardDrive
    "network":      "\uE968",  # NetworkTower
    "shield":       "\uE72E",  # Shield
    "key":          "\uE8D7",  # Permissions
    "robot":        "\uE99A",  # Robot

    # 문서/콘텐츠
    "document":     "\uE8A5",  # Document
    "folder":       "\uE8B7",  # FolderOpen
    "clipboard":    "\uE77F",  # Paste
    "edit":         "\uE70F",  # Edit
    "save":         "\uE74E",  # Save
    "delete":       "\uE74D",  # Delete

    # 프로세스/흐름
    "refresh":      "\uE72C",  # Refresh
    "sync":         "\uE895",  # Sync
    "play":         "\uE768",  # Play
    "pause":        "\uE769",  # Pause
    "stop":         "\uE71A",  # Stop
    "fast_forward": "\uEB9D",  # FastForward

    # 비즈니스 프로세스
    "handshake":    "\uE8FA",  # Relationship (handshake)
    "money":        "\uE8C6",  # Money (currency)
    "certificate":  "\uE734",  # Certificate
    "growth":       "\uE8A9",  # Trending up
    "target":       "\uE7C1",  # Target (flag as proxy)
    "light_bulb":   "\uEA80",  # Light bulb (idea)
    "rocket":       "\uE7C8",  # Launch (rocket)
}


def draw_icon(
    c: Canvas,
    icon_name: str,
    *,
    x: float,
    y: float,
    size: float = 24,
    color: str = "accent",
    region: Region | None = None,
) -> None:
    """MDL2 아이콘을 지정 위치에 렌더.

    Args:
        icon_name: ICON_MAP 키 (e.g. "database", "check", "chart_bar")
        x, y: 위치 (인치, region 있으면 상대)
        size: 폰트 크기 pt
        color: 색상 별칭 또는 hex
    """
    glyph = ICON_MAP.get(icon_name, ICON_MAP.get("info", "\uE946"))
    # 아이콘 크기에 맞는 박스
    box_size = size * 0.02  # pt → 인치 근사
    c.text(
        glyph,
        x=x, y=y, w=box_size, h=box_size,
        size=size, color=color, font=ICON_FONT,
        align="center", anchor="middle",
        region=region,
    )


def draw_icon_with_label(
    c: Canvas,
    icon_name: str,
    label: str,
    *,
    x: float,
    y: float,
    w: float = 1.5,
    h: float = 0.8,
    icon_size: float = 20,
    label_size: float = 8,
    color: str = "accent",
    label_color: str = "grey_900",
    layout: str = "vertical",  # "vertical" (아이콘 위, 텍스트 아래) / "horizontal" (좌 아이콘, 우 텍스트)
    region: Region | None = None,
) -> None:
    """아이콘 + 라벨 조합."""
    if layout == "vertical":
        icon_h = h * 0.55
        c.text(
            ICON_MAP.get(icon_name, "\uE946"),
            x=x, y=y, w=w, h=icon_h,
            size=icon_size, color=color, font=ICON_FONT,
            align="center", anchor="middle",
            region=region,
        )
        c.text(
            label,
            x=x, y=y + icon_h, w=w, h=h - icon_h,
            size=label_size, bold=True, color=label_color,
            align="center", anchor="top",
            region=region,
        )
    else:  # horizontal
        icon_w = 0.45
        c.text(
            ICON_MAP.get(icon_name, "\uE946"),
            x=x, y=y, w=icon_w, h=h,
            size=icon_size, color=color, font=ICON_FONT,
            align="center", anchor="middle",
            region=region,
        )
        c.text(
            label,
            x=x + icon_w + 0.08, y=y, w=w - icon_w - 0.08, h=h,
            size=label_size, bold=True, color=label_color,
            anchor="middle",
            region=region,
        )
