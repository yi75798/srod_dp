#!/usr/bin/python
# -*- encoding: utf-8 -*-
# File    :   app.py
# Time    :   2026/08/13 13:55:55
# Author  :   Hsu, Liang-Yi 
# Email:   yi75798@gmail.com
# Description :

import streamlit as st
import streamlit.components.v1 as components

import pandas as pd
import numpy as np

from io import BytesIO
from html import escape


# =========================================================
# 網頁基本設定
# =========================================================
st.set_page_config(
    page_title="甘特圖工具",
    page_icon="📅",
    layout="wide"
)

st.title("📅 甘特圖製作工具")

st.write(
    """
    上傳 Excel 後，系統會將每個工作的完成期限顯示在時間軸上。

    你可以建立工作之間的連動關係、計算日曆天或工作天、
    設定休假日，並將完整專案儲存成 Excel，
    或匯出成不需要重新上傳資料的靜態 HTML。
    """
)


# =========================================================
# Session State 初始化
# =========================================================
if "links" not in st.session_state:
    st.session_state.links = []

if "holidays" not in st.session_state:
    st.session_state.holidays = []

if "day_mode" not in st.session_state:
    st.session_state.day_mode = "日曆天"

if "workdays" not in st.session_state:
    st.session_state.workdays = {
        "monday": True,
        "tuesday": True,
        "wednesday": True,
        "thursday": True,
        "friday": True,
        "saturday": False,
        "sunday": False
    }

if "loaded_project_name" not in st.session_state:
    st.session_state.loaded_project_name = None


# =========================================================
# 日期格式
# =========================================================
def format_date(date):
    """
    2026-08-13 -> 26/8/13
    """
    date = pd.Timestamp(date)

    return (
        f"{date.strftime('%y')}/"
        f"{date.month}/"
        f"{date.day}"
    )


def format_month(date):
    """
    2026-08 -> 26/8
    """
    date = pd.Timestamp(date)

    return (
        f"{date.strftime('%y')}/"
        f"{date.month}"
    )


def format_short_date(date):
    """
    2026-08-13 -> 8/13
    """
    date = pd.Timestamp(date)

    return (
        f"{date.month}/"
        f"{date.day}"
    )


# =========================================================
# 計算日曆天
# =========================================================
def calculate_calendar_days(
    start_date,
    end_date
):
    return (
        end_date - start_date
    ).days + 1


# =========================================================
# 計算工作天
# =========================================================
def calculate_workdays(
    start_date,
    end_date,
    weekmask,
    holidays
):
    start_np = np.datetime64(
        start_date.date()
    )

    end_np = (
        np.datetime64(
            end_date.date()
        )
        +
        np.timedelta64(
            1,
            "D"
        )
    )

    holiday_array = np.array(
        [
            np.datetime64(
                holiday.strftime(
                    "%Y-%m-%d"
                )
            )
            for holiday in holidays
        ],
        dtype="datetime64[D]"
    )

    return int(
        np.busday_count(
            start_np,
            end_np,
            weekmask=weekmask,
            holidays=holiday_array
        )
    )


# =========================================================
# 建立 weekmask
# =========================================================
def create_weekmask(workdays):
    keys = [
        "monday",
        "tuesday",
        "wednesday",
        "thursday",
        "friday",
        "saturday",
        "sunday"
    ]

    return "".join(
        "1" if workdays[key] else "0"
        for key in keys
    )


# =========================================================
# 匯出完整專案 Excel
# =========================================================
def create_project_excel(
    task_df,
    links,
    holidays,
    day_mode,
    workdays
):
    output = BytesIO()

    task_export = task_df.copy()

    task_export[
        "完成期限"
    ] = pd.to_datetime(
        task_export[
            "完成期限"
        ]
    )

    links_export = pd.DataFrame(
        links
    )

    if links_export.empty:
        links_export = pd.DataFrame(
            columns=[
                "start",
                "end"
            ]
        )

    holidays_export = pd.DataFrame(
        {
            "日期": [
                holiday.strftime(
                    "%Y-%m-%d"
                )
                for holiday in holidays
            ]
        }
    )

    settings_export = pd.DataFrame(
        {
            "設定": [
                "day_mode",
                "monday",
                "tuesday",
                "wednesday",
                "thursday",
                "friday",
                "saturday",
                "sunday"
            ],
            "值": [
                day_mode,
                workdays["monday"],
                workdays["tuesday"],
                workdays["wednesday"],
                workdays["thursday"],
                workdays["friday"],
                workdays["saturday"],
                workdays["sunday"]
            ]
        }
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        task_export.to_excel(
            writer,
            sheet_name="工作項目",
            index=False
        )

        links_export.to_excel(
            writer,
            sheet_name="連動關係",
            index=False
        )

        holidays_export.to_excel(
            writer,
            sheet_name="排除日期",
            index=False
        )

        settings_export.to_excel(
            writer,
            sheet_name="系統設定",
            index=False
        )

    output.seek(0)

    return output


# =========================================================
# 載入完整專案設定
# =========================================================
def load_project_settings(
    uploaded_file,
    task_list
):
    excel_file = pd.ExcelFile(
        uploaded_file
    )

    sheet_names = (
        excel_file.sheet_names
    )

    loaded_links = []
    loaded_holidays = []
    loaded_day_mode = "日曆天"

    loaded_workdays = {
        "monday": True,
        "tuesday": True,
        "wednesday": True,
        "thursday": True,
        "friday": True,
        "saturday": False,
        "sunday": False
    }

    invalid_links = []

    # -----------------------------------------------------
    # 連動
    # -----------------------------------------------------
    if "連動關係" in sheet_names:

        links_df = pd.read_excel(
            uploaded_file,
            sheet_name="連動關係"
        )

        if (
            "start" in links_df.columns
            and
            "end" in links_df.columns
        ):

            for _, row in links_df.iterrows():

                start_name = row["start"]
                end_name = row["end"]

                if (
                    pd.isna(start_name)
                    or
                    pd.isna(end_name)
                ):
                    continue

                start_name = str(
                    start_name
                )

                end_name = str(
                    end_name
                )

                if (
                    start_name in task_list
                    and
                    end_name in task_list
                ):
                    loaded_links.append(
                        {
                            "start": start_name,
                            "end": end_name
                        }
                    )

                else:
                    invalid_links.append(
                        {
                            "start": start_name,
                            "end": end_name
                        }
                    )

    # -----------------------------------------------------
    # 排除日期
    # -----------------------------------------------------
    if "排除日期" in sheet_names:

        holidays_df = pd.read_excel(
            uploaded_file,
            sheet_name="排除日期"
        )

        if "日期" in holidays_df.columns:

            for value in holidays_df[
                "日期"
            ]:

                if pd.notna(value):

                    date = pd.to_datetime(
                        value,
                        errors="coerce"
                    )

                    if pd.notna(date):

                        loaded_holidays.append(
                            pd.Timestamp(
                                date
                            ).normalize()
                        )

    # -----------------------------------------------------
    # 系統設定
    # -----------------------------------------------------
    if "系統設定" in sheet_names:

        settings_df = pd.read_excel(
            uploaded_file,
            sheet_name="系統設定"
        )

        if (
            "設定" in settings_df.columns
            and
            "值" in settings_df.columns
        ):

            settings_dict = dict(
                zip(
                    settings_df[
                        "設定"
                    ],
                    settings_df[
                        "值"
                    ]
                )
            )

            if (
                "day_mode"
                in settings_dict
            ):

                value = str(
                    settings_dict[
                        "day_mode"
                    ]
                )

                if value in [
                    "日曆天",
                    "工作天"
                ]:
                    loaded_day_mode = (
                        value
                    )

            for key in (
                loaded_workdays.keys()
            ):

                if key in settings_dict:

                    value = (
                        settings_dict[
                            key
                        ]
                    )

                    if isinstance(
                        value,
                        (
                            bool,
                            np.bool_
                        )
                    ):
                        loaded_workdays[
                            key
                        ] = bool(
                            value
                        )

                    else:
                        loaded_workdays[
                            key
                        ] = (
                            str(
                                value
                            ).lower()
                            in [
                                "true",
                                "1",
                                "yes"
                            ]
                        )

    return (
        loaded_links,
        loaded_holidays,
        loaded_day_mode,
        loaded_workdays,
        invalid_links
    )


# =========================================================
# 建立甘特圖 HTML
# =========================================================
def create_gantt_html(
    gantt_df,
    links,
    day_mode,
    weekmask,
    holidays,
    project_title,
    static_mode=False
):

    # =====================================================
    # 外觀參數
    # =====================================================
    DAY_WIDTH = 28

    ROW_HEIGHT = 70

    HEADER_HEIGHT = 75

    FOOTER_HEIGHT = 60

    LABEL_WIDTH = 250

    LABEL_MIN_WIDTH = 150

    LABEL_MAX_WIDTH = 650

    RESIZER_WIDTH = 7

    SCROLLBAR_HEIGHT = 18

    DATE_PADDING = 7


    server_today = (
        pd.Timestamp.today()
        .normalize()
    )


    # =====================================================
    # 日期範圍
    # =====================================================
    all_dates = (
        gantt_df[
            "完成期限"
        ]
        .dropna()
    )

    min_task_date = (
        all_dates.min()
    )

    max_task_date = (
        all_dates.max()
    )


    if static_mode:

        start_date = (
            min_task_date
            -
            pd.Timedelta(
                days=60
            )
        ).normalize()

        end_date = (
            max_task_date
            +
            pd.Timedelta(
                days=365
            )
        ).normalize()

    else:

        min_date = min(
            min_task_date,
            server_today
        )

        max_date = max(
            max_task_date,
            server_today
        )

        start_date = (
            min_date
            -
            pd.Timedelta(
                days=DATE_PADDING
            )
        ).normalize()

        end_date = (
            max_date
            +
            pd.Timedelta(
                days=DATE_PADDING
            )
        ).normalize()


    # =====================================================
    # 時間軸尺寸
    # =====================================================
    total_days = (
        end_date
        -
        start_date
    ).days + 1

    timeline_width = (
        total_days
        *
        DAY_WIDTH
    )

    task_area_height = (
        len(gantt_df)
        *
        ROW_HEIGHT
    )

    footer_start_y = (
        HEADER_HEIGHT
        +
        task_area_height
    )

    svg_height = (
        HEADER_HEIGHT
        +
        task_area_height
        +
        FOOTER_HEIGHT
    )


    # =====================================================
    # 日期轉 X
    # =====================================================
    def date_to_x(date):

        date = (
            pd.Timestamp(date)
            .normalize()
        )

        day_index = (
            date
            -
            start_date
        ).days

        return (
            day_index
            *
            DAY_WIDTH
            +
            DAY_WIDTH / 2
        )


    # =====================================================
    # 工作項目 Y
    # =====================================================
    task_positions = {}

    for index, row in (
        gantt_df
        .reset_index(
            drop=True
        )
        .iterrows()
    ):

        task_positions[
            row["工作項目"]
        ] = (
            HEADER_HEIGHT
            +
            index
            *
            ROW_HEIGHT
            +
            ROW_HEIGHT / 2
        )


    # =====================================================
    # 左側工作名稱
    # =====================================================
    labels_html = """
    <div class="task-label-header">
        工作項目
    </div>
    """

    for _, row in gantt_df.iterrows():

        task_name = escape(
            str(
                row["工作項目"]
            )
        )

        labels_html += f"""
        <div
            class="task-label-row"
            title="{task_name}"
        >
            <div class="task-label-text">
                {task_name}
            </div>
        </div>
        """

    labels_html += f"""
    <div class="task-label-footer">
        日期
    </div>
    """


    # =====================================================
    # SVG
    # =====================================================
    svg_parts = []


    # -----------------------------------------------------
    # 背景
    # -----------------------------------------------------
    svg_parts.append(
        f"""
        <rect
            x="0"
            y="0"
            width="{timeline_width}"
            height="{svg_height}"
            fill="#080808"
        />
        """
    )


    # =====================================================
    # 每日背景 / 格線
    # =====================================================
    for i in range(total_days):

        current_date = (
            start_date
            +
            pd.Timedelta(
                days=i
            )
        )

        x = (
            i
            *
            DAY_WIDTH
        )

        # 週末
        if current_date.weekday() >= 5:

            svg_parts.append(
                f"""
                <rect
                    x="{x}"
                    y="0"
                    width="{DAY_WIDTH}"
                    height="{svg_height}"
                    fill="#101010"
                />
                """
            )

        # 每日格線
        svg_parts.append(
            f"""
            <line
                x1="{x}"
                y1="0"
                x2="{x}"
                y2="{svg_height}"
                stroke="#242424"
                stroke-width="1"
            />
            """
        )


    # =====================================================
    # 上下月份
    # =====================================================
    previous_month = None

    for i in range(total_days):

        current_date = (
            start_date
            +
            pd.Timedelta(
                days=i
            )
        )

        month_key = (
            current_date.year,
            current_date.month
        )

        if month_key != previous_month:

            x = (
                i
                *
                DAY_WIDTH
            )

            # 月份分隔線
            svg_parts.append(
                f"""
                <line
                    x1="{x}"
                    y1="0"
                    x2="{x}"
                    y2="{svg_height}"
                    stroke="#666666"
                    stroke-width="2"
                />
                """
            )

            # 上方月份
            svg_parts.append(
                f"""
                <text
                    x="{x + 6}"
                    y="20"
                    fill="#FFFFFF"
                    font-size="13"
                    font-weight="600"
                >
                    {format_month(current_date)}
                </text>
                """
            )

            # 下方月份
            svg_parts.append(
                f"""
                <text
                    x="{x + 6}"
                    y="{footer_start_y + 22}"
                    fill="#FFFFFF"
                    font-size="13"
                    font-weight="600"
                >
                    {format_month(current_date)}
                </text>
                """
            )

            previous_month = (
                month_key
            )


    # =====================================================
    # 上下日期刻度
    # =====================================================
    for i in range(total_days):

        current_date = (
            start_date
            +
            pd.Timedelta(
                days=i
            )
        )

        if (
            i % 7 == 0
            or
            current_date.day == 1
        ):

            x = (
                i
                *
                DAY_WIDTH
                +
                DAY_WIDTH / 2
            )

            date_text = (
                format_short_date(
                    current_date
                )
            )

            # 上方
            svg_parts.append(
                f"""
                <text
                    x="{x}"
                    y="52"
                    fill="#CFCFCF"
                    font-size="11"
                    text-anchor="middle"
                >
                    {date_text}
                </text>
                """
            )

            # 下方
            svg_parts.append(
                f"""
                <text
                    x="{x}"
                    y="{footer_start_y + 48}"
                    fill="#CFCFCF"
                    font-size="11"
                    text-anchor="middle"
                >
                    {date_text}
                </text>
                """
            )


    # =====================================================
    # 水平工作格線
    # =====================================================
    for i in range(
        len(gantt_df) + 1
    ):

        y = (
            HEADER_HEIGHT
            +
            i
            *
            ROW_HEIGHT
        )

        svg_parts.append(
            f"""
            <line
                x1="0"
                y1="{y}"
                x2="{timeline_width}"
                y2="{y}"
                stroke="#303030"
                stroke-width="1"
            />
            """
        )


    # =====================================================
    # Streamlit 預覽版今天線
    # =====================================================
    if not static_mode:

        today_x = (
            date_to_x(
                server_today
            )
        )

        today_label = (
            "今天 "
            +
            format_date(
                server_today
            )
        )

        svg_parts.append(
            f"""
            <g id="todayMarker">

                <line
                    x1="{today_x}"
                    y1="0"
                    x2="{today_x}"
                    y2="{svg_height}"
                    stroke="#FF3B30"
                    stroke-width="3"
                    stroke-dasharray="7 5"
                />

                <rect
                    x="{today_x - 46}"
                    y="4"
                    width="92"
                    height="22"
                    rx="4"
                    fill="#D32F2F"
                />

                <text
                    x="{today_x}"
                    y="19"
                    fill="#FFFFFF"
                    font-size="11"
                    font-weight="600"
                    text-anchor="middle"
                >
                    {today_label}
                </text>

            </g>
            """
        )


    # =====================================================
    # 工作連動
    # =====================================================
    for link in links:

        start_name = (
            link["start"]
        )

        end_name = (
            link["end"]
        )

        start_rows = gantt_df[
            gantt_df[
                "工作項目"
            ]
            ==
            start_name
        ]

        end_rows = gantt_df[
            gantt_df[
                "工作項目"
            ]
            ==
            end_name
        ]

        if (
            start_rows.empty
            or
            end_rows.empty
        ):
            continue

        start_task_date = (
            start_rows[
                "完成期限"
            ].iloc[0]
        )

        end_task_date = (
            end_rows[
                "完成期限"
            ].iloc[0]
        )

        start_x = (
            date_to_x(
                start_task_date
            )
        )

        end_x = (
            date_to_x(
                end_task_date
            )
        )

        start_y = (
            task_positions[
                start_name
            ]
        )

        end_y = (
            task_positions[
                end_name
            ]
        )

        calendar_days = (
            calculate_calendar_days(
                start_task_date,
                end_task_date
            )
        )

        work_days = (
            calculate_workdays(
                start_task_date,
                end_task_date,
                weekmask,
                holidays
            )
        )

        if day_mode == "日曆天":
            display_text = (
                f"{calendar_days} 天"
            )
        else:
            display_text = (
                f"{work_days} 工作天"
            )

        LINE_COLOR = (
            "#4DA3FF"
        )

        # 水平線
        svg_parts.append(
            f"""
            <line
                x1="{start_x}"
                y1="{start_y}"
                x2="{end_x}"
                y2="{start_y}"
                stroke="{LINE_COLOR}"
                stroke-width="7"
                stroke-linecap="round"
            />
            """
        )

        # 垂直線
        if start_y != end_y:

            svg_parts.append(
                f"""
                <line
                    x1="{end_x}"
                    y1="{start_y}"
                    x2="{end_x}"
                    y2="{end_y}"
                    stroke="{LINE_COLOR}"
                    stroke-width="3"
                    stroke-dasharray="5 4"
                />
                """
            )

        # 天數
        middle_x = (
            start_x
            +
            end_x
        ) / 2

        label_width = max(
            62,
            len(
                display_text
            )
            *
            15
        )

        svg_parts.append(
            f"""
            <rect
                x="{
                    middle_x
                    -
                    label_width / 2
                }"
                y="{start_y - 17}"
                width="{label_width}"
                height="26"
                rx="6"
                fill="#1B1B1B"
                stroke="#FFFFFF"
                stroke-width="1"
            />

            <text
                x="{middle_x}"
                y="{start_y + 1}"
                fill="#FFFFFF"
                font-size="12"
                font-weight="600"
                text-anchor="middle"
            >
                {escape(display_text)}
            </text>
            """
        )


    # =====================================================
    # 里程碑
    # =====================================================
    for _, row in gantt_df.iterrows():

        task_name = (
            row[
                "工作項目"
            ]
        )

        milestone_date = (
            row[
                "完成期限"
            ]
        )

        x = (
            date_to_x(
                milestone_date
            )
        )

        y = (
            task_positions[
                task_name
            ]
        )

        date_text = (
            format_date(
                milestone_date
            )
        )

        # 日期
        svg_parts.append(
            f"""
            <text
                x="{x}"
                y="{y - 18}"
                fill="#FFFFFF"
                font-size="11"
                font-weight="600"
                text-anchor="middle"
            >
                {date_text}
            </text>
            """
        )

        # 菱形
        diamond_size = 7

        points = (
            f"{x},{y - diamond_size} "
            f"{x + diamond_size},{y} "
            f"{x},{y + diamond_size} "
            f"{x - diamond_size},{y}"
        )

        svg_parts.append(
            f"""
            <polygon
                points="{points}"
                fill="#FFD54F"
                stroke="#FFFFFF"
                stroke-width="1.5"
            >
                <title>
                    {
                        escape(
                            str(
                                task_name
                            )
                        )
                    }
                    -
                    {date_text}
                </title>
            </polygon>
            """
        )


    svg_content = "".join(
        svg_parts
    )


    # =====================================================
    # 初始位置
    # =====================================================
    initial_x = (
        date_to_x(
            server_today
        )
    )

    initial_scroll = max(
        0,
        initial_x - 450
    )


    safe_title = escape(
        project_title
    )


    # =====================================================
    # 匯出時間
    # =====================================================
    export_time = (
        pd.Timestamp.now()
    )

    export_time_text = (
        f"{format_date(export_time)} "
        f"{export_time.strftime('%H:%M')}"
    )


    # =====================================================
    # BODY
    # =====================================================
    body_content = f"""

    <div class="page-header">

        <div>

            <h1>
                📊 {safe_title}
            </h1>

            {
                f'''
                <div class="update-time">
                    資料匯出時間：
                    {export_time_text}
                </div>
                '''
                if static_mode
                else
                ""
            }

        </div>

        <div class="header-actions">

            <button
                id="todayButton"
                type="button"
            >
                跳到今天
            </button>

        </div>

    </div>


    <div
        id="ganttWrapper"
        class="gantt-wrapper"
    >

        <!-- 上方捲軸左側空白 -->
        <div class="top-left-spacer"></div>

        <div class="top-resizer-spacer"></div>

        <!-- 上方水平捲軸 -->
        <div
            id="topScroll"
            class="external-scroll"
        >
            <div
                class="scroll-width"
            ></div>
        </div>


        <!-- 左側工作項目 -->
        <div
            class="task-labels"
        >
            {labels_html}
        </div>


        <!-- 欄寬調整 -->
        <div
            id="columnResizer"
            class="column-resizer"
            title="左右拖曳調整工作項目欄寬"
        >
        </div>


        <!-- 主時間軸 -->
        <div
            id="timelineScroll"
            class="timeline-scroll"
        >

            <svg
                id="timelineSvg"
                class="timeline-svg"
                width="{timeline_width}"
                height="{svg_height}"
                viewBox="
                    0
                    0
                    {timeline_width}
                    {svg_height}
                "
            >
                {svg_content}
            </svg>

        </div>


        <!-- 下方捲軸左側空白 -->
        <div class="bottom-left-spacer"></div>

        <div class="bottom-resizer-spacer"></div>

        <!-- 下方水平捲軸 -->
        <div
            id="bottomScroll"
            class="external-scroll"
        >
            <div
                class="scroll-width"
            ></div>
        </div>

    </div>


    <div class="legend">

        <div class="legend-item">
            <div class="legend-line"></div>
            工作連動期間
        </div>

        <div class="legend-item">
            <div class="legend-diamond"></div>
            工作完成期限
        </div>

        <div class="legend-item">
            <div class="legend-today"></div>
            今天
        </div>

        <div class="resize-hint">
            ↔ 拖曳分隔線可調整工作項目欄寬
        </div>

    </div>
    """


    # =====================================================
    # CSS
    # =====================================================
    css = f"""
    <style>

        * {{
            box-sizing: border-box;
        }}

        html,
        body {{
            margin: 0;
            padding: 0;
            background: #050505;
            color: #FFFFFF;

            font-family:
                -apple-system,
                BlinkMacSystemFont,
                "Segoe UI",
                "Noto Sans TC",
                sans-serif;
        }}


        .page-container {{
            padding: 18px;
            width: 100%;
        }}


        .page-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            margin-bottom: 16px;
        }}


        .page-header h1 {{
            margin: 0;
            color: #FFFFFF;
            font-size: 24px;
        }}


        .update-time {{
            margin-top: 6px;
            color: #AAAAAA;
            font-size: 12px;
        }}


        #todayButton {{
            border: 1px solid #666666;
            background: #202020;
            color: #FFFFFF;
            padding: 8px 14px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
        }}


        #todayButton:hover {{
            background: #303030;
        }}


        /* =================================================
           主 Grid
           ================================================= */

        .gantt-wrapper {{

            --label-width:
                {LABEL_WIDTH}px;

            display: grid;

            grid-template-columns:
                var(--label-width)
                {RESIZER_WIDTH}px
                minmax(0, 1fr);

            grid-template-rows:
                {SCROLLBAR_HEIGHT}px
                auto
                {SCROLLBAR_HEIGHT}px;

            width: 100%;

            border:
                1px solid #444444;

            border-radius:
                8px;

            overflow:
                hidden;

            background:
                #080808;
        }}


        .top-left-spacer,
        .bottom-left-spacer {{
            background: #111111;
        }}


        .top-resizer-spacer,
        .bottom-resizer-spacer {{
            background: #333333;
        }}


        /* =================================================
           左側工作項目
           ================================================= */

        .task-labels {{
            background: #111111;
            color: #FFFFFF;
            min-width: 0;
        }}


        .task-label-header {{
            height: {HEADER_HEIGHT}px;

            display: flex;
            align-items: center;

            padding: 0 14px;

            font-size: 13px;
            font-weight: 700;

            border-bottom:
                1px solid #444444;

            background:
                #181818;
        }}


        .task-label-row {{
            height: {ROW_HEIGHT}px;

            display: flex;
            align-items: center;

            padding: 7px 14px;

            color: #FFFFFF;

            border-bottom:
                1px solid #303030;

            min-width: 0;
            overflow: hidden;
        }}


        .task-label-text {{
            width: 100%;

            white-space: normal;

            overflow-wrap:
                anywhere;

            word-break:
                break-word;

            line-height:
                1.35;

            font-size:
                13px;

            max-height:
                calc(
                    {ROW_HEIGHT}px
                    -
                    14px
                );

            overflow:
                hidden;
        }}


        .task-label-footer {{
            height: {FOOTER_HEIGHT}px;

            display: flex;
            align-items: center;

            padding: 0 14px;

            color: #AAAAAA;

            font-size: 12px;

            background:
                #181818;

            border-top:
                1px solid #444444;
        }}


        /* =================================================
           欄寬拖曳
           ================================================= */

        .column-resizer {{
            width:
                {RESIZER_WIDTH}px;

            height: 100%;

            background:
                #333333;

            cursor:
                col-resize;

            position:
                relative;

            user-select:
                none;
        }}


        .column-resizer:hover,
        .column-resizer.dragging {{
            background:
                #4DA3FF;
        }}


        .column-resizer::after {{
            content: "";

            position: absolute;

            left: 2px;
            top: 0;
            bottom: 0;

            width: 2px;

            background:
                rgba(
                    255,
                    255,
                    255,
                    0.35
                );
        }}


        /* =================================================
           主時間軸
           ================================================= */

        .timeline-scroll {{
            width: 100%;
            min-width: 0;

            overflow-x:
                auto;

            overflow-y:
                hidden;

            background:
                #080808;

            scrollbar-width:
                none;
        }}


        .timeline-scroll::-webkit-scrollbar {{
            display: none;
        }}


        .timeline-svg {{
            display: block;

            width:
                {timeline_width}px;

            min-width:
                {timeline_width}px;

            height:
                {svg_height}px;
        }}


        /* =================================================
           上下捲軸
           ================================================= */

        .external-scroll {{
            width: 100%;
            min-width: 0;

            overflow-x:
                scroll;

            overflow-y:
                hidden;

            height:
                {SCROLLBAR_HEIGHT}px;

            background:
                #151515;

            scrollbar-color:
                #777777
                #1A1A1A;

            scrollbar-width:
                auto;
        }}


        .external-scroll::-webkit-scrollbar {{
            height:
                {SCROLLBAR_HEIGHT}px;
        }}


        .external-scroll::-webkit-scrollbar-track {{
            background:
                #1A1A1A;
        }}


        .external-scroll::-webkit-scrollbar-thumb {{
            background:
                #777777;

            border-radius:
                8px;

            border:
                3px solid
                #1A1A1A;
        }}


        .external-scroll::-webkit-scrollbar-thumb:hover {{
            background:
                #AAAAAA;
        }}


        .scroll-width {{
            width:
                {timeline_width}px;

            min-width:
                {timeline_width}px;

            height:
                1px;
        }}


        /* =================================================
           圖例
           ================================================= */

        .legend {{
            margin-top: 10px;

            display: flex;

            gap: 22px;

            flex-wrap: wrap;

            color: #FFFFFF;

            font-size: 12px;

            padding:
                10px 12px;

            background:
                #111111;

            border-radius:
                6px;
        }}


        .legend-item {{
            display: flex;
            align-items: center;
            gap: 7px;
        }}


        .legend-line {{
            width: 24px;
            height: 5px;

            background:
                #4DA3FF;

            border-radius:
                3px;
        }}


        .legend-diamond {{
            width: 10px;
            height: 10px;

            background:
                #FFD54F;

            border:
                1px solid #FFFFFF;

            transform:
                rotate(45deg);
        }}


        .legend-today {{
            width: 3px;
            height: 18px;

            background:
                #FF3B30;
        }}


        .resize-hint {{
            margin-left: auto;

            color:
                #AAAAAA;

            font-size: 11px;
        }}

    </style>
    """


    # =====================================================
    # JavaScript
    # =====================================================
    start_iso = (
        start_date.strftime(
            "%Y-%m-%d"
        )
    )


    js = f"""
    <script>

        const DAY_WIDTH =
            {DAY_WIDTH};

        const TIMELINE_START =
            "{start_iso}";

        const SVG_HEIGHT =
            {svg_height};

        const STATIC_MODE =
            {
                "true"
                if static_mode
                else
                "false"
            };


        const topScroller =
            document.getElementById(
                "topScroll"
            );

        const mainScroller =
            document.getElementById(
                "timelineScroll"
            );

        const bottomScroller =
            document.getElementById(
                "bottomScroll"
            );

        const wrapper =
            document.getElementById(
                "ganttWrapper"
            );

        const resizer =
            document.getElementById(
                "columnResizer"
            );

        const svg =
            document.getElementById(
                "timelineSvg"
            );

        const todayButton =
            document.getElementById(
                "todayButton"
            );


        // =================================================
        // 日期工具
        // =================================================

        function localToday() {{

            const now =
                new Date();

            return new Date(
                now.getFullYear(),
                now.getMonth(),
                now.getDate()
            );
        }}


        function parseDate(
            value
        ) {{

            const p =
                value.split(
                    "-"
                );

            return new Date(
                Number(
                    p[0]
                ),
                Number(
                    p[1]
                ) - 1,
                Number(
                    p[2]
                )
            );
        }}


        function daysBetween(
            start,
            end
        ) {{

            const oneDay =
                86400000;

            const a =
                Date.UTC(
                    start.getFullYear(),
                    start.getMonth(),
                    start.getDate()
                );

            const b =
                Date.UTC(
                    end.getFullYear(),
                    end.getMonth(),
                    end.getDate()
                );

            return Math.round(
                (
                    b - a
                )
                /
                oneDay
            );
        }}


        function dateToX(
            date
        ) {{

            const start =
                parseDate(
                    TIMELINE_START
                );

            return (
                daysBetween(
                    start,
                    date
                )
                *
                DAY_WIDTH
                +
                DAY_WIDTH / 2
            );
        }}


        function formatDisplayDate(
            date
        ) {{

            return (
                String(
                    date.getFullYear()
                ).slice(
                    -2
                )
                +
                "/"
                +
                (
                    date.getMonth()
                    +
                    1
                )
                +
                "/"
                +
                date.getDate()
            );
        }}


        // =================================================
        // 三個橫向位置同步
        // =================================================

        let syncingScroll =
            false;


        function syncScroll(
            source
        ) {{

            if (
                syncingScroll
            ) {{
                return;
            }}

            syncingScroll =
                true;

            const left =
                source.scrollLeft;


            if (
                source
                !==
                topScroller
            ) {{
                topScroller.scrollLeft =
                    left;
            }}


            if (
                source
                !==
                mainScroller
            ) {{
                mainScroller.scrollLeft =
                    left;
            }}


            if (
                source
                !==
                bottomScroller
            ) {{
                bottomScroller.scrollLeft =
                    left;
            }}


            requestAnimationFrame(
                function() {{

                    syncingScroll =
                        false;

                }}
            );
        }}


        topScroller.addEventListener(
            "scroll",
            function() {{
                syncScroll(
                    topScroller
                );
            }}
        );


        mainScroller.addEventListener(
            "scroll",
            function() {{
                syncScroll(
                    mainScroller
                );
            }}
        );


        bottomScroller.addEventListener(
            "scroll",
            function() {{
                syncScroll(
                    bottomScroller
                );
            }}
        );


        function setScrollLeft(
            left
        ) {{

            topScroller.scrollLeft =
                left;

            mainScroller.scrollLeft =
                left;

            bottomScroller.scrollLeft =
                left;
        }}


        // =================================================
        // 靜態版今天線
        // =================================================

        function drawTodayMarker() {{

            if (
                !STATIC_MODE
                ||
                !svg
            ) {{
                return;
            }}


            const old =
                document.getElementById(
                    "todayMarker"
                );

            if (old) {{
                old.remove();
            }}


            const today =
                localToday();

            const x =
                dateToX(
                    today
                );


            if (
                x < 0
                ||
                x > {timeline_width}
            ) {{
                return;
            }}


            const NS =
                "http://www.w3.org/2000/svg";


            const group =
                document.createElementNS(
                    NS,
                    "g"
                );

            group.setAttribute(
                "id",
                "todayMarker"
            );


            const line =
                document.createElementNS(
                    NS,
                    "line"
                );

            line.setAttribute(
                "x1",
                x
            );

            line.setAttribute(
                "x2",
                x
            );

            line.setAttribute(
                "y1",
                0
            );

            line.setAttribute(
                "y2",
                SVG_HEIGHT
            );

            line.setAttribute(
                "stroke",
                "#FF3B30"
            );

            line.setAttribute(
                "stroke-width",
                3
            );

            line.setAttribute(
                "stroke-dasharray",
                "7 5"
            );


            const rect =
                document.createElementNS(
                    NS,
                    "rect"
                );

            rect.setAttribute(
                "x",
                x - 46
            );

            rect.setAttribute(
                "y",
                4
            );

            rect.setAttribute(
                "width",
                92
            );

            rect.setAttribute(
                "height",
                22
            );

            rect.setAttribute(
                "rx",
                4
            );

            rect.setAttribute(
                "fill",
                "#D32F2F"
            );


            const text =
                document.createElementNS(
                    NS,
                    "text"
                );

            text.setAttribute(
                "x",
                x
            );

            text.setAttribute(
                "y",
                19
            );

            text.setAttribute(
                "fill",
                "#FFFFFF"
            );

            text.setAttribute(
                "font-size",
                11
            );

            text.setAttribute(
                "font-weight",
                600
            );

            text.setAttribute(
                "text-anchor",
                "middle"
            );

            text.textContent =
                "今天 "
                +
                formatDisplayDate(
                    today
                );


            group.appendChild(
                line
            );

            group.appendChild(
                rect
            );

            group.appendChild(
                text
            );

            svg.appendChild(
                group
            );
        }}


        // =================================================
        // 跳到今天
        // =================================================

        function scrollToToday() {{

            const today =
                localToday();

            const x =
                dateToX(
                    today
                );

            const target =
                Math.max(
                    0,
                    x
                    -
                    mainScroller.clientWidth
                    /
                    2
                );

            setScrollLeft(
                target
            );
        }}


        if (
            todayButton
        ) {{

            todayButton.addEventListener(
                "click",
                scrollToToday
            );
        }}


        // =================================================
        // 初始位置
        // =================================================

        setScrollLeft(
            {initial_scroll}
        );


        // =================================================
        // Excel 式欄寬拖曳
        // =================================================

        let isResizing =
            false;

        let startMouseX =
            0;

        let startWidth =
            {LABEL_WIDTH};

        const minWidth =
            {LABEL_MIN_WIDTH};

        const maxWidth =
            {LABEL_MAX_WIDTH};


        resizer.addEventListener(
            "mousedown",
            function(
                event
            ) {{

                isResizing =
                    true;

                startMouseX =
                    event.clientX;

                const currentWidth =
                    getComputedStyle(
                        wrapper
                    )
                    .getPropertyValue(
                        "--label-width"
                    );

                startWidth =
                    parseFloat(
                        currentWidth
                    );

                resizer.classList.add(
                    "dragging"
                );

                document.body.style.cursor =
                    "col-resize";

                document.body.style.userSelect =
                    "none";
            }}
        );


        window.addEventListener(
            "mousemove",
            function(
                event
            ) {{

                if (
                    !isResizing
                ) {{
                    return;
                }}

                let width =
                    startWidth
                    +
                    event.clientX
                    -
                    startMouseX;

                width =
                    Math.max(
                        minWidth,
                        Math.min(
                            maxWidth,
                            width
                        )
                    );

                wrapper.style.setProperty(
                    "--label-width",
                    width
                    +
                    "px"
                );
            }}
        );


        window.addEventListener(
            "mouseup",
            function() {{

                if (
                    !isResizing
                ) {{
                    return;
                }}

                isResizing =
                    false;

                resizer.classList.remove(
                    "dragging"
                );

                document.body.style.cursor =
                    "";

                document.body.style.userSelect =
                    "";
            }}
        );


        drawTodayMarker();

    </script>
    """


    # =====================================================
    # 靜態 HTML
    # =====================================================
    if static_mode:

        html = f"""
        <!DOCTYPE html>

        <html lang="zh-Hant">

        <head>

            <meta charset="UTF-8">

            <meta
                name="viewport"
                content="
                    width=device-width,
                    initial-scale=1.0
                "
            >

            <title>
                {safe_title}
            </title>

            {css}

        </head>

        <body>

            <div class="page-container">

                {body_content}

            </div>

            {js}

        </body>

        </html>
        """

    else:

        html = f"""
        {css}

        <div class="page-container">

            {body_content}

        </div>

        {js}
        """


    return (
        html,
        svg_height
    )


# =========================================================
# 上傳 Excel
# =========================================================
uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案，甘特圖標題會是上傳的檔名",
    type=["xlsx"],
    help=(
        "可以上傳只有「工作項目」與「完成期限」的一般 Excel，"
        "也可以直接上傳本工具輸出的完整專案 Excel。"
    )
)


# =========================================================
# Excel 已上傳
# =========================================================
if uploaded_file is not None:

    # -----------------------------------------------------
    # Excel 檔名作為專案名稱
    # -----------------------------------------------------
    original_filename = (
        uploaded_file.name
    )

    if original_filename.lower().endswith(
        ".xlsx"
    ):
        project_title = (
            original_filename[:-5]
        )
    else:
        project_title = (
            original_filename
        )


    # -----------------------------------------------------
    # 讀取 Excel
    # -----------------------------------------------------
    try:

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        sheet_names = (
            excel_file.sheet_names
        )

    except Exception as e:

        st.error(
            f"Excel 讀取失敗：{e}"
        )

        st.stop()


    # =====================================================
    # 判斷專案檔
    # =====================================================
    if "工作項目" in sheet_names:

        df = pd.read_excel(
            uploaded_file,
            sheet_name="工作項目"
        )

        is_project_file = True

    else:

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0
        )

        is_project_file = False


    if is_project_file:

        st.success(
            "已辨識為完整甘特圖專案檔。"
        )

    else:

        st.info(
            "已辨識為一般工作資料 Excel。"
        )


    # =====================================================
    # 必要欄位
    # =====================================================
    required_columns = [
        "工作項目",
        "完成期限"
    ]

    missing_columns = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        st.error(
            "Excel 缺少以下必要欄位："
            +
            "、".join(
                missing_columns
            )
        )

        st.stop()


    # =====================================================
    # 整理資料
    # =====================================================
    df["工作項目"] = (
        df[
            "工作項目"
        ]
        .astype(
            str
        )
        .str.strip()
    )


    df["完成期限"] = (
        pd.to_datetime(
            df[
                "完成期限"
            ],
            errors="coerce"
        )
    )


    invalid_date = df[
        df[
            "完成期限"
        ].isna()
    ]

    if not invalid_date.empty:

        st.warning(
            f"有 {len(invalid_date)} 筆完成期限無法辨識，"
            "這些工作不會顯示在甘特圖。"
        )

        st.dataframe(
            invalid_date,
            use_container_width=True
        )


    gantt_df = (
        df.dropna(
            subset=[
                "工作項目",
                "完成期限"
            ]
        )
        .copy()
    )


    if gantt_df.empty:

        st.error(
            "目前沒有可用的工作資料。"
        )

        st.stop()


    # =====================================================
    # 重複工作名稱
    # =====================================================
    duplicated_tasks = gantt_df[
        gantt_df[
            "工作項目"
        ].duplicated(
            keep=False
        )
    ]


    if not duplicated_tasks.empty:

        st.warning(
            "偵測到重複的工作項目名稱。"
            "目前連動功能以工作名稱辨識工作，"
            "建議每一個工作項目使用不同名稱。"
        )


    task_list = gantt_df[
        "工作項目"
    ].tolist()


    # =====================================================
    # 載入專案設定
    # =====================================================
    current_file_name = (
        uploaded_file.name
    )


    if (
        is_project_file
        and
        st.session_state.loaded_project_name
        != current_file_name
    ):

        try:

            (
                loaded_links,
                loaded_holidays,
                loaded_day_mode,
                loaded_workdays,
                invalid_links
            ) = load_project_settings(
                uploaded_file,
                task_list
            )

            st.session_state.links = (
                loaded_links
            )

            st.session_state.holidays = (
                loaded_holidays
            )

            st.session_state.day_mode = (
                loaded_day_mode
            )

            st.session_state.workdays = (
                loaded_workdays
            )

            st.session_state.loaded_project_name = (
                current_file_name
            )

            if invalid_links:

                st.warning(
                    f"有 {len(invalid_links)} 組連動"
                    "因工作項目不存在而未載入。"
                )

        except Exception as e:

            st.warning(
                f"專案設定讀取失敗：{e}"
            )


    # =====================================================
    # 工作資料
    # =====================================================
    st.divider()

    st.subheader(
        "📋 工作資料"
    )


    display_df = gantt_df.copy()


    display_df[
        "完成期限"
    ] = (
        display_df[
            "完成期限"
        ].apply(
            format_date
        )
    )


    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )


    # =====================================================
    # 天數計算
    # =====================================================
    st.divider()

    st.subheader(
        "⚙️ 天數計算設定"
    )


    day_mode = st.radio(
        "計算方式",
        [
            "日曆天",
            "工作天"
        ],
        index=(
            0
            if
            st.session_state.day_mode
            ==
            "日曆天"
            else
            1
        ),
        horizontal=True
    )


    st.session_state.day_mode = (
        day_mode
    )


    # =====================================================
    # 工作天設定
    # =====================================================
    if day_mode == "工作天":

        st.write(
            "### 工作日設定"
        )

        st.caption(
            "勾選哪些星期要視為正常工作日。"
        )


        c1, c2, c3, c4 = (
            st.columns(4)
        )


        with c1:

            monday = st.checkbox(
                "星期一",
                value=(
                    st.session_state
                    .workdays[
                        "monday"
                    ]
                )
            )

            friday = st.checkbox(
                "星期五",
                value=(
                    st.session_state
                    .workdays[
                        "friday"
                    ]
                )
            )


        with c2:

            tuesday = st.checkbox(
                "星期二",
                value=(
                    st.session_state
                    .workdays[
                        "tuesday"
                    ]
                )
            )

            saturday = st.checkbox(
                "星期六",
                value=(
                    st.session_state
                    .workdays[
                        "saturday"
                    ]
                )
            )


        with c3:

            wednesday = st.checkbox(
                "星期三",
                value=(
                    st.session_state
                    .workdays[
                        "wednesday"
                    ]
                )
            )

            sunday = st.checkbox(
                "星期日",
                value=(
                    st.session_state
                    .workdays[
                        "sunday"
                    ]
                )
            )


        with c4:

            thursday = st.checkbox(
                "星期四",
                value=(
                    st.session_state
                    .workdays[
                        "thursday"
                    ]
                )
            )


        st.session_state.workdays = {
            "monday": monday,
            "tuesday": tuesday,
            "wednesday": wednesday,
            "thursday": thursday,
            "friday": friday,
            "saturday": saturday,
            "sunday": sunday
        }


        # -------------------------------------------------
        # 排除日期
        # -------------------------------------------------
        st.write(
            "### 🏖️ 國定假日 / 排除日期"
        )


        holiday_date = st.date_input(
            "選擇要新增的排除日期"
        )


        if st.button(
            "＋ 新增排除日期"
        ):

            holiday_timestamp = (
                pd.Timestamp(
                    holiday_date
                )
                .normalize()
            )


            normalized_holidays = [
                holiday.normalize()
                for holiday
                in st.session_state.holidays
            ]


            if (
                holiday_timestamp
                not in
                normalized_holidays
            ):

                st.session_state.holidays.append(
                    holiday_timestamp
                )

                st.session_state.holidays.sort()

                st.rerun()

            else:

                st.warning(
                    "這個日期已經在排除清單中。"
                )


        if (
            st.session_state.holidays
        ):

            st.write(
                "#### 目前排除日期"
            )


            for index, holiday in enumerate(
                st.session_state.holidays
            ):

                col_date, col_delete = (
                    st.columns(
                        [
                            4,
                            1
                        ]
                    )
                )


                with col_date:

                    st.write(
                        format_date(
                            holiday
                        )
                    )


                with col_delete:

                    if st.button(
                        "刪除",
                        key=(
                            f"delete_holiday_"
                            f"{index}"
                        )
                    ):

                        st.session_state.holidays.pop(
                            index
                        )

                        st.rerun()


            if st.button(
                "清除全部排除日期"
            ):

                st.session_state.holidays = []

                st.rerun()


    # =====================================================
    # weekmask
    # =====================================================
    weekmask = create_weekmask(
        st.session_state.workdays
    )


    if "1" not in weekmask:

        st.error(
            "至少必須設定一個工作日。"
        )

        st.stop()


    # =====================================================
    # 建立連動
    # =====================================================
    st.divider()

    st.subheader(
        "🔗 建立工作連動"
    )


    col1, col2, col3 = (
        st.columns(
            [
                2,
                2,
                1
            ]
        )
    )


    with col1:

        start_task = st.selectbox(
            "開始項目",
            task_list,
            key="start_task"
        )


    with col2:

        end_task = st.selectbox(
            "結束項目",
            task_list,
            key="end_task"
        )


    with col3:

        st.write("")
        st.write("")

        add_link = st.button(
            "＋ 建立連動",
            use_container_width=True
        )


    # =====================================================
    # 新增連動
    # =====================================================
    if add_link:

        if (
            start_task
            ==
            end_task
        ):

            st.warning(
                "開始項目與結束項目不能相同。"
            )

        else:

            start_task_date = (
                gantt_df.loc[
                    gantt_df[
                        "工作項目"
                    ]
                    ==
                    start_task,
                    "完成期限"
                ].iloc[0]
            )


            end_task_date = (
                gantt_df.loc[
                    gantt_df[
                        "工作項目"
                    ]
                    ==
                    end_task,
                    "完成期限"
                ].iloc[0]
            )


            if (
                end_task_date
                <
                start_task_date
            ):

                st.error(
                    "結束項目的日期早於開始項目，"
                    "無法建立連動。"
                )

            else:

                duplicate = any(

                    link[
                        "start"
                    ]
                    ==
                    start_task

                    and

                    link[
                        "end"
                    ]
                    ==
                    end_task

                    for link
                    in st.session_state.links
                )


                if duplicate:

                    st.warning(
                        "這組連動已經存在。"
                    )

                else:

                    st.session_state.links.append(
                        {
                            "start":
                                start_task,

                            "end":
                                end_task
                        }
                    )


                    st.success(
                        f"已建立："
                        f"{start_task}"
                        f" → "
                        f"{end_task}"
                    )


    # =====================================================
    # 目前連動
    # =====================================================
    if (
        st.session_state.links
    ):

        st.write(
            "### 目前連動關係"
        )


        link_table = []


        for link in (
            st.session_state.links
        ):

            start_name = (
                link["start"]
            )

            end_name = (
                link["end"]
            )


            start_rows = gantt_df[
                gantt_df[
                    "工作項目"
                ]
                ==
                start_name
            ]


            end_rows = gantt_df[
                gantt_df[
                    "工作項目"
                ]
                ==
                end_name
            ]


            if (
                start_rows.empty
                or
                end_rows.empty
            ):

                continue


            start_task_date = (
                start_rows[
                    "完成期限"
                ].iloc[0]
            )


            end_task_date = (
                end_rows[
                    "完成期限"
                ].iloc[0]
            )


            calendar_days = (
                calculate_calendar_days(
                    start_task_date,
                    end_task_date
                )
            )


            work_days = (
                calculate_workdays(
                    start_task_date,
                    end_task_date,
                    weekmask,
                    st.session_state.holidays
                )
            )


            link_table.append(
                {
                    "開始項目":
                        start_name,

                    "開始日期":
                        format_date(
                            start_task_date
                        ),

                    "結束項目":
                        end_name,

                    "結束日期":
                        format_date(
                            end_task_date
                        ),

                    "日曆天":
                        calendar_days,

                    "工作天":
                        work_days
                }
            )


        if link_table:

            st.dataframe(
                pd.DataFrame(
                    link_table
                ),
                use_container_width=True,
                hide_index=True
            )


        # -------------------------------------------------
        # 刪除連動
        # -------------------------------------------------
        st.write(
            "#### 🗑️ 刪除連動"
        )


        link_options = [
            (
                i,
                f"{link['start']}"
                f" → "
                f"{link['end']}"
            )
            for i, link
            in enumerate(
                st.session_state.links
            )
        ]


        selected_link = (
            st.selectbox(
                "選擇要刪除的連動",
                options=link_options,
                format_func=lambda x: x[1]
            )
        )


        col_delete, col_clear = (
            st.columns(2)
        )


        with col_delete:

            if st.button(
                "刪除選取的連動",
                use_container_width=True
            ):

                st.session_state.links.pop(
                    selected_link[0]
                )

                st.rerun()


        with col_clear:

            if st.button(
                "清除全部連動",
                use_container_width=True
            ):

                st.session_state.links = []

                st.rerun()


    # =====================================================
    # 甘特圖
    # =====================================================
    st.divider()


    st.subheader(
        f"📊 {project_title}"
    )


    st.caption(
        "甘特圖上下都有調整日期的水平捲軸；"
        "工作項目可拖曳分隔線調整欄寬。"
    )


    gantt_html, gantt_height = (
        create_gantt_html(

            gantt_df=
                gantt_df,

            links=
                st.session_state.links,

            day_mode=
                day_mode,

            weekmask=
                weekmask,

            holidays=
                st.session_state.holidays,

            project_title=
                project_title,

            static_mode=
                False
        )
    )


    components.html(

        gantt_html,

        height=(
            gantt_height
            +
            190
        ),

        scrolling=False
    )


    # =====================================================
    # 匯出 / 分享
    # =====================================================
    st.divider()

    st.subheader(
        "📤 匯出"
    )


    st.write(
        """
        你可以儲存完整專案 Excel，
        或直接產生靜態 HTML。
        """
    )


    # =====================================================
    # 建立專案 Excel
    # =====================================================
    project_excel = (
        create_project_excel(

            task_df=
                gantt_df,

            links=
                st.session_state.links,

            holidays=
                st.session_state.holidays,

            day_mode=
                st.session_state.day_mode,

            workdays=
                st.session_state.workdays
        )
    )


    # =====================================================
    # 建立靜態 HTML
    # =====================================================
    static_html, _ = (
        create_gantt_html(

            gantt_df=
                gantt_df,

            links=
                st.session_state.links,

            day_mode=
                day_mode,

            weekmask=
                weekmask,

            holidays=
                st.session_state.holidays,

            project_title=
                project_title,

            static_mode=
                True
        )
    )


    export_col1, export_col2 = (
        st.columns(2)
    )


    with export_col1:

        st.download_button(

            label=(
                "⬇️ 下載完整專案 Excel"
            ),

            data=
                project_excel,

            file_name=(
                f"{project_title}_甘特圖專案.xlsx"
            ),

            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True
        )


    with export_col2:

        st.download_button(

            label=(
                "🌐 下載靜態甘特圖 HTML"
            ),

            data=
                static_html.encode(
                    "utf-8"
                ),

            file_name=(
                f"{project_title}_甘特圖.html"
            ),

            mime=(
                "text/html"
            ),

            use_container_width=True
        )