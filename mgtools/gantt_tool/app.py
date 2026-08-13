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

st.title("📅 工作時程甘特圖工具")

st.write(
    """
    上傳 Excel 後，系統會將每個工作的完成期限顯示在時間軸上。

    你可以建立工作之間的連動關係、計算日曆天或工作天、
    設定休假日，並將完整專案儲存成一份 Excel 檔。
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
# 日期顯示格式
# 2026-08-13 → 26/8/13
# =========================================================
def format_date(date):

    date = pd.Timestamp(date)

    return (
        f"{date.strftime('%y')}/"
        f"{date.month}/"
        f"{date.day}"
    )


# =========================================================
# 月份顯示格式
# 2026-08 → 26/8
# =========================================================
def format_month(date):

    date = pd.Timestamp(date)

    return (
        f"{date.strftime('%y')}/"
        f"{date.month}"
    )


# =========================================================
# 短日期格式
# 2026-08-13 → 8/13
# =========================================================
def format_short_date(date):

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

    # busday_count 不包含結束日期
    # 因此加一天，採含頭含尾
    end_np = (
        np.datetime64(end_date.date())
        + np.timedelta64(1, "D")
    )

    holiday_array = np.array(
        [
            np.datetime64(
                holiday.strftime("%Y-%m-%d")
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

    # -----------------------------------------------------
    # 工作項目
    # -----------------------------------------------------
    task_export = task_df.copy()

    task_export["完成期限"] = pd.to_datetime(
        task_export["完成期限"]
    )

    # -----------------------------------------------------
    # 連動關係
    # -----------------------------------------------------
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

    # -----------------------------------------------------
    # 排除日期
    # Excel 內部仍使用標準日期格式
    # -----------------------------------------------------
    holidays_export = pd.DataFrame(
        {
            "日期": [
                holiday.strftime("%Y-%m-%d")
                for holiday in holidays
            ]
        }
    )

    # -----------------------------------------------------
    # 系統設定
    # -----------------------------------------------------
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

    # -----------------------------------------------------
    # 寫入 Excel
    # -----------------------------------------------------
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
# 讀取完整專案設定
# =========================================================
def load_project_settings(
    uploaded_file,
    task_list
):

    excel_file = pd.ExcelFile(
        uploaded_file
    )

    sheet_names = excel_file.sheet_names

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
    # 連動關係
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

                start_name = str(start_name)
                end_name = str(end_name)

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

            for value in holidays_df["日期"]:

                if pd.notna(value):

                    date = pd.to_datetime(
                        value,
                        errors="coerce"
                    )

                    if pd.notna(date):

                        loaded_holidays.append(
                            pd.Timestamp(date).normalize()
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
                    settings_df["設定"],
                    settings_df["值"]
                )
            )


            if "day_mode" in settings_dict:

                value = str(
                    settings_dict["day_mode"]
                )

                if value in [
                    "日曆天",
                    "工作天"
                ]:

                    loaded_day_mode = value


            for key in loaded_workdays.keys():

                if key in settings_dict:

                    value = settings_dict[key]

                    if isinstance(
                        value,
                        (bool, np.bool_)
                    ):

                        loaded_workdays[key] = bool(
                            value
                        )

                    else:

                        loaded_workdays[key] = (
                            str(value).lower()
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
# 建立 HTML 甘特圖
# =========================================================
def create_gantt_html(
    gantt_df,
    links,
    day_mode,
    weekmask,
    holidays
):

    # =====================================================
    # 甘特圖尺寸設定
    # =====================================================

    # 每一天固定 28px
    DAY_WIDTH = 28

    # 每一工作列高度
    ROW_HEIGHT = 70

    # 表頭高度
    HEADER_HEIGHT = 75

    # 工作名稱預設寬度
    LABEL_WIDTH = 250

    # 工作名稱允許拖曳的最小 / 最大寬度
    LABEL_MIN_WIDTH = 150
    LABEL_MAX_WIDTH = 650

    # 拖曳分隔線寬度
    RESIZER_WIDTH = 7

    # 日期前後留白
    DATE_PADDING = 7


    today = pd.Timestamp.today().normalize()


    # =====================================================
    # 日期範圍
    # =====================================================
    all_dates = gantt_df[
        "完成期限"
    ].dropna()


    min_date = min(
        all_dates.min(),
        today
    )

    max_date = max(
        all_dates.max(),
        today
    )


    start_date = (
        min_date
        - pd.Timedelta(
            days=DATE_PADDING
        )
    ).normalize()


    end_date = (
        max_date
        + pd.Timedelta(
            days=DATE_PADDING
        )
    ).normalize()


    total_days = (
        end_date
        - start_date
    ).days + 1


    timeline_width = (
        total_days
        * DAY_WIDTH
    )


    svg_height = (
        HEADER_HEIGHT
        +
        len(gantt_df) * ROW_HEIGHT
    )


    # =====================================================
    # 日期 → X 座標
    # =====================================================
    def date_to_x(date):

        date = pd.Timestamp(
            date
        ).normalize()

        day_index = (
            date
            - start_date
        ).days

        return (
            day_index * DAY_WIDTH
            +
            DAY_WIDTH / 2
        )


    # =====================================================
    # 工作項目 → Y 座標
    # =====================================================
    task_positions = {}

    for index, row in gantt_df.reset_index(
        drop=True
    ).iterrows():

        task_positions[
            row["工作項目"]
        ] = (
            HEADER_HEIGHT
            +
            index * ROW_HEIGHT
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


    # =====================================================
    # SVG
    # =====================================================
    svg_parts = []


    # -----------------------------------------------------
    # 黑色背景
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
    # 每日背景與垂直線
    # =====================================================
    for i in range(total_days):

        current_date = (
            start_date
            +
            pd.Timedelta(days=i)
        )


        x = i * DAY_WIDTH


        # -------------------------------------------------
        # 週末
        # -------------------------------------------------
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


        # -------------------------------------------------
        # 每日格線
        # -------------------------------------------------
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
    # 月份表頭
    # =====================================================
    previous_month = None


    for i in range(total_days):

        current_date = (
            start_date
            +
            pd.Timedelta(days=i)
        )


        month_key = (
            current_date.year,
            current_date.month
        )


        if month_key != previous_month:

            x = i * DAY_WIDTH


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


            previous_month = month_key


    # =====================================================
    # 日期表頭
    # =====================================================
    for i in range(total_days):

        current_date = (
            start_date
            +
            pd.Timedelta(days=i)
        )


        if (
            i % 7 == 0
            or
            current_date.day == 1
        ):

            x = (
                i * DAY_WIDTH
                +
                DAY_WIDTH / 2
            )


            svg_parts.append(
                f"""
                <text
                    x="{x}"
                    y="52"
                    fill="#CFCFCF"
                    font-size="11"
                    text-anchor="middle"
                >
                    {format_short_date(current_date)}
                </text>
                """
            )


    # =====================================================
    # 水平工作列格線
    # =====================================================
    for i in range(
        len(gantt_df) + 1
    ):

        y = (
            HEADER_HEIGHT
            +
            i * ROW_HEIGHT
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
    # 今天
    # =====================================================
    today_x = date_to_x(
        today
    )


    svg_parts.append(
        f"""
        <line
            x1="{today_x}"
            y1="0"
            x2="{today_x}"
            y2="{svg_height}"
            stroke="#FF3B30"
            stroke-width="3"
            stroke-dasharray="7 5"
        />
        """
    )


    today_label = (
        "今天 "
        +
        format_date(today)
    )


    today_label_width = 92


    svg_parts.append(
        f"""
        <rect
            x="{today_x - today_label_width / 2}"
            y="4"
            width="{today_label_width}"
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
        """
    )


    # =====================================================
    # 工作連動
    # =====================================================
    for link in links:

        start_name = link[
            "start"
        ]

        end_name = link[
            "end"
        ]


        start_rows = gantt_df[
            gantt_df["工作項目"]
            == start_name
        ]

        end_rows = gantt_df[
            gantt_df["工作項目"]
            == end_name
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


        start_x = date_to_x(
            start_task_date
        )

        end_x = date_to_x(
            end_task_date
        )


        start_y = task_positions[
            start_name
        ]

        end_y = task_positions[
            end_name
        ]


        # -------------------------------------------------
        # 天數
        # -------------------------------------------------
        calendar_days = (
            calculate_calendar_days(
                start_task_date,
                end_task_date
            )
        )


        work_days = calculate_workdays(
            start_task_date,
            end_task_date,
            weekmask,
            holidays
        )


        if day_mode == "日曆天":

            display_text = (
                f"{calendar_days} 天"
            )

        else:

            display_text = (
                f"{work_days} 工作天"
            )


        LINE_COLOR = "#4DA3FF"


        # -------------------------------------------------
        # 水平線
        # -------------------------------------------------
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


        # -------------------------------------------------
        # 垂直連線
        # -------------------------------------------------
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


        # -------------------------------------------------
        # 天數標籤
        # -------------------------------------------------
        middle_x = (
            start_x
            +
            end_x
        ) / 2


        label_width = max(
            62,
            len(display_text) * 15
        )


        svg_parts.append(
            f"""
            <rect
                x="{middle_x - label_width / 2}"
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
    # 所有里程碑
    # =====================================================
    for _, row in gantt_df.iterrows():

        task_name = row[
            "工作項目"
        ]

        milestone_date = row[
            "完成期限"
        ]


        x = date_to_x(
            milestone_date
        )

        y = task_positions[
            task_name
        ]


        date_text = format_date(
            milestone_date
        )


        # -------------------------------------------------
        # 日期在菱形上方
        # -------------------------------------------------
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


        # -------------------------------------------------
        # 菱形
        # -------------------------------------------------
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
                    {escape(str(task_name))}
                    - {date_text}
                </title>
            </polygon>
            """
        )


    # =====================================================
    # 組合 SVG
    # =====================================================
    svg_content = "".join(
        svg_parts
    )


    # =====================================================
    # 預設移動至今天附近
    # =====================================================
    initial_scroll = max(
        0,
        today_x - 450
    )


    # =====================================================
    # HTML
    # =====================================================
    html = f"""
    <style>

        * {{
            box-sizing: border-box;
        }}


        body {{

            margin: 0;

            padding: 0;

            background: transparent;

            font-family:
                -apple-system,
                BlinkMacSystemFont,
                "Segoe UI",
                "Noto Sans TC",
                sans-serif;
        }}


        /* =================================================
           甘特圖主框
           ================================================= */

        .gantt-wrapper {{

            --label-width:
                {LABEL_WIDTH}px;

            display: grid;

            grid-template-columns:
                var(--label-width)
                {RESIZER_WIDTH}px
                minmax(0, 1fr);

            width: 100%;

            border:
                1px solid
                #444444;

            border-radius:
                8px;

            overflow: hidden;

            background:
                #080808;
        }}


        /* =================================================
           左側工作名稱區
           ================================================= */

        .task-labels {{

            background:
                #111111;

            color:
                #FFFFFF;

            min-width: 0;

            z-index: 2;
        }}


        .task-label-header {{

            height:
                {HEADER_HEIGHT}px;

            display: flex;

            align-items: center;

            padding:
                0 14px;

            font-size:
                13px;

            font-weight:
                700;

            border-bottom:
                1px solid
                #444444;

            background:
                #181818;
        }}


        /* =================================================
           工作項目格子

           重要：
           取消 ellipsis
           改成自動換行
           ================================================= */

        .task-label-row {{

            height:
                {ROW_HEIGHT}px;

            display: flex;

            align-items: center;

            padding:
                7px 14px;

            font-size:
                13px;

            font-weight:
                500;

            color:
                #FFFFFF;

            border-bottom:
                1px solid
                #303030;

            min-width: 0;

            overflow: hidden;
        }}


        .task-label-text {{

            width: 100%;

            white-space:
                normal;

            overflow-wrap:
                anywhere;

            word-break:
                break-word;

            line-height:
                1.35;

            color:
                #FFFFFF;

            max-height:
                calc(
                    {ROW_HEIGHT}px
                    - 14px
                );

            overflow:
                hidden;
        }}


        /* =================================================
           可拖曳欄寬分隔線
           ================================================= */

        .column-resizer {{

            width:
                {RESIZER_WIDTH}px;

            height:
                100%;

            background:
                #333333;

            cursor:
                col-resize;

            position:
                relative;

            z-index:
                10;

            user-select:
                none;
        }}


        .column-resizer:hover {{

            background:
                #4DA3FF;
        }}


        .column-resizer.dragging {{

            background:
                #4DA3FF;
        }}


        .column-resizer::after {{

            content:
                "";

            position:
                absolute;

            top:
                0;

            bottom:
                0;

            left:
                2px;

            width:
                2px;

            background:
                rgba(
                    255,
                    255,
                    255,
                    0.35
                );
        }}


        /* =================================================
           時間軸 + 真正水平捲軸
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

            scrollbar-color:
                #777777
                #1A1A1A;

            scrollbar-width:
                auto;
        }}


        .timeline-scroll::-webkit-scrollbar {{

            height:
                16px;
        }}


        .timeline-scroll::-webkit-scrollbar-track {{

            background:
                #1A1A1A;
        }}


        .timeline-scroll::-webkit-scrollbar-thumb {{

            background:
                #777777;

            border-radius:
                8px;

            border:
                3px solid
                #1A1A1A;
        }}


        .timeline-scroll::-webkit-scrollbar-thumb:hover {{

            background:
                #AAAAAA;
        }}


        .timeline-svg {{

            display:
                block;

            width:
                {timeline_width}px;

            min-width:
                {timeline_width}px;

            height:
                {svg_height}px;
        }}


        /* =================================================
           圖例
           ================================================= */

        .legend {{

            margin-top:
                10px;

            display:
                flex;

            gap:
                22px;

            flex-wrap:
                wrap;

            color:
                #FFFFFF;

            font-size:
                12px;

            padding:
                10px 12px;

            background:
                #111111;

            border-radius:
                6px;
        }}


        .legend-item {{

            display:
                flex;

            align-items:
                center;

            gap:
                7px;
        }}


        .legend-line {{

            width:
                24px;

            height:
                5px;

            background:
                #4DA3FF;

            border-radius:
                3px;
        }}


        .legend-diamond {{

            width:
                10px;

            height:
                10px;

            background:
                #FFD54F;

            border:
                1px solid
                #FFFFFF;

            transform:
                rotate(45deg);
        }}


        .legend-today {{

            width:
                3px;

            height:
                18px;

            background:
                #FF3B30;
        }}


        .resize-hint {{

            margin-left:
                auto;

            color:
                #AAAAAA;

            font-size:
                11px;
        }}

    </style>


    <div
        id="ganttWrapper"
        class="gantt-wrapper"
    >

        <!-- ==============================================
             左側工作項目
             ============================================== -->

        <div class="task-labels">

            {labels_html}

        </div>


        <!-- ==============================================
             Excel 類似的拖曳欄寬分隔線
             ============================================== -->

        <div
            id="columnResizer"
            class="column-resizer"
            title="左右拖曳調整工作項目欄寬"
        >
        </div>


        <!-- ==============================================
             右側時間軸
             ============================================== -->

        <div
            id="timelineScroll"
            class="timeline-scroll"
        >

            <svg
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

    </div>


    <!-- =================================================
         圖例
         ================================================= -->

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


        <div class="legend-item">

            每一天固定 {DAY_WIDTH}px

        </div>


        <div class="resize-hint">

            ↔ 拖曳分隔線可調整工作項目欄寬

        </div>

    </div>


    <script>

        // =================================================
        // 預設將時間軸移到今天附近
        // =================================================

        const scroller =
            document.getElementById(
                "timelineScroll"
            );


        if (scroller) {{

            scroller.scrollLeft =
                {initial_scroll};

        }}


        // =================================================
        // Excel 類似的欄寬拖曳功能
        // =================================================

        const wrapper =
            document.getElementById(
                "ganttWrapper"
            );


        const resizer =
            document.getElementById(
                "columnResizer"
            );


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


        if (
            wrapper
            &&
            resizer
        ) {{

            // ---------------------------------------------
            // 開始拖曳
            // ---------------------------------------------

            resizer.addEventListener(
                "mousedown",
                function(event) {{

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


            // ---------------------------------------------
            // 拖曳中
            // ---------------------------------------------

            window.addEventListener(
                "mousemove",
                function(event) {{

                    if (
                        !isResizing
                    ) {{

                        return;

                    }}


                    const delta =
                        event.clientX
                        -
                        startMouseX;


                    let newWidth =
                        startWidth
                        +
                        delta;


                    // 最小欄寬
                    if (
                        newWidth
                        <
                        minWidth
                    ) {{

                        newWidth =
                            minWidth;

                    }}


                    // 最大欄寬
                    if (
                        newWidth
                        >
                        maxWidth
                    ) {{

                        newWidth =
                            maxWidth;

                    }}


                    wrapper.style.setProperty(
                        "--label-width",
                        newWidth
                        +
                        "px"
                    );

                }}
            );


            // ---------------------------------------------
            // 停止拖曳
            // ---------------------------------------------

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

        }}

    </script>
    """


    return html, svg_height


# =========================================================
# 上傳 Excel
# =========================================================
uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案",
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
    # 判斷專案檔 / 一般 Excel
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
    # 整理工作名稱
    # =====================================================
    df["工作項目"] = (
        df["工作項目"]
        .astype(str)
        .str.strip()
    )


    # =====================================================
    # 日期轉換
    # =====================================================
    df["完成期限"] = (
        pd.to_datetime(
            df["完成期限"],
            errors="coerce"
        )
    )


    # =====================================================
    # 日期錯誤
    # =====================================================
    invalid_date = df[
        df["完成期限"].isna()
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


    # =====================================================
    # 有效工作資料
    # =====================================================
    gantt_df = df.dropna(
        subset=[
            "工作項目",
            "完成期限"
        ]
    ).copy()


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


    task_list = (
        gantt_df[
            "工作項目"
        ].tolist()
    )


    # =====================================================
    # 自動載入專案設定
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


    display_df = (
        gantt_df.copy()
    )


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
    # 天數計算設定
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
            if st.session_state.day_mode
            == "日曆天"
            else 1
        ),
        horizontal=True
    )


    st.session_state.day_mode = (
        day_mode
    )


    # =====================================================
    # 工作天模式
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
                value=st.session_state.workdays[
                    "monday"
                ]
            )

            friday = st.checkbox(
                "星期五",
                value=st.session_state.workdays[
                    "friday"
                ]
            )


        with c2:

            tuesday = st.checkbox(
                "星期二",
                value=st.session_state.workdays[
                    "tuesday"
                ]
            )

            saturday = st.checkbox(
                "星期六",
                value=st.session_state.workdays[
                    "saturday"
                ]
            )


        with c3:

            wednesday = st.checkbox(
                "星期三",
                value=st.session_state.workdays[
                    "wednesday"
                ]
            )

            sunday = st.checkbox(
                "星期日",
                value=st.session_state.workdays[
                    "sunday"
                ]
            )


        with c4:

            thursday = st.checkbox(
                "星期四",
                value=st.session_state.workdays[
                    "thursday"
                ]
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


        # =================================================
        # 排除日期
        # =================================================
        st.write(
            "### 🏖️ 國定假日 / 排除日期"
        )


        st.caption(
            "這些日期即使原本是工作日，也不會計入工作天。"
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
                not in normalized_holidays
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


        # =================================================
        # 顯示排除日期
        # =================================================
        if st.session_state.holidays:

            st.write(
                "#### 目前排除日期"
            )


            for index, holiday in enumerate(
                st.session_state.holidays
            ):

                col_date, col_delete = (
                    st.columns(
                        [4, 1]
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


        else:

            st.info(
                "目前沒有設定排除日期。"
            )


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
    # 建立工作連動
    # =====================================================
    st.divider()

    st.subheader(
        "🔗 建立工作連動"
    )


    col1, col2, col3 = (
        st.columns(
            [2, 2, 1]
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

        if start_task == end_task:

            st.warning(
                "開始項目與結束項目不能相同。"
            )


        else:

            start_task_date = gantt_df.loc[
                gantt_df["工作項目"]
                == start_task,
                "完成期限"
            ].iloc[0]


            end_task_date = gantt_df.loc[
                gantt_df["工作項目"]
                == end_task,
                "完成期限"
            ].iloc[0]


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

                    link["start"]
                    == start_task

                    and

                    link["end"]
                    == end_task

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
    # 目前連動關係
    # =====================================================
    if st.session_state.links:

        st.write(
            "### 目前連動關係"
        )


        link_table = []


        for link in st.session_state.links:

            start_name = (
                link["start"]
            )

            end_name = (
                link["end"]
            )


            start_rows = gantt_df[
                gantt_df["工作項目"]
                == start_name
            ]


            end_rows = gantt_df[
                gantt_df["工作項目"]
                == end_name
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

            link_df = pd.DataFrame(
                link_table
            )


            st.dataframe(
                link_df,
                use_container_width=True,
                hide_index=True
            )


        # =================================================
        # 刪除連動
        # =================================================
        st.write(
            "#### 🗑️ 刪除連動"
        )


        link_options = []


        for i, link in enumerate(
            st.session_state.links
        ):

            label = (
                f"{link['start']}"
                f" → "
                f"{link['end']}"
            )


            link_options.append(
                (
                    i,
                    label
                )
            )


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


    else:

        st.info(
            "目前尚未建立任何工作連動。"
        )


    # =====================================================
    # 甘特圖
    # =====================================================
    st.divider()

    st.subheader(
        "📊 網調時程表"
    )


    st.caption(
        ""
    )


    # =====================================================
    # 建立 HTML 甘特圖
    # =====================================================
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
                st.session_state.holidays
        )
    )


    # =====================================================
    # 顯示甘特圖
    # =====================================================
    components.html(

        gantt_html,

        height=(
            gantt_height
            +
            90
        ),

        scrolling=False
    )


    # =====================================================
    # 使用說明
    # =====================================================
    st.info(
        """
        **甘特圖操作方式**

        - 使用底部的水平捲軸可以調整時間軸。
        - 可拖曳工作項目欄右側的分隔線調整欄寬。
        - 黃色菱形代表工作完成期限。
        - 菱形上方的白色日期為該工作的完成期限。
        - 藍色線段代表兩個工作項目之間的期間。
        - 紅色虛線代表今天。
        """
    )


    # =====================================================
    # 儲存完整專案
    # =====================================================
    st.divider()

    st.subheader(
        "💾 儲存完整甘特圖專案"
    )


    st.write(
        """
        將目前的工作資料、連動關係、
        排除日期與工作日設定全部儲存成一份 Excel。

        下次重新上傳這份 Excel，
        即可恢復目前專案。
        """
    )


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


    st.download_button(

        label=(
            "⬇️ 下載完整專案 Excel"
        ),

        data=
            project_excel,

        file_name=(
            "甘特圖專案.xlsx"
        ),

        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        use_container_width=True
    )