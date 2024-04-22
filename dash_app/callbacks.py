from dash import callback, Input, Output, State, ctx
from pathlib import Path
import pandas as pd
import datetime as dt
import plotly.graph_objects as go
import datetime
import numpy as np
import dash_bootstrap_components as dbc
from dash import dcc
from src.font_size import get_text_dims



def find_lowest_missing(arr):
    # Sort the array
    arr_sorted = np.sort(arr)

    # Check for the first gap between consecutive elements
    for i in range(len(arr_sorted)):
        if arr_sorted[i] != i:
            return i
    # If no gap found, the lowest missing value is n+1
    return len(arr_sorted)


def read_data(table_name: str) -> tuple:
    """Load table data and return it in a json friendly format.

    Args:
        table_name (str): table name

    Returns:
        list: list of data from table
    """
    result = pd.read_csv(Path.cwd() / table_name)
    assert result.size > 0
    return result.to_dict("records"), [
        {"name": i, "id": i} for i in result.columns
    ]


@callback(
    Output("timetable", "data"),
    Output("timetable", "columns", allow_duplicate=True),
    Input("editing-rows-button", "n_clicks"),
    Input("10_min", "n_intervals"),
    State("timetable", "data"),
    State("timetable", "columns"),
    prevent_initial_call="initial_duplicate",
)
def add_row_tt(n_clicks, _, rows, columns):
    if ctx.triggered_id == "editing-rows-button":
        if n_clicks > 0:
            rows.append({c["id"]: "" for c in columns})
        return rows, columns
    return read_data("timetable.csv")


@callback(
    Output("font_sizes", "data"),
    Output("font_sizes", "columns"),
    Input("editing-rows-button-fs", "n_clicks"),
    Input("10_min", "n_intervals"),
    State("font_sizes", "data"),
    State("font_sizes", "columns"),
)
def add_row_fs(n_clicks, _, rows, columns):
    if ctx.triggered_id == "editing-rows-button-fs":
        if n_clicks > 0:
            rows.append({c["id"]: "" for c in columns})
        return rows, columns
    return read_data("font_size.csv")


@callback(
    Output("timetable", "dropdown"),
    Output("timetable", "columns", allow_duplicate=True),
    Input("timetable", "data"),
    State("timetable", "columns"),
    prevent_initial_call="initial_duplicate",
)
def add_table_dropdowns(data, columns):
    _ = data
    if columns is not None:
        dd_dict = {
            "Tag": {
                "options": [
                    {"label": "Montag", "value": "Montag"},
                    {"label": "Dienstag", "value": "Dienstag"},
                    {"label": "Mittwoch", "value": "Mittwoch"},
                    {"label": "Donnerstag", "value": "Donnerstag"},
                    {"label": "Freitag", "value": "Freitag"},
                    {"label": "Samstag", "value": "Samstag"},
                    {"label": "Sonntag", "value": "Sonntag"},
                ]
            }
        }
        columns = [
            {"id": col["id"], "name": col["id"], "presentation": "dropdown"}
            if col["id"] in dd_dict
            else {"id": col["id"], "name": col["id"]}
            for col in columns
        ]
        return dd_dict, columns
    return {}, []


def convert_to_offset(w_str: str) -> int:
    wd_dict = {
        "Montag": 0,
        "Dienstag": 1,
        "Mittwoch": 2,
        "Donnerstag": 3,
        "Freitag": 4,
        "Samstag": 5,
        "Sonntag": 6,
    }
    res = wd_dict.get(w_str)
    assert res is not None
    return res


@callback(Output("day_width", "children"), Input("end_day_dd", "value"))
def create_day_width_fields(end_day_val) -> list:
    if end_day_val:
        return [
            dbc.Col(dcc.Input(value=1, type="number", min=1, step=1))
            for i in range(end_day_val)
        ]
    return []


def add_annotations(
    row: pd.Series,
    fig: go.Figure,
    font: str,
    px_font_size: int,
    max_width: int,
    day_width: np.ndarray,
) -> None:
    width = day_width / (max_width)
    offset = width * (row["offset"])
    startzeit = row["Startzeit"]
    endzeit = row["Endzeit"]
    time = row["Startzeitpunkt"].time()
    for i, text in enumerate([row["Tutor:in"], row["Schwerpunkt"]]):
        text_dims = get_text_dims(
            str(text), int(px_font_size * (96 / 72)), font
        )
        fig.add_annotation(
            x=float(row["plot_day"] + offset),
            y=float(
                time.second
                + 60 * time.minute
                + 60**2 * time.hour
                # + row["diff"]
            ),
            text=str(text),
            ax=0,
            ay=0,
            showarrow=False,
            align="left",
            valign="bottom",
            font=dict(
                family=font,
                size=px_font_size,
            ),
            xanchor="left",
            yanchor="middle",
            xshift=0,  # -text_dims[0]/2,
            yshift=-text_dims[1] * (2 * i + 1) / 2,
        )
    room_txt = row["Raum"] or ""
    text_dims = get_text_dims(room_txt, int(px_font_size * (96 / 72)), font)
    fig.add_annotation(
        x=float(row["plot_day"] + offset + width),
        y=float(
            time.second + 60 * time.minute + 60**2 * time.hour + row["diff"]
        ),
        text=room_txt,
        ax=0,
        ay=0,
        showarrow=False,
        align="left",
        valign="bottom",
        font=dict(
            family=font,
            size=px_font_size,
        ),
        xanchor="right",
        yanchor="middle",
        xshift=0,  # -text_dims[0]/2,
        yshift=text_dims[1] / 2,
    )
    text_dims = get_text_dims(
        f"{startzeit}-{endzeit}", int(px_font_size * (96 / 72)), font
    )
    fig.add_annotation(
        x=float(row["plot_day"] + offset + width),
        y=float(time.second + 60 * time.minute + 60**2 * time.hour),
        text=f"{startzeit}-{endzeit}",
        ax=0,
        ay=0,
        showarrow=False,
        align="left",
        valign="bottom",
        font=dict(
            family=font,
            size=px_font_size,
        ),
        xanchor="right",
        yanchor="middle",
        xshift=0,  # -text_dims[0]/2,
        yshift=-text_dims[1] / 2,
    )
    text_dims = get_text_dims(
        str(row["Ort"]), int(px_font_size * (96 / 72)), font
    )
    fig.add_annotation(
        x=float(row["plot_day"] + offset),
        y=float(
            time.second + 60 * time.minute + 60**2 * time.hour + row["diff"]
        ),
        text=str(row["Ort"]),
        ax=0,
        ay=0,
        showarrow=False,
        align="left",
        valign="bottom",
        font=dict(
            family=font,
            size=px_font_size,
        ),
        xanchor="left",
        yanchor="middle",
        xshift=0,  # -text_dims[0]/2,
        yshift=text_dims[1] / 2,
    )


@callback(
    Output("collapse-fs", "is_open"),
    [Input("collapse-button-fs", "n_clicks")],
    [State("collapse-fs", "is_open")],
)
def toggle_collapse_fs(n, is_open):
    if n:
        return not is_open
    return is_open


@callback(
    Output("tt-graph", "figure"),
    Input("create-plan", "n_clicks"),
    State("timetable", "data"),
    State("date-picker", "value"),
    State("end_day_dd", "value"),
    State("day_width", "children"),
    State("font_sizes", "data"),
    prevent_initial_call=True,
)
def create_chart(_, data, date_str, end_day_val, day_widths, fs_data):
    if (
        data is not None
        and date_str is not None
        and end_day_val is not None
        and day_widths is not None
    ):
        widths = [
            dw.get("props", {})
            .get("children", {})
            .get("props", {})
            .get("value")
            for dw in day_widths
        ]
        day_offset = np.cumsum(widths)
        day_offset = np.insert(day_offset, 0, 0, axis=0)
        df = pd.DataFrame(data)
        sel_week = dt.datetime.fromisoformat(date_str).strftime("%V %G")
        date = dt.datetime.strptime(f"1 {sel_week}", "%w %V %G").date()
        df["Startzeitpunkt"] = df.apply(
            lambda x: dt.datetime.combine(
                date,
                dt.datetime.strptime(x["Startzeit"], "%H:%M").time(),
            ),
            axis=1,
        )
        df["Endzeitpunkt"] = df.apply(
            lambda x: dt.datetime.combine(
                date,
                dt.datetime.strptime(x["Endzeit"], "%H:%M").time(),
            ),
            axis=1,
        )
        df["diff"] = (
            df["Endzeitpunkt"] - df["Startzeitpunkt"]
        ).dt.total_seconds()
        df["week_day"] = df["Tag"].apply(convert_to_offset)
        df["plot_day"] = df.apply(
            lambda x: day_offset[x.week_day],
            axis=1,
        )
        df["color"] = df["color"].fillna(
            df["Ort"].map(
                {
                    "Rüsselsheim": "#E2EFDA",
                    "WBS": "#DDEBF7",
                    "Online": "#FFF2CC",
                }
            )
        )
        df["offset"] = pd.Series([pd.NA] * df.shape[0])
        df = df.sort_values(["Startzeitpunkt", "Endzeitpunkt"]).reset_index(
            drop=True
        )
        fs_df = (
            pd.DataFrame(fs_data)
            .astype(int)
            .sort_values("Dauer")
            .reset_index(drop=True)
        )
        fs_df.to_csv("./font_size.csv", index=False)
        layout = go.Layout(paper_bgcolor="#fff", plot_bgcolor="#fff")
        fig = go.Figure(layout=layout)
        sd_df = {
            tag: df.loc[df.Tag == tag, :].reset_index(names=["old_ind"])
            for tag in df.Tag.unique()
        }
        show_legend = {"Rüsselsheim": True, "WBS": True, "Online": True}

        x_tick_labels = [
            "Montag",
            "Dienstag",
            "Mittwoch",
            "Donnerstag",
            "Freitag",
            "Samstag",
            "Sonntag",
        ]

        x_tick_vals = [
            day_offset[i] + (day_offset[i + 1] - day_offset[i]) / 2
            for i in range(len(day_offset) - 1)
        ]
        y_tick_vals = np.arange(0, 3600 * 24, 3600)
        y_tick_labels = [
            datetime.time(tick // 3600).strftime("%H:%M")
            for tick in y_tick_vals
        ]
        fig.update_yaxes(
            autorange="reversed",
            mirror="ticks",
            gridcolor="black",
            tickmode="array",
            tickvals=y_tick_vals,
            ticktext=y_tick_labels,
        )

        fig.update_xaxes(
            range=[
                -0.5,
                day_offset[-1] + 0.5,
            ],
            tickmode="array",
            tickvals=x_tick_vals,
            ticktext=x_tick_labels[:end_day_val],
            side="top",
        )  # ,ticklabelmode="period")#, )
        fig.update_layout(
            barmode="group",
            bargap=0.1,
            bargroupgap=0,
        )

        for _, wd_df in sd_df.items():
            for _, row in wd_df.iterrows():
                overlap_df = wd_df.loc[
                    (
                        (wd_df.Startzeitpunkt <= row.Startzeitpunkt)
                        & (row.Startzeitpunkt < wd_df.Endzeitpunkt)
                    )
                    | (
                        (row.Startzeitpunkt <= wd_df.Startzeitpunkt)
                        & (wd_df.Startzeitpunkt < row.Endzeitpunkt)
                    )
                    | (
                        (row.Endzeitpunkt >= wd_df.Endzeitpunkt)
                        & (wd_df.Endzeitpunkt > row.Startzeitpunkt)
                    )
                    | (
                        (wd_df.Endzeitpunkt >= row.Endzeitpunkt)
                        & (row.Endzeitpunkt > wd_df.Startzeitpunkt)
                    )
                ].reset_index(drop=True)
                used_offsets = (
                    df.loc[overlap_df["old_ind"], "offset"]
                    .to_frame()
                    .reset_index(drop=True)
                    .loc[~overlap_df["offset"].isna()]
                    .values
                )
                if np.size(used_offsets) == 0:
                    cur_offset = 0
                else:
                    cur_offset = find_lowest_missing(used_offsets)
                tutor = row["Tutor:in"]
                ind = (
                    df.query(
                        f"`Tutor:in` == '{tutor}' "
                        f" and Ort == '{row.Ort}'"
                        f" and Schwerpunkt == '{row.Schwerpunkt}'"
                    )
                    .loc[
                        (df.Startzeitpunkt == row.Startzeitpunkt)
                        & (df.Endzeitpunkt == row.Endzeitpunkt)
                    ]
                    .index.values[0]
                )
                df.loc[ind, "offset"] = cur_offset
                wd_df.loc[wd_df.old_ind == ind, "offset"] = cur_offset
            max_width = wd_df["offset"].max() + 1
            for k, row in wd_df.iterrows():
                day_width = (
                    day_offset[row.week_day + 1] - day_offset[row.week_day]
                )
                width = day_width / (max_width)
                offset = width * (row["offset"])
                time = row["Startzeitpunkt"].time()
                trace = go.Bar(
                    x=[row["plot_day"]],
                    y=[row["diff"]],
                    base=[time.second + 60 * time.minute + 60**2 * time.hour],
                    marker_color=row["color"],
                    name=row["Ort"],
                    width=width,
                    offset=offset,
                    legendgroup=row["Ort"],
                    showlegend=show_legend[row["Ort"]],
                )
                if show_legend[row["Ort"]]:
                    show_legend[row["Ort"]] = False
                fig.add_trace(trace)
                font = "arial"
                fs_df["next_size"] = fs_df["Dauer"].shift(-1)
                duration = np.divide(row["diff"], 3600)
                px_font_size = int(
                    fs_df.loc[
                        (fs_df["Dauer"] <= duration)
                        & (
                            (fs_df["next_size"] > duration)
                            ^ (pd.isna(fs_df["next_size"]))
                        ),
                        "Schriftgröße",
                    ].values[0]
                )
                add_annotations(
                    row, fig, font, px_font_size, max_width, day_width
                )
        fig.update_layout(
            margin=dict(r=50, l=50, b=50, t=125),
            title={
                "text": "Helpdeskplan für KW 13",
                "y": 0.9,
                "x": 0.5,
                "xanchor": "center",
                "yanchor": "top",
            },
        )
        fig.add_layout_image(
            source="https://upload.wikimedia.org/wikipedia/commons/thumb/5/53/Logo-Hochschule-RheinMain.svg/2560px-Logo-Hochschule-RheinMain.svg.png",
            xref="paper",
            yref="paper",
            x=1,
            y=1,
            sizex=0.6,
            sizey=0.6,
            xanchor="right",
            yanchor="bottom",
        )

        return fig


@callback(Output("cw-output", "children"), Input("date-picker", "value"))
def date_to_cw(date_str: str) -> str:
    date = pd.to_datetime(date_str, yearfirst=True, format="%Y-%m-%d")
    cw_str = date.strftime("%V %G")
    return cw_str
