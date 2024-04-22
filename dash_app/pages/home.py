"""
This module contains the layout for the home page.

Author: Jonas Schrage
Date: 15.04.2023

"""
import dash
import dash_bootstrap_components as dbc
from dash import html, dash_table, dcc
import dash_mantine_components as dmc
from datetime import datetime, date

dash.register_page(
    __name__,
    path="/",
    name="Helpdesk Planner",
    title="Helpdesk Planner",
    description="Landing page.",  # image="img_home.png"
)

date_picker = dbc.Row(
    [
        dbc.Col(
            dmc.DatePicker(
                id="date-picker",
                label="Start Date",
                description="Wähle das Wochendatum aus:",
                minDate=date(2020, 8, 5),
                value=datetime.now().date(),
                style={"width": 200},
            ),
            width=2,
        ),
        dbc.Col(
            [html.P(["Kalenderwoche: ", html.Span(id="cw-output")])], width=4
        ),
    ]
)

fs_table = dash_table.DataTable(
    id="font_sizes", editable=True, row_deletable=True
)

col_btn2 = dbc.Button(
            "Open font size control",
            id="collapse-button-fs",
            className="mb-3",
            color="primary",
            n_clicks=0,
        ),
collapse2 = html.Div(
    [
        
        dbc.Collapse(
            [dbc.Row(fs_table),dbc.Row(
            dbc.Col(
                html.Button("Add Row", id="editing-rows-button-fs", n_clicks=0),
                width=3,
            )
        ),],
            id="collapse-fs",
            is_open=False,
        ),
    ]
)

data_table = dash_table.DataTable(
    id="timetable", editable=True, row_deletable=True
)

color_btn_link = dbc.Button("Link to available colors",href='https://developer.mozilla.org/en-US/docs/Web/CSS/named-color')

layout = dbc.Container(
    [
        dbc.Row(html.H1("Home")),
        date_picker,
        dbc.Row(data_table),
        dbc.Row(
            dbc.Col(
                html.Button("Add Row", id="editing-rows-button", n_clicks=0),
                width=3,
            )
        ),
        dbc.Row([dbc.Col(col_btn2),dbc.Col(color_btn_link)]),
        dbc.Row(collapse2),
        dbc.Row(
            dbc.Col(
                dcc.Dropdown(
                    [
                        {"value": 5, "label": "FR"},
                        {"value": 6, "label": "SA"},
                        {"value": 7, "label": "SO"},
                    ],
                    id="end_day_dd",
                    placeholder="Choose the last day of the plan...",
                )
            )
        ),
        dbc.Row(id="day_width"),
        dbc.Row(
            dbc.Col(
                html.Button("Create Plan", id="create-plan", n_clicks=0),
                width=3,
            )
        ),
        dbc.Row(dcc.Graph(id="tt-graph")),
    ],
    fluid=True,
)
