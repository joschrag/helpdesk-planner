import dash
import dash_bootstrap_components as dbc
from dash import dcc, html
import dash_app.callbacks
app = dash.Dash(
    __name__,
    use_pages=True,
    external_stylesheets=[dbc.themes.SLATE],
)

app.layout = dbc.Container(
    children=[
        dbc.Row(html.Div(dash.page_container)),
        dcc.Interval(id="10_min", interval=1000*10*60),
    ],
    className="dbc",
    fluid=True,
)