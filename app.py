import dash
from dash import Dash, dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import io
import base64
import os
import plotly.io as pio
import kaleido  # required by plotly.io.write_image
from datetime import datetime

external_stylesheets = [
    dbc.themes.FLATLY,
    "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css"
]


# Fallback sample data in case Excel is missing or empty
fallback_data = pd.DataFrame({
    "Name of IRO": [
        "Energy Efficiency", "Water Use", "Labor Practices", "GHG Emissions",
        "Diversity & Inclusion", "Product Safety", "Supply Chain Ethics",
        "Board Independence", "Climate Risk Strategy", "Customer Privacy"
    ],
    "Impact": [4, 3, 5, 5, 2, 4, 3, 1, 5, 2],
    "Risk":    [5, 2, 4, 5, 3, 3, 2, 1, 5, 4],
    "Sub-Topic": [
        "Environmental", "Environmental", "Social", "Environmental",
        "Social", "Social", "Social", "Governance", "Environmental", "Governance"
    ]
})

# Try loading the default template
TEMPLATE_PATH = os.path.join("data", "Materiality_Template.xlsx")
try:
    df_default = pd.read_excel(TEMPLATE_PATH)
    if df_default.empty or not {"Name of IRO", "Impact", "Risk", "Sub-Topic"}.issubset(df_default.columns):
        df_default = fallback_data
except Exception as e:
    print(f"Failed to load Excel file: {e}")
    df_default = fallback_data

# Function to generate the plot
def generate_figure(df, company_name=""):

    if "Sub-Topic" in df.columns:
        def wrap_label(text):
            clean = str(text).replace("_", " ")
            words = clean.split()
            # break into lines of 4 words
            lines = [' '.join(words[i:i + 4]) for i in range(0, len(words), 4)]
            return '<br>'.join(lines)

        df["Sub-Topic"] = df["Sub-Topic"].apply(wrap_label)

    title_text = f"<b>{company_name} : Double Materiality Map</b>" if company_name else "<b>Double Materiality Map</b>"

    fig = px.scatter(
        df,
        x="Risk",
        y="Impact",
        color="Sub-Topic",
        hover_name="Name of IRO",
        title=title_text
    )

    fig.update_traces(marker=dict(
        size=12,
        line=dict(width=0),  # no border
        opacity=1.0
    ), selector=dict(mode='markers'))

    fig.update_layout(
        template='simple_white',
        autosize=False,
        height=620,
        width=850,
        uirevision='static',
        margin=dict(l=40, r=40, t=60, b=60),
        font=dict(size=18, family="Arial", color="#333"),
        title_font=dict(size=20, family="Arial", color="#333"),
        title_x=0.25,  # Center the title relative to the graph
        legend=dict(
            font=dict(size=18),
            title=dict(text="<b>Sub-Topic<br>", font=dict(size=19), side="top"),
            x=1.1,
            xanchor='left',
            y=1.0,
            yanchor='top'
        ),
        xaxis=dict(
            title="Financial Materiality (Risk or Opportunity)",
            range=[0, 5.1],         # lock x-axis
            fixedrange=True,      # no zooming or panning
            tickmode='array',
            tickvals=[1, 2, 3, 4, 5],  # Show only 1–5
            ticktext=["1", "2", "3", "4", "5"],
            gridcolor='lightgray',
            gridwidth=0.5,
            showgrid=True,
            scaleanchor="y",
            layer='below traces'
        ),
        yaxis=dict(
            title="Impact Materiality",
            range=[0, 5.1],         # lock y-axis
            fixedrange=True,      # no zooming or panning
            tickmode='array',
            tickvals=[1, 2, 3, 4, 5],  # Show only 1–5
            ticktext=["1", "2", "3", "4", "5"],
            gridcolor='lightgray',
            gridwidth=1,
            showgrid=True,
            layer='below traces'
        )
    )

    return fig


# Dash app init
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
app.title = "DMM"
app.config.suppress_callback_exceptions = True

# Layout
app.layout = dbc.Container([
    html.Img(src='/assets/bm.jpeg', className='logo'),

    html.H2("Bemari : Double Materiality Map", className="mt-4 mb-4 text-center"),

    html.Div(style={"marginBottom": "60px"}),  # Add vertical space here

    html.Div(
        dbc.Row([
            # Drag and Drop Upload
            dbc.Col(
                html.Div([
                    html.Label("Plot YOUR data on this graph", style={
                        "fontWeight": "bold",
                        "marginBottom": "4px"
                    }),
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            html.Span('Drag and Drop  /  ', style={"fontSize": "22px"}),
                            html.A('Select Excel File (.xlsx)', style={"fontSize": "22px"})
                        ]),
                        style={
                            'width': '500px',
                            'height': '45px',
                            'lineHeight': '45px',
                            'borderWidth': '0.1px',
                            'borderStyle': 'dashed',
                            'borderRadius': '5px',
                            'textAlign': 'center',
                            'padding': '0px 10px'
                        },
                        accept=".xlsx"
                    )
                ]),
                width="auto",
                style={"marginRight": "60px", "minWidth": "340px"}
            ),
            # Company Name Input
            dbc.Col(
                html.Div([
                    html.Label("Company Name", style={"fontWeight": "bold", "display": "block", "marginBottom": "4px"}),
                    dbc.Input(
                        id="input-company",
                        type="text",
                        placeholder="Enter company name (optional)",
                        style={
                            "minWidth": "250px",
                            "border": "0.1px solid #555",  # darker, thicker border
                            "borderRadius": "5px",  # optional: round corners
                            "boxShadow": "0 0 2px #888888"  # subtle glow (optional)
                        }
                    )
                ]),
                style={"marginRight": "60px"}
            ),

            # Download Template Button
            dbc.Col(
                html.Div([
                    html.Label(" ", style={"display": "block", "marginBottom": "6px"}),  # For vertical alignment
                    html.A([
                        html.I(className="fas fa-file-excel", style={"marginRight": "15px", "fontSize": "22px"}),
                        "Download template input file"
                    ],
                        href="/static/Materiality_Template.xlsx",
                        target="_blank",
                        className="btn btn-outline-primary"
                    )
                ]),
                style={"minWidth": "250px", "marginRight": "10px"}
            ),
            # Download Chart Button
            dbc.Col(
                html.Div([
                    html.Label(" ", style={"display": "block", "marginBottom": "4px"}),
                    html.Button([
                        html.I(className="fas fa-download", style={"marginRight": "15px", "fontSize": "22px"}),
                        "Download this Chart"
                    ],
                        id="btn-download-fig",
                        style={
                            "backgroundColor": "white",
                            "color": "black",
                            "border": "1px solid #ccc",
                            "padding": "6px 12px",
                            "borderRadius": "5px",
                            "fontWeight": "500"
                        }),
                    dcc.Download(id="download-fig")
                ])
            )
        ],
            className="mb-4",
            align="end",
            justify="start"  # This ensures left alignment of the row contents
        )
    ),

    html.Div([
        html.Div(
            dcc.Graph(id='scatter-plot', figure=generate_figure(df_default),
                      style={"height": "100%", "width": "100%", "marginTop": "60px"}),
            className="graph-container"
        )
    ], className="graph-wrapper")
], fluid=True)

# Callback for file upload
@app.callback(
    Output('scatter-plot', 'figure'),
    Input('upload-data', 'contents'),
    Input('input-company', 'value'),
    State('upload-data', 'filename')
)

def update_graph(contents, company_name, filename):
    if contents is None:
        return generate_figure(df_default, company_name or "")

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        df = pd.read_excel(io.BytesIO(decoded))
        required_cols = {"Name of IRO", "Impact", "Risk", "Sub-Topic"}
        if not required_cols.issubset(df.columns):
            return px.scatter(title="Missing required columns: Name of IRO, Impact, Risk, Sub-Topic")
        return generate_figure(df, company_name or "")
    except Exception as e:
        return px.scatter(title=f"Error reading file: {str(e)}")

@app.callback(
    Output("download-fig", "data"),
    Input("btn-download-fig", "n_clicks"),
    State("input-company", "value"),
    prevent_initial_call=True
)
def download_chart(n_clicks, company_name):
    fig = generate_figure(df_default, company_name or "")
    filename = f"materiality_map_{(company_name or 'company').lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.png"
    return dcc.send_bytes(lambda: fig.to_image(format="png", width=900, height=600, scale=2), filename)


test_fig = generate_figure(df_default)
# test_fig.write_image("test_plot.png")

# if __name__ == '__main__':
#     app.run(debug=True)

server = app.server