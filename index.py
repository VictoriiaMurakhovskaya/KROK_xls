import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
import dash
import dash_daq as daq
from main import installed_fonts, make_file
from dash.dependencies import Input, Output, State
import os
import xlsxwriter
import flask

fonts = installed_fonts()
file_name = None
link = None


def base_layout():
    main_controls = dbc.Card([
        dbc.FormGroup([
            dbc.Label('Тип шрифта'),
            dcc.Dropdown(id='main_font', options=fonts, value=1)
        ]),
        dbc.FormGroup([
            dbc.Label('Размер шрифта'),
            daq.NumericInput(id='font_size', value=9, max=14)
        ]),
        dcc.Upload(id='input_file', children=[dbc.Button('Загрузить файл', color='primary')])
    ], style={'width': '12rem', 'height': '16rem'}, body=True)

    width_controls_1 = dbc.Card([
        dbc.Label('Колонка 1'),
        daq.NumericInput(id='col_1', value=10, max=20),
        dbc.FormGroup([
            dbc.Label('Колонка 2'),
            daq.NumericInput(id='col_2', value=40, max=100)
        ]),
        dbc.FormGroup([
            dbc.Label('Колонка 3'),
            daq.NumericInput(id='col_3', value=10, max=20)
        ]),
    ], style={'width': '8rem', 'height': '16rem'}, body=True)

    width_controls_2 = dbc.Card([
        dbc.Label('Колонка 4'),
        daq.NumericInput(id='col_4', value=8, max=15),
        dbc.FormGroup([
            dbc.Label('Колонка 5'),
            daq.NumericInput(id='col_5', value=8, max=15)
        ]),
        dbc.FormGroup([
            dbc.Label('Колонка 6'),
            daq.NumericInput(id='col_6', value=8, max=15)
        ]),
    ], style={'width': '8rem', 'height': '16rem'}, body=True)

    layout = dbc.Container([
        dbc.Row([
            dbc.Col([
                html.H1('Индивидуальные листы оценок')
            ], md=5, width='auto')
        ]),
        dbc.Row([
            dbc.Col([
                html.Hr()
            ], md=5, style={'align': 'center'})
        ]),
        dbc.Row([
            dbc.Col([
                main_controls
            ], md=2),
            dbc.Col([
                width_controls_1
            ], width='auto'),
            dbc.Col([
                width_controls_2
            ], md=2)
        ], style={'justify': 'around'}),
        dbc.Row([
            dbc.Col([
                html.Div(id='download_file', children=['Файл не обработан'])
            ], md=6, width={'offset': 3}, style={'justify': 'center', 'margin': '14px'})
        ])

    ], fluid=True)
    return layout


app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.config.suppress_callback_exceptions = False
app.layout = html.Div([dcc.Location(id='loc', refresh=True),
                       html.Div(id='page-content', children=[base_layout()]),
                       ])


@app.callback(Output('download_file', 'children'),
              [Input('input_file', 'filename')],
              [State('col_1', 'value'),
               State('col_2', 'value'),
               State('col_3', 'value'),
               State('col_4', 'value'),
               State('col_5', 'value'),
               State('col_6', 'value'),
               State('main_font', 'value'),
               State('font_size', 'value')]
              )
def input_file(name, w1, w2, w3, w4, w5, w6, font, size):
    global file_name
    if name:
        try:
            file_name = name
            wb = xlsxwriter.Workbook('result.xlsx', {'in_memory': True})
            wb = make_file(name, wb, w1, w2, w3, w4, w5, w6, fonts[font], size)
            wb.close()
            link = html.A('Загрузить результат',
                          href='/result.xlsx', download='result.xlsx')
        except:
            link = html.H3('Неудачная обработка файла')
    else:
        link = None
    return link


@app.server.route("/result.xlsx")
def serve_static():
    return flask.send_file(os.getcwd()+'/result.xlsx',
                           mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == '__main__':
    app.run_server(debug=True)
