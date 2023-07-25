import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import pandas as pd

df = pd.read_excel('turnover.xlsx')
pivot_table = df.pivot_table(
    values=['Общая текучесть','Текучесть 3 месяца', 'Текучесть 6 месяца', 'Текучесть 12 месяца'],
    index=['Организация', 'Группа подразделений', 'Группа должностей', 'Пол'],
    aggfunc='mean'
) * 100

app = dash.Dash(__name__, external_stylesheets=['https://codepen.io/chriddyp/pen/bWLwgP.css'])
server = app.server
filter_options = ['Организация', 'Группа подразделений', 'Группа должностей', 'Пол']

app.layout = html.Div([
    html.H1("Текучесть кадров"),

    dcc.Dropdown(
        id='filter-dropdown',
        options=[{'label': i, 'value': i} for i in filter_options],
        multi=False,
        value='Группа подразделений',  # Set default filter
    ),
    
    html.Div(children=[
        html.Div(dcc.Graph(id='total-turnover-graph'), className="six columns"),
        html.Div(dcc.Graph(id='3m-turnover-graph'), className="six columns"),
    ], className="row"),

    html.Div(children=[
        html.Div(dcc.Graph(id='6m-turnover-graph'), className="six columns"),
        html.Div(dcc.Graph(id='12m-turnover-graph'), className="six columns"),
    ], className="row")

], className="ten columns offset-by-one")

@app.callback(
    Output('total-turnover-graph', 'figure'),
    Output('3m-turnover-graph', 'figure'),
    Output('6m-turnover-graph', 'figure'),
    Output('12m-turnover-graph', 'figure'),
    Input('filter-dropdown', 'value')
)
def update_graphs(selected_filter):
    graphs = []

    for turnover_metric in ['Общая текучесть', 'Текучесть 3 месяца', 'Текучесть 6 месяца', 'Текучесть 12 месяца']:
        if selected_filter:
            filtered_data = pivot_table.reset_index()
            filtered_data = filtered_data[filtered_data[selected_filter].notna()]
            unique_names = filtered_data[selected_filter].unique()

            graph = {
                'data': [
                    {
                        'x': [name], 
                        'y': [filtered_data.loc[filtered_data[selected_filter] == name, turnover_metric].mean()],
                        'type': 'bar', 
                        'name': name,
                        'text': [f'{filtered_data.loc[filtered_data[selected_filter] == name, turnover_metric].mean():.2f}%'],
                        'hovertext': [name],  # hover text same as legend
                        'hoverinfo': 'text',  # display only the hover text
                        'textposition': 'outside',
                    }
                    for name in unique_names
                ],
                'layout': {
                    'title': turnover_metric,
                    'xaxis': {'visible': False, 'showticklabels': False},  # Hide x-axis
                    'yaxis': {'visible': False, 'showticklabels': False},  # Hide y-axis
                    'margin': {'t': 50},  # Add a top margin
                }
            }
            graphs.append(graph)
        else:
            graphs.append({'data': [], 'layout': {'title': turnover_metric, 'xaxis': {'visible': False, 'showticklabels': False}, 'yaxis': {'visible': False, 'showticklabels': False}}})

    return graphs[0], graphs[1], graphs[2], graphs[3]

if __name__ == '__main__':
    app.run_server(debug=False)
