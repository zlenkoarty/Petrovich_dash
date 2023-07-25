import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table
from dash.dependencies import Input, Output
import pandas as pd

df = pd.read_excel('текучесть.xlsx')
pivot_table = df.pivot_table(
    values=['Общая текучесть','Текучесть 3 месяца', 'Текучесть 6 месяца', 'Текучесть 12 месяца'],
    index=['Организация', 'Группа подразделений', 'Группа должностей', 'Пол'],
    aggfunc='mean'
) * 100

# Round and format DataFrame for the table
formatted_table = pivot_table.reset_index()
for col in ['Общая текучесть','Текучесть 3 месяца', 'Текучесть 6 месяца', 'Текучесть 12 месяца']:
    formatted_table[col] = formatted_table[col].round(2).apply(lambda x: f'{x}%')

# Define column order
column_order = ['Организация', 'Группа подразделений', 'Группа должностей', 'Пол', 'Общая текучесть', 'Текучесть 3 месяца', 'Текучесть 6 месяца', 'Текучесть 12 месяца']

# Reorder columns
formatted_table = formatted_table.reindex(columns=column_order)

# Write DataFrame to an excel
formatted_table.to_excel('formatted_table.xlsx', index=False)

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
    ], className="row"),

    html.Hr(),  # Horizontal line

    #html.A('Скачать данные', href='/download_excel/', download="formatted_table.xlsx", id='download-link'),

   
    html.Div(dcc.Markdown('''
    **Анализ текучести кадров**

    1. **Общая текучесть:**
       - Общая текучесть для всей организации "ООО ПЕТРОВИЧ-ТЕХ" составляет 0,0%. Это означает, что за наблюдаемый период сотрудники не покидали компанию.

    2. **Текучесть по полу:**
       - Во всех подразделениях и группах должностей текучесть как для мужчин, так и для женщин относительно низкая и составляет 0,0%. Это говорит о том, что компания успешно удерживает сотрудников обоих полов.

    3. **Текучесть по подразделениям:**
       - В подразделении "ГО" текучесть для мужчин составляет 20,0%, что является наивысшим показателем среди всех подразделений. Однако для женщин в этом подразделении оттока нет.

    4. **Текучесть по группам должностей:**
       - Для группы должностей "Грузчик" в подразделении "РЦ Пискаревский" наблюдается наивысшая текучесть среди всех должностей и составляет 46,15% за 12 месяцев. Это указывает на возможные проблемы с удержанием сотрудников на этой должности.

    5. **Текучесть по подразделениям и полу:**
       - В подразделении "КЦ" текучесть как для мужчин (35,97%), так и для женщин (31,74%) относительно высокая по сравнению с другими подразделениями.

    6. **Текучесть по подразделениям и группам должностей:**
       - В подразделении "РЦ Новосаратовка" наблюдается значительная текучесть в различных группах должностей, особенно для мужчин на должностях "Кладовщики площадка" (22,22%) и "Комплектовщики САХ" (36,07%).

    В целом, компания имеет низкую текучесть, что является положительным сигналом и указывает на стабильность трудового коллектива. Однако есть конкретные подразделения и должности, где текучесть относительно выше, особенно среди мужчин. Для разработки соответствующих стратегий удержания необходимо провести дополнительный анализ и исследование причин текучести в этих областях. Следует отметить, что показатели текучести могут меняться со временем, поэтому непрерывный мониторинг и анализ являются важными для решения потенциальных проблем удержания сотрудников.
    
    Ниже представленна таблица, лежащей в основе построяния графиков. Скачать таблицу можно, нажав кнопку "EXPORT"
    '''), className="row"),

    html.Hr(),
    html.Div(
        dash_table.DataTable(
            id='table',
            columns=[{"name": i, "id": i} for i in formatted_table.columns],
            data=formatted_table.to_dict('records'),
            export_format="xlsx"
        ),
        className="row"
    ),
    html.Hr(),
    

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
    app.run_server(debug=True)
