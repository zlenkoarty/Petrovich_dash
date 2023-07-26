import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table
from dash.dependencies import Input, Output
import pandas as pd


df = pd.read_excel('turnover.xlsx')
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

    1. **Компания "ООО ПЕТРОВИЧ-ТЕХ":**
       - Несмотря на то, что у нас отсутствует детализация по группам должностей для этой компании, но факт стабильности, когда нет текучести кадров на протяжении длительного времени, может говорить о хорошем управлении и условиях труда.

    2. **Компания "ООО СТД Петрович":** 
       - Компания показывает более сложную картину. Есть определённые группы должностей и подразделения, где текучесть персонала превышает средние значения.

    3. **РЦ "Пискаревский":** 
       - Особенно в этом подразделении внимание привлекает высокая текучесть среди грузчиков. Можно предположить, что условия труда, возможно, нагрузка или оплата труда в этой должности не устраивают сотрудников, поэтому они часто уходят.
       
    4. **Грузчики в других подразделениях:** 
       - Сравнение показателей текучести для грузчиков в РЦ "Новосаратовка" и РЦ "Салова" показывает, что проблемы, возможно, связаны не только с самой должностью, но и с условиями в конкретном подразделении.
       
    5. **Женщины-кладовщики:** 
       - Стабильность в этих должностях может говорить о том, что для женщин условия труда более приемлемы или, возможно, менее существует альтернатив на рынке труда.

    6. **Несоответствие общей текучести и текучести по должностям:** 
       - Это может говорить о том, что в этих подразделениях имеются должности с низкой текучкой, которые компенсируют высокую текучесть на других должностях. Например, в руководящих или административных должностях текучесть может быть ниже.
    
    7. **Центральный офис (ЦО):** 
       - Видим, что уровень текучести примерно одинаковый для мужчин и женщин. Это может говорить о том, что условия труда и возможности для развития представлены равноценно для обоих полов.

    В заключение, для "ООО СТД Петрович" имеет смысл провести дополнительные исследования и опросы сотрудников для выявления конкретных причин текучести и разработки стратегии по удержанию персонала. Например, улучшение условий труда, повышение заработной платы, проведение карьерных консультаций и тренингов для развития.
    

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
    app.run_server(debug=False)
