import pandas as pd
from dash.dependencies import Input, Output, State
from dash import Dash, html, Input, Output,dcc,dash_table
import dash_bootstrap_components as dbc
from dash.exceptions import PreventUpdate
from dash_iconify import DashIconify
import plotly.express as px

app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.BOOTSTRAP], 
                meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1.0", }],
                )

#Load Image
image_path = 'assets/survey.jpg'


### --- LOAD DATAFRAME
excel_file = 'Survey_Results.xlsx'
sheet_name = 'DATA'

df = pd.read_excel(excel_file,
                   sheet_name=sheet_name,
                   usecols='B:D',
                   header=3)

df_participants = pd.read_excel(excel_file,
                                sheet_name= sheet_name,
                                usecols='F:G',
                                header=3)
df_participants.dropna(inplace=True)

# --- STREAMLIT SELECTION
department = df['Department'].unique().tolist()
ages = df['Age'].unique().tolist()

age_selection = dcc.RangeSlider(min(ages), max(ages), id='age',  value=[min(ages), max(ages)])

department_selection = dcc.Dropdown( department, value=[department[0], department[1]], id='department',
                                    multi=True)



app.layout=dbc.Container([dbc.Row([dbc.Col([html.P('Survey Results 2021', className='title'), html.P('Redesign Data App With Dash Plotly (YT : @walidata1996, Tiktok: @walidata18)', className='subtitle'),   dbc.Col(html.Img(src=image_path, className='img')
)])]),
                          dbc.Row([
    dbc.Col([html.P('Select Range Age'), age_selection, html.P('Select Department'), department_selection], className='card',),
    
    dbc.Col(dash_table.DataTable(id='table', style_cell={'textAlign': 'left'}, style_data={"font-family":'Segoe UI'}, style_header={'color':'#035064',"font-family":'Segoe UI','backgroundColor': '#d1f3fd', 'fontWeight':'bold'}),  
            style={'height':'200px', 'overflowY': 'auto'}, className='card'),
  

], className='g-4'),
dbc.Row([
    dbc.Col(dcc.Graph(id='graph1', ), className='card'),
    dbc.Col(dcc.Graph(id='graph2', ), className='card'),
    ]), 
    dbc.Row([dbc.Col(html.Div(className='spinner'))
        ,dbc.Col(html.P('I have got this inspiration from his streamlit work : https://github.com/Sven-Bo/excel-webapp-streamlit'), )])

])

@app.callback(Output('graph1', 'figure'), Output('graph2', 'figure'), Output('table', 'data'), Input('age', 'value'), Input('department', 'value'))
def graph(age, department):
    # --- FILTER DATAFRAME BASED ON SELECTION
    mask = (df['Age'].between(*age)) & (df['Department'].isin(department))
    number_of_result = df[mask].shape[0]

    # --- GROUP DATAFRAME AFTER SELECTION
    df_grouped = df[mask].groupby(by=['Rating']).count()[['Age']]
    df_grouped = df_grouped.rename(columns={'Age': 'Votes'})
    df_grouped = df_grouped.reset_index()

    bar_chart = px.bar(df_grouped,
                    title='Rating vs Votes',
                    x='Rating',
                    y='Votes',
                    text='Votes',
                    #color_discrete_sequence = ['#F63366']*len(df_grouped),
                    )
    
    bar_chart.update_layout(paper_bgcolor = 'rgba(0,0,0,0)',
                      plot_bgcolor = 'rgba(0,0,0,0)',
                      title_x=0.5,
                      bargap=0.1,
                      title_font_size = 20,
                      title_font_family = 'Segoe UI'),

    bar_chart.update_xaxes(gridcolor='#ECF2FF', minor_griddash="dot")
    bar_chart.update_traces(marker_color='#035064')

    pie_chart = px.pie(df_participants,
                    title='Total No. of Participants',
                    values='Participants',
                    names='Departments',
                    color_discrete_sequence=px.colors.sequential.Blues_r)
    pie_chart.update_layout(paper_bgcolor = 'rgba(0,0,0,0)',
                      plot_bgcolor = 'rgba(0,0,0,0)',
                      title_x=0.5,
                      bargap=0.1,
                      title_font_size = 20,
                      title_font_family = 'Segoe UI'),
    
    table = df[mask].to_dict('records')
    
    return bar_chart, pie_chart, table

if __name__ == '__main__':
    app.run_server(host='0.0.0.0', debug=True,)