
import os
import numpy as np
import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px
import webbrowser


months = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']


df = pd.read_excel("data.xlsx")
print("Data loaded from 'data.xlsx' successfully!")

category_sales = df.groupby("Category")["Sales"].sum().reset_index()

monthly_sales = (
    df.groupby("Month", as_index=False)["Sales"]
    .sum()
    .set_index("Month")
    .reindex(months)
    .fillna(0)  
    .reset_index()
)


heatmap_data = pd.pivot_table(df, values='Sales', index='Category', columns='Month', aggfunc=np.sum)

app = dash.Dash(__name__)


app.layout = html.Div([
    html.H1("Excel Data Visualization Dashboard", style={'text-align': 'center', 'margin-bottom': '20px'}),

 
    html.Div([
        html.Label("Select a Category:", style={'font-size': '18px'}),
        dcc.Dropdown(
            id='category-dropdown',
            options=[{'label': 'All Categories', 'value': 'all'}] +
                    [{'label': category, 'value': category} for category in df['Category'].unique()],
            value='all',  
            style={'width': '50%', 'margin-bottom': '20px'}
        ),
    ], style={'text-align': 'center', 'margin-bottom': '20px'}),

    html.Div(id='log-output', style={'text-align': 'center', 'font-size': '16px', 'margin-bottom': '20px'}),

    
    dcc.Graph(id='bar-chart'),

   
    dcc.Graph(id='line-chart'),

    
    dcc.Graph(id='scatter-plot'),

   
    dcc.Graph(id='pie-chart'),

   
    dcc.Graph(id='heatmap'),

    
    dcc.Graph(id='box-plot'),
])


@app.callback(
    [
        Output('bar-chart', 'figure'),
        Output('line-chart', 'figure'),
        Output('scatter-plot', 'figure'),
        Output('pie-chart', 'figure'),
        Output('heatmap', 'figure'),
        Output('box-plot', 'figure'),
        Output('log-output', 'children')
    ],
    [Input('category-dropdown', 'value')]
)
def update_charts(selected_category):
   
    if selected_category == 'all':
        filtered_df = df
        message = "Showing data for all categories."
    else:
        filtered_df = df[df['Category'] == selected_category]
        message = f"Showing data for category: '{selected_category}'"

  
    category_sales_filtered = filtered_df.groupby("Category")["Sales"].sum().reset_index()
    bar_chart = px.bar(category_sales_filtered, x="Category", y="Sales", title="Total Sales by Category", color="Category")

   
    monthly_sales_filtered = (
        filtered_df.groupby("Month", as_index=False)["Sales"]
        .sum()
        .set_index("Month")
        .reindex(months)
        .fillna(0)
        .reset_index()
    )
    line_chart = px.line(monthly_sales_filtered, x="Month", y="Sales", title="Monthly Sales Trend", markers=True)

   
    scatter_plot = px.scatter(filtered_df, x="Sales", y="Profit", color="Category",
                              title="Sales vs Profit (Colored by Category)", opacity=0.6)

  
    pie_chart = px.pie(monthly_sales_filtered, names="Month", values="Sales", title="Sales Distribution by Month")

  
    heatmap_data_filtered = pd.pivot_table(filtered_df, values='Sales', index='Category', columns='Month', aggfunc=np.sum)
    heatmap_chart = px.imshow(heatmap_data_filtered, labels=dict(x="Month", y="Category", color="Sales"),
                               title="Sales Heatmap by Category and Month", aspect="auto")

    
    box_chart = px.box(filtered_df, x='Category', y='Sales', title="Box Plot - Sales Distribution by Category")

    return bar_chart, line_chart, scatter_plot, pie_chart, heatmap_chart, box_chart, message


if __name__ == '__main__':
    url = "http://127.0.0.1:8050/"
    print(f"Starting Dash app... Click to open: {url}")

   
    if os.environ.get("RENDER_ENV") != "production":
        webbrowser.open_new(url)


    app.run_server(host='0.0.0.0', port=int(os.environ.get("PORT", 8050)))
