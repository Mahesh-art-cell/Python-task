

import os
import pandas as pd
import numpy as np
import dash
from dash import dcc, html
import plotly.express as px
import webbrowser


np.random.seed(42)

categories = ['Electronics', 'Clothing', 'Furniture', 'Food', 'Toys']
months = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']


data = {
    'Category': np.random.choice(categories, 1000),
    'Month': np.random.choice(months, 1000),
    'Sales': np.random.randint(500, 5000, 1000),
    'Profit': np.random.randint(50, 1000, 1000)
}


df = pd.DataFrame(data)


df.to_excel("data.xlsx", index=False, engine='openpyxl')

print(" Data saved to 'data.xlsx' successfully!")


df = pd.read_excel("data.xlsx")
print(" Data loaded from 'data.xlsx' successfully!")


category_sales = df.groupby("Category")["Sales"].sum().reset_index()

monthly_sales = (
    df.groupby("Month", as_index=False)["Sales"]
    .sum()
    .set_index("Month")
    .reindex(months)
    .reset_index()
)

# Scatter Plot Data
scatter_data = df

# Heatmap Data Preparation
heatmap_data = pd.pivot_table(df, values='Sales', index='Category', columns='Month', aggfunc=np.sum)

# Box Plot Data Preparation
box_data = df[['Category', 'Sales']]



# Initialize Dash app
app = dash.Dash(__name__)

# Define Layout for Dash Dashboard
app.layout = html.Div([
    html.H1(" Excel Data Visualization Dashboard", style={'text-align': 'center', 'margin-bottom': '20px'}),

    # Bar Chart - Total Sales by Category
    dcc.Graph(
        id='bar-chart',
        figure=px.bar(category_sales, x="Category", y="Sales", title="Total Sales by Category", color="Category")
    ),

    # Line Chart - Monthly Sales Trend
    dcc.Graph(
        id='line-chart',
        figure=px.line(monthly_sales, x="Month", y="Sales", title="Monthly Sales Trend", markers=True)
    ),

    # Scatter Plot - Sales vs Profit with Color by Category
    dcc.Graph(
        id='scatter-plot',
        figure=px.scatter(scatter_data, x="Sales", y="Profit", color="Category",
                          title="Sales vs Profit (Colored by Category)", opacity=0.6)
    ),

    # Pie Chart - Sales Distribution by Month
    dcc.Graph(
        id='pie-chart',
        figure=px.pie(monthly_sales, names="Month", values="Sales", title="Sales Distribution by Month")
    ),

    # Heatmap - Sales by Category and Month
    dcc.Graph(
        id='heatmap',
        figure=px.imshow(heatmap_data, labels=dict(x="Month", y="Category", color="Sales"),
                         title="Sales Heatmap by Category and Month", aspect="auto")
    ),

    # Box Plot - Sales Distribution by Category
    dcc.Graph(
        id='box-plot',
        figure=px.box(box_data, x='Category', y='Sales', title="Box Plot - Sales Distribution by Category")
    ),
])




if __name__ == '__main__':
    url = "http://127.0.0.1:8050/"
    print(f"Starting Dash app... Click to open: {url}")

    # Automatically open the browser in development (but not in production on Render)
    if os.environ.get("RENDER_ENV") != "production":
        webbrowser.open_new(url)

    # Run the Dash server with 0.0.0.0 for Render deployment
    app.run_server(host='0.0.0.0', port=int(os.environ.get("PORT", 8050)))


