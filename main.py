# import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt
# import seaborn as sns
# import dash
# from dash import dcc, html
# import plotly.express as px

# # 📌 Step 1: Generate Sample Data and Save to Excel
# np.random.seed(42)

# categories = ['Electronics', 'Clothing', 'Furniture', 'Food', 'Toys']
# months = ['January', 'February', 'March', 'April', 'May', 'June', 
#           'July', 'August', 'September', 'October', 'November', 'December']

# data = {
#     'Category': np.random.choice(categories, 1000),
#     'Month': np.random.choice(months, 1000),
#     'Sales': np.random.randint(500, 5000, 1000),
#     'Profit': np.random.randint(50, 1000, 1000)
# }

# df = pd.DataFrame(data)
# # df.to_excel("data.xlsx", index=False)
# df.to_excel("data.xlsx", index=False, engine='openpyxl')


# # 📌 Step 2: Read the Excel File
# df = pd.read_excel("data.xlsx")

# # 📌 Step 3: Generate Visualizations

# # 1️⃣ Bar Chart - Total Sales by Category
# plt.figure(figsize=(10,5))
# sns.barplot(x=df['Category'], y=df['Sales'], estimator=sum, ci=None, palette="viridis")
# plt.title("Total Sales by Category")
# plt.xlabel("Category")
# plt.ylabel("Total Sales")
# plt.xticks(rotation=45)
# plt.show()

# # 2️⃣ Line Graph - Monthly Sales Trend
# plt.figure(figsize=(10,5))
# monthly_sales = df.groupby("Month")["Sales"].sum().reindex(months)
# sns.lineplot(x=monthly_sales.index, y=monthly_sales.values, marker='o', color='b')
# plt.title("Monthly Sales Trend")
# plt.xlabel("Month")
# plt.ylabel("Total Sales")
# plt.xticks(rotation=45)
# plt.grid()
# plt.show()

# # 3️⃣ Scatter Plot - Sales vs Profit
# plt.figure(figsize=(8,5))
# sns.scatterplot(x=df["Sales"], y=df["Profit"], alpha=0.6)
# plt.title("Sales vs Profit")
# plt.xlabel("Sales")
# plt.ylabel("Profit")
# plt.show()

# # 4️⃣ Pie Chart - Sales Distribution by Month
# plt.figure(figsize=(8,8))
# monthly_sales = df.groupby("Month")["Sales"].sum().reindex(months)
# plt.pie(monthly_sales, labels=monthly_sales.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette("pastel"))
# plt.title("Sales Distribution by Month")
# plt.show()

# # 📌 Step 4: Build a Web Dashboard
# category_sales = df.groupby("Category")["Sales"].sum().reset_index()
# scatter_data = df

# app = dash.Dash(__name__)

# app.layout = html.Div([
#     html.H1("Excel Data Visualization Dashboard", style={'text-align': 'center'}),

#     dcc.Graph(
#         id='bar-chart',
#         figure=px.bar(category_sales, x="Category", y="Sales", title="Total Sales by Category", color="Category")
#     ),
    
#     dcc.Graph(
#         id='line-chart',
#         figure=px.line(monthly_sales.reset_index(), x="Month", y="Sales", title="Monthly Sales Trend", markers=True)
#     ),

#     dcc.Graph(
#         id='scatter-plot',
#         figure=px.scatter(scatter_data, x="Sales", y="Profit", title="Sales vs Profit", opacity=0.6)
#     ),
    
#     dcc.Graph(
#         id='pie-chart',
#         figure=px.pie(monthly_sales.reset_index(), names="Month", values="Sales", title="Sales Distribution by Month")
#     ),
# ])

# if __name__ == '__main__':
#     app.run_server(debug=True)





# 📌 Import Required Libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import dash
from dash import dcc, html
import plotly.express as px
import webbrowser

# -----------------------------
# 🎯 Step 1: Generate Sample Data and Save to Excel
# -----------------------------
np.random.seed(42)

# Define categories and months
categories = ['Electronics', 'Clothing', 'Furniture', 'Food', 'Toys']
months = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']

# Create a dataset with random values
data = {
    'Category': np.random.choice(categories, 1000),
    'Month': np.random.choice(months, 1000),
    'Sales': np.random.randint(500, 5000, 1000),
    'Profit': np.random.randint(50, 1000, 1000)
}

# Create DataFrame
df = pd.DataFrame(data)

# Save data to Excel file using openpyxl engine
df.to_excel("data.xlsx", index=False, engine='openpyxl')

print("✅ Data saved to 'data.xlsx' successfully!")

# -----------------------------
# 🎯 Step 2: Read the Excel File
# -----------------------------
df = pd.read_excel("data.xlsx")
print("✅ Data loaded from 'data.xlsx' successfully!")

# -----------------------------
# 🎯 Step 3: Generate Visualizations
# -----------------------------

# 1️⃣ Bar Chart - Total Sales by Category
plt.figure(figsize=(10, 5))
sns.barplot(x=df['Category'], y=df['Sales'], estimator=sum, ci=None, palette="viridis")
plt.title("Total Sales by Category")
plt.xlabel("Category")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.show()

# 2️⃣ Line Graph - Monthly Sales Trend
plt.figure(figsize=(10, 5))
monthly_sales = df.groupby("Month")["Sales"].sum().reindex(months)
sns.lineplot(x=monthly_sales.index, y=monthly_sales.values, marker='o', color='b')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.grid()
plt.show()

# 3️⃣ Scatter Plot - Sales vs Profit
plt.figure(figsize=(8, 5))
sns.scatterplot(x=df["Sales"], y=df["Profit"], alpha=0.6)
plt.title("Sales vs Profit")
plt.xlabel("Sales")
plt.ylabel("Profit")
plt.show()

# 4️⃣ Pie Chart - Sales Distribution by Month
plt.figure(figsize=(8, 8))
monthly_sales = df.groupby("Month")["Sales"].sum().reindex(months)
plt.pie(monthly_sales, labels=monthly_sales.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette("pastel"))
plt.title("Sales Distribution by Month")
plt.show()

print("✅ Visualizations generated successfully!")

# -----------------------------
# 🎯 Step 4: Build a Web Dashboard using Dash
# -----------------------------

# Group data for dashboard visualizations
category_sales = df.groupby("Category")["Sales"].sum().reset_index()
scatter_data = df

# Initialize Dash app
app = dash.Dash(__name__)

# Define Layout for Dash Dashboard
app.layout = html.Div([
    html.H1("📊 Excel Data Visualization Dashboard", style={'text-align': 'center'}),

    # Bar Chart in Dash
    dcc.Graph(
        id='bar-chart',
        figure=px.bar(category_sales, x="Category", y="Sales", title="Total Sales by Category", color="Category")
    ),

    # Line Chart in Dash
    dcc.Graph(
        id='line-chart',
        figure=px.line(monthly_sales.reset_index(), x="Month", y="Sales", title="Monthly Sales Trend", markers=True)
    ),

    # Scatter Plot in Dash
    dcc.Graph(
        id='scatter-plot',
        figure=px.scatter(scatter_data, x="Sales", y="Profit", title="Sales vs Profit", opacity=0.6)
    ),

    # Pie Chart in Dash
    dcc.Graph(
        id='pie-chart',
        figure=px.pie(monthly_sales.reset_index(), names="Month", values="Sales", title="Sales Distribution by Month")
    ),
])

# -----------------------------
# 🎯 Step 5: Run the Dash app
# -----------------------------

if __name__ == '__main__':
    url = "http://127.0.0.1:8050/"
    print(f"🚀 Starting Dash app... Click to open: {url}")
    
    # Automatically open the browser with the URL
    webbrowser.open_new(url)

    # Run the Dash server
    app.run_server(debug=True)
