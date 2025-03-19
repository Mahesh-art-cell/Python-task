import os
import pandas as pd
import numpy as np
import dash
from dash import dcc,html
import plotly.express as px
import webbrowser


# step:1
# ----------------

np.random.seed(42)

# define
categories = ["electronics","clothing","furniture","food","toys"]

months = ["janu","feb"]


# dataset
data = {
    "categories" : np.random.choice(categories,1000),
    "month" : np.random.choice(months,1000),
    "sales" : np.random.randit(500,5000,1000),
    "profit" :np.random.randit(500,5000,1000)

}


# create dataframe

df = pd.DataFrame(data)


df.to_excel("data.xlsx",index=False,engine="openpyxl")


print("success")


# step-2
# --------------------


df = pd.read_excel("data.xlsx")



# step-3
# ---------------------


category_sales = df.groupby("category")["sales"].sum().reset_index()