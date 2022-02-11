import pandas as pd
import streamlit as st
st.set_page_config(page_title='Supply Chain Transformation',layout="wide")
from multiapp import MultiApp
from apps import home,data
import plotly.express as px
import streamlit.components.v1 as components
from PIL import Image
import base64

#st.set_page_config(page_title='Supply Chain Transformation')

#st.image("image")
st.title('Supply Chain Digital Transformation')
#st.write('Supply Chain Digital Transformation')

app=MultiApp()

app.add_app("Home",home.app)
app.add_app("ETA REQUEST",data.app)

app.run()