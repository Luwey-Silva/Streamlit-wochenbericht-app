import streamlit as st
import pandas as pd

st.title("Wochenbericht APP LDS")
file = "Wochenberichte.csv"
df = pd.read_csv(file)
st.header("Existing File")
st.write(df)

st.sidebar.header("Options")
options_form = st.sidebar.form("options_form")
user_name = options_form.text_input("Name")
user_age = options_form.text_input("Age")
add_data = options_form.form_submit_button()
if add_data:
    new_data = {"name": user_name, "age": int(user_age)}
    df = df. append(new_data, ignore_index=True)
    df.to_csv(file, index=False)
