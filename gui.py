import base64

import streamlit as st
from io import StringIO
from src.libraries.presenze import Presenze
from datetime import datetime, timedelta
import pandas as pd
from datetime import date


uploaded_file = st.file_uploader("Choose a file")
if uploaded_file is not None:

    form = st.form("form")

    base = date.today().replace(day=1)
    m = pd.date_range(
        start=base - timedelta(days=365), end=base, freq='MS').sort_values(ascending=False).strftime('%m/%Y').tolist()
    year_month = form.selectbox('Mese', m, index=1)

    submitted = form.form_submit_button("Elabora")

    if submitted:
        presenze = Presenze()
        month, year = year_month.split('/')
        dipendenti = presenze.read_file(uploaded_file, uploaded_file.name, int(month), int(year))
        # st.write(dipendenti)


        filename = presenze.create_xml(dipendenti, int(month))
        # st.write(presenze)
        # zip = presenze.create_zip(month_number)
        with open(filename, encoding="utf8") as f:
            st.download_button("Download file", f, file_name=filename)
