import pdfplumber
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt

with pdfplumber.open("table.pdf") as pdf:
    first_page = pdf.pages[0]
    print(first_page.chars[0])