import requests
from bs4 import BeautifulSoup
import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from tempfile import NamedTemporaryFile
from openpyxl.styles import Color, PatternFill


def check_amazon(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            if soup.find(string="No disponible por el momento."):
                return "INACTIVO", 0, "-", "-"
            else:
                rating = soup.find("span", "a-icon-alt")
                review = soup.find("span", id="acrCustomerReviewText")
                return ("ACTIVO", "-",
                        (rating.text if rating is not None else "-"),
                        (review.text if review is not None else "-"))
        else:
            return "PAGINA NO ENCONTRADA", 0, "-", "-"
    except requests.RequestException:
        return "Failed to fetch the page", 0, "-", "-"


def check_mercadolibre(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            if "publicación pausada" in response.text.lower():
                return "INACTIVO", 0, "-", "-"
            else:
                price = soup.find("span", "andes-money-amount__fraction")
                rating = soup.find("span", "ui-pdp-review__rating")
                review = soup.find(
                    "p",
                    class_="ui-review-ui-review-capability__rating__label")
                return ("ACTIVO", (price.text if price is not None else "-"),
                        (rating.text if rating is not None else "-"),
                        (review.text if review is not None else "-"))
        else:
            return "PAGINA NO ENCONTRADA", 0, "-", "-"
    except requests.RequestException:
        return "Failed to fetch the page", 0, "-", "-"


# def check_walmart(url):
#     try:
#         response = requests.get(url)
#         if response.status_code == 200:
#             soup = BeautifulSoup(response.text, 'html.parser')
#             # Check specific indicators of availability
#             if "agotado" in response.text.lower():
#                 return "INACTIVO"
#             else:
#                 return "ACTIVO"
#     except requests.RequestException:
#         return "Failed to fetch the page"


def check_liverpool(url):
    try:
        response = requests.get(url)
        if "lo sentimos, la página ha sido actualizada o no existe" in response.text.lower(
        ):
            return "PAGINA NO ENCONTRADA", 0, "-", "-"
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            discount_price = soup.find(
                "p", class_="a-product__paragraphDiscountPrice")
            if discount_price is None:
                discount_price = soup.find(
                    "p", class_="a-product__paragraphRegularPrice")
            rating = soup.find("span", "TTreviewSummaryAverageRating")
            reviews = soup.find("div", "TTreviewCount")
            return ("ACTIVO", (discount_price.text
                               if discount_price is not None else "-"),
                    (rating.text if rating is not None else "-"),
                    (review.text if review is not None else "-"))
        else:
            return "INACTIVO", 0, "-", "-"

    except requests.RequestException:
        return "PAGINA NO ENCONTRADA", 0, "-", "-"


# def check_coppel(url):
#     try:
#         response = requests.get(url)
#         if response.status_code == 404:
#             return "Product link not available"
#         else:
#             return "ACTIVO"
#     except requests.RequestException:
#         return "Failed to fetch the page"


def main():
    # mas de un archivo
    # descarga del zip creado de facturapi
    st.title("Marketplace Product Status Extractor")

    dataset = st.file_uploader("Upload Excel file (.xlsx)", type=['xlsx'])
    results = {}
    if dataset is not None:
        wb = openpyxl.load_workbook(dataset, read_only=True)
        st.info(f"File uploaded: {dataset.name}")

        ws = wb.active

        ###
        # End Result Excel Variables
        result_wb = openpyxl.Workbook()
        result_ws = result_wb.active

        keys = [
            "Marketplace", "Codigo", "Descripcion", "Link", "Estatus",
            "Precio", "Calificación", "# Reseñas"
        ]
        result_row_num = 1

        for col_num, column_title in enumerate(keys, 1):
            cell = result_ws.cell(row=result_row_num, column=col_num)
            cell.value = column_title

        ###
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            result_row_num += 1

            marketplace = row[0]
            product_code = row[1]
            product_name = row[2]
            link = row[3]

            result = ""
            price = "-"
            rating = "-"
            reviews = "-"
            if marketplace == 'Amazon':
                result, price, rating, reviews = check_amazon(link)

            elif marketplace == 'ML':
                result, price, rating, reviews = check_mercadolibre(link)

            elif marketplace == 'Liverpool':
                result, price, rating, reviews = check_liverpool(link)

            row = [
                marketplace, product_code, product_name, link, result, price,
                rating, reviews
            ]

            for col_num, cell_value in enumerate(row, 1):
                cell = result_ws.cell(row=result_row_num, column=col_num)
                cell.value = cell_value
        for e_column in result_ws['E']:
            if e_column.value == "ACTIVO":
                e_column.fill = PatternFill(start_color='38B856',
                                            end_color='38B856',
                                            fill_type='solid')
            if e_column.value == "INACTIVO":
                e_column.fill = PatternFill(start_color='d30000',
                                            end_color='d30000',
                                            fill_type='solid')
            if e_column.value in [
                    "PAGINA NO ENCONTRADA", "Failed to fetch the page"
            ]:
                e_column.fill = PatternFill(start_color='808080',
                                            end_color='808080',
                                            fill_type='solid')
        with NamedTemporaryFile() as tmp:

            result_wb.save(tmp.name)
            data = BytesIO(tmp.read())

        st.subheader("Resultados")
        st.download_button("Descargar Archivo",
                           data=data,
                           mime='xlsx',
                           file_name="resultados.xlsx")


if __name__ == "__main__":
    main()
