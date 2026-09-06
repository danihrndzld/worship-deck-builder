#!/usr/bin/env python3
"""Generates reference/example-hymnal.pdf: a small, fully original 2-page
sample hymnal used only to demonstrate the verse/chorus text format this
project's parser expects (numbered verses, a "Coro.-" marker). Every line of
text here was written for this repo — it is not transcribed from any real
hymnal. Point the real tool at your own church's hymnal PDF instead.
"""
from fpdf import FPDF
from fpdf.enums import XPos, YPos

PAGES = [
    {
        "number": 1,
        "title": "CANTO DE EJEMPLO UNO",
        "subtitle": "Ejemplo de Melodia. Do mayor.",
        "verses": [
            [
                "Esta es la primera linea de ejemplo,",
                "esta es la segunda linea de ejemplo,",
                "esta es la tercera linea de ejemplo,",
                "esta es la cuarta linea de ejemplo.",
            ],
            [
                "Aqui va el segundo verso de muestra,",
                "con otra linea que lo acompana,",
                "una tercera linea de relleno,",
                "y una cuarta que lo cierra bien.",
            ],
        ],
        "chorus": [
            "Este es el coro de ejemplo aqui,",
            "con dos lineas cortas asi,",
            "este es el coro de ejemplo aqui,",
            "con dos lineas cortas asi.",
        ],
        "credit": "-Ejemplo.",
    },
    {
        "number": 2,
        "title": "CANTO DE EJEMPLO DOS",
        "subtitle": "Otra Melodia de Muestra. Fa mayor.",
        "verses": [
            [
                "Un segundo canto de muestra empieza,",
                "para probar el formato de este PDF,",
                "cada linea representa una frase,",
                "que el analizador debera reconocer.",
            ],
            [
                "Este es el verso numero dos,",
                "solo para variar el contenido,",
                "y confirmar que el patron se repite,",
                "verso tras verso, coro tras coro.",
            ],
            [
                "Un tercer verso cierra el ejemplo,",
                "con una linea final de muestra,",
                "para que el analizador vea tres versos,",
                "y un coro corto entre ellos.",
            ],
        ],
        "chorus": [
            "Un coro breve de muestra aqui,",
            "solo tres lineas asi,",
            "para probar el caso corto.",
        ],
        "credit": "-Muestra.",
    },
]


def build():
    pdf = FPDF(format="Letter")
    pdf.set_auto_page_break(auto=True, margin=20)
    for page in PAGES:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, str(page["number"]), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 10, page["title"], new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "I", 11)
        pdf.cell(0, 8, page["subtitle"], new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(4)
        pdf.set_font("Helvetica", "", 12)
        for i, verse in enumerate(page["verses"], start=1):
            pdf.multi_cell(0, 7, f"{i} {verse[0]}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            for line in verse[1:]:
                pdf.multi_cell(0, 7, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(3)
            if i == 1:
                pdf.set_font("Helvetica", "I", 12)
                pdf.cell(0, 7, "Coro.-", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                for line in page["chorus"]:
                    pdf.multi_cell(0, 7, line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.set_font("Helvetica", "", 12)
                pdf.ln(3)
        pdf.set_font("Helvetica", "I", 10)
        pdf.cell(0, 7, page["credit"], new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 10)
        pdf.ln(6)
        pdf.cell(0, 7, str(page["number"]), align="C")
    pdf.output("reference/example-hymnal.pdf")
    print("wrote reference/example-hymnal.pdf")


if __name__ == "__main__":
    build()
