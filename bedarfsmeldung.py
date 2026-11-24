
from decimal import Decimal, ROUND_CEILING
from io import BytesIO
from xml.sax.saxutils import escape
import logging

from django.core.mail import EmailMultiAlternatives
from django.core.validators import validate_email
from django.db import transaction

from plugin import InvenTreePlugin
from plugin.mixins import SettingsMixin, ScheduleMixin
from part.models import Part

from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
)

logger = logging.getLogger(__name__)


class BedarfsmeldungPlugin(
    SettingsMixin, ScheduleMixin, InvenTreePlugin
):
    NAME = "Bedarfsmeldung"
    SLUG = "bedarfsmeldung"
    TITLE = "Automatische Bedarfsmeldung"
    DESCRIPTION = "Bedarfsmeldung direkt als E-Mail-Formular"
    VERSION = "1.4.0"

    SCHEDULED_TASKS = {
        "bestandspruefung": {
            "func": "check_stock",
            "schedule": "I",
            "minutes": 1,
        },
    }

    SETTINGS = {
        "ENABLED": {
            "name": "Automatischen Versand aktivieren",
            "default": False,
            "validator": bool,
        },
        "RECIPIENT": {
            "name": "E-Mail-Empfänger",
            "default": "",
        },
        "COMPANY": {
            "name": "Firma",
            "default": "",
        },
        "APPLICANT": {
            "name": "Antragsteller",
            "default": "",
        },
        "TEST_PART_ID": {
            "name": "Testartikel-ID",
            "default": 1,
            "validator": int,
        },
        "ROUNDING": {
            "name": "Bestellmengen-Rundung",
            "default": 10,
            "validator": int,
        },
        "INCLUDE_OPEN_ORDERS": {
            "name": "Offene Bestellungen berücksichtigen",
            "default": True,
            "validator": bool,
        },
        "PREVENT_DUPLICATES": {
            "name": "Doppelte Bedarfsmeldungen verhindern",
            "default": True,
            "validator": bool,
        },
        "ATTACH_PDF": {
            "name": "PDF anhängen",
            "description": "Optionaler PDF-Anhang",
            "default": False,
            "validator": bool,
        },
        "ATTACH_WORD": {
            "name": "Word-Datei anhängen",
            "description": "Optionaler Word-Anhang",
            "default": False,
            "validator": bool,
        },
    }

    def number(self, value):
        if callable(value):
            value = value()
        return Decimal(str(value or 0))

    def fmt(self, value):
        result = format(self.number(value), "f")
        if "." in result:
            result = result.rstrip("0").rstrip(".")
        return result.replace(".", ",")

    def money(self, value):
        if value is None:
            return "Zu prüfen"

        return (
            f"{value:,.2f}"
            .replace(",", "_")
            .replace(".", ",")
            .replace("_", ".")
            + " €"
        )

    def enabled(self, key):
        value = self.get_setting(key)
        return str(value).strip().lower() in (
            "true", "1", "yes", "on"
        )

    def short_reason(self, data):
        return (
            "Mindestbestand unterschritten.\n"
            f"Lagerbestand: {data['stock']} Stück\n"
            f"Maximalbestand: {data['maximum']} Stück"
        )

    def header_rows(self, data):
        return [
            ["Firma:", data["company"]],
            ["Antragsteller:", data["applicant"]],
            ["Kostenstelle oder:", ""],
            ["Innenauftrags Nr oder:", ""],
            ["Kontierungs Nr.:", ""],
            ["Projekt Nr/Name:", ""],
            ["Kunde:", ""],
            ["Einsatzort:", ""],
            ["Termin:", ""],
            [
                "Ausführliche Begründung:",
                self.short_reason(data),
            ],
        ]

    def article_rows(self, data):
        return [
            [
                "Menge",
                "Bezeichnung / Artikel Nr.",
                "Lieferant / Hersteller",
                "Preis netto €/ME",
                "Preis total",
            ],
            [
                str(data["quantity"]),
                data["name"]
                + "\nArtikel Nr.: "
                + data["ipn"],
                data["supplier"],
                data["unit_price"],
                data["total_price"],
            ],
            [
                "Gesamtsumme:",
                "",
                "",
                "",
                data["total_price"],
            ],
        ]

    # ==========================================
    # HTML-FORMULAR DIREKT IN DER E-MAIL
    # ==========================================

    def html_value(self, value):
        if not value:
            return "&nbsp;"

        return escape(str(value)).replace(
            "\n", "<br>"
        )

    def create_html(self, data):
        table_style = (
            "border-collapse:collapse;"
            "table-layout:fixed;"
            "width:100%;"
            "font-family:Arial,Helvetica,sans-serif;"
            "font-size:16px;"
            "line-height:1.3;"
            "color:#000000;"
            "background-color:#ffffff;"
            "margin:0;"
            "border-spacing:0;"
        )

        cell_style = (
            "border:1px solid #000000;"
            "padding:7px;"
            "vertical-align:middle;"
            "overflow-wrap:break-word;"
            "color:#000000;"
        )

        gray_style = (
            cell_style
            + "background-color:#BFBFBF;"
            + "text-align:center;"
        )

        white_style = (
            cell_style
            + "background-color:#FFFFFF;"
        )

        html = (
            '<div style="width:100%;'
            'max-width:1000px;margin:0;padding:0;">'
        )

        html += (
            '<table cellpadding="0" cellspacing="0"'
            f' style="{table_style}">'
        )

        for label, value in self.header_rows(data):
            html += (
                "<tr>"
                f'<td width="29%" style="{gray_style}">'
                + self.html_value(label)
                + "</td>"
                f'<td width="71%" style="{white_style}">'
                + self.html_value(value)
                + "</td>"
                "</tr>"
            )

        html += "</table>"

        html += (
            '<table cellpadding="0" cellspacing="0"'
            f' style="{table_style}">'
        )

        rows = self.article_rows(data)

        widths = [
            "9%",
            "37%",
            "20%",
            "17%",
            "17%",
        ]

        html += "<tr>"

        for index, label in enumerate(rows[0]):
            html += (
                f'<td width="{widths[index]}"'
                f' style="{gray_style}">'
                + self.html_value(label)
                + "</td>"
            )

        html += "</tr><tr>"

        for index, value in enumerate(rows[1]):
            alignment = (
                "left" if index == 1 else "center"
            )

            html += (
                f'<td style="{white_style}'
                f'text-align:{alignment};">'
                + self.html_value(value)
                + "</td>"
            )

        html += "</tr>"

        html += (
            "<tr>"
            f'<td colspan="4" style="{white_style}'
            'text-align:center;">'
            "Gesamtsumme:"
            "</td>"
            f'<td style="{white_style}'
            'text-align:center;">'
            + self.html_value(data["total_price"])
            + "</td>"
            "</tr>"
        )

        html += "</table></div>"

        return html

    # ==========================================
    # OPTIONALER WORD-ANHANG
    # ==========================================

    def shade(self, cell):
        properties = cell._tc.get_or_add_tcPr()

        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "BFBFBF")

        properties.append(shading)

    def format_cell(self, cell, centered=False):
        cell.vertical_alignment = (
            WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )

        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0

            if centered:
                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                )

    def word_table(self, document, rows, widths):
        table = document.add_table(
            rows=len(rows),
            cols=len(widths),
        )

        table.style = "Table Grid"
        table.autofit = False

        for i, width in enumerate(widths):
            table.columns[i].width = Cm(width)

        for r, values in enumerate(rows):
            for c, value in enumerate(values):
                cell = table.cell(r, c)
                cell.width = Cm(widths[c])
                cell.text = str(value)
                self.format_cell(cell)

        return table

    def create_word(self, data):
        document = Document()

        section = document.sections[0]

        section.page_width = Cm(21)
        section.page_height = Cm(29.7)

        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1)
        section.right_margin = Cm(1)

        normal = document.styles["Normal"]

        normal.font.name = "Arial"
        normal.font.size = Pt(12)

        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.line_spacing = 1.0

        header = self.word_table(
            document,
            self.header_rows(data),
            [5.5, 13.5],
        )

        for row in header.rows:
            self.shade(row.cells[0])
            self.format_cell(
                row.cells[0],
                centered=True,
            )

        articles = self.word_table(
            document,
            self.article_rows(data),
            [1.7, 7, 3.9, 3.2, 3.2],
        )

        for cell in articles.rows[0].cells:
            self.shade(cell)
            self.format_cell(cell, centered=True)

        for index in (0, 2, 3, 4):
            self.format_cell(
                articles.cell(1, index),
                centered=True,
            )

        last = articles.rows[-1]

        merged = last.cells[0].merge(last.cells[3])
        merged.text = "Gesamtsumme:"

        self.format_cell(merged, centered=True)

        self.format_cell(
            last.cells[4],
            centered=True,
        )

        output = BytesIO()
        document.save(output)

        return output.getvalue()

    # ==========================================
    # OPTIONALER PDF-ANHANG
    # ==========================================

    def pdf_table(self, rows, widths, centered=None):
        centered = centered or set()

        normal = ParagraphStyle(
            "NormalCell",
            fontName="Helvetica",
            fontSize=12,
            leading=14,
            spaceBefore=0,
            spaceAfter=0,
        )

        center = ParagraphStyle(
            "CenteredCell",
            parent=normal,
            alignment=1,
        )

        converted = []

        for r, row in enumerate(rows):
            converted_row = []

            for c, value in enumerate(row):
                style = (
                    center
                    if (r, c) in centered
                    else normal
                )

                converted_row.append(
                    Paragraph(
                        escape(str(value)).replace(
                            "\n", "<br/>"
                        ),
                        style,
                    )
                )

            converted.append(converted_row)

        table = Table(
            converted,
            colWidths=widths,
            spaceBefore=0,
            spaceAfter=0,
        )

        table.setStyle(TableStyle([
            (
                "GRID",
                (0, 0), (-1, -1),
                0.6, colors.black,
            ),
            (
                "VALIGN",
                (0, 0), (-1, -1),
                "MIDDLE",
            ),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))

        return table

    def create_pdf(self, data):
        output = BytesIO()

        document = SimpleDocTemplate(
            output,
            pagesize=A4,
            leftMargin=28.35,
            rightMargin=28.35,
            topMargin=42.52,
            bottomMargin=42.52,
        )

        gray = colors.HexColor("#BFBFBF")

        header_rows = self.header_rows(data)

        header = self.pdf_table(
            header_rows,
            [153, 374],
            centered={
                (r, 0)
                for r in range(len(header_rows))
            },
        )

        header.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), gray),
        ]))

        centered = {
            (0, c) for c in range(5)
        }

        centered.update({
            (1, 0),
            (1, 2),
            (1, 3),
            (1, 4),
            (2, 0),
            (2, 4),
        })

        articles = self.pdf_table(
            self.article_rows(data),
            [47, 194, 108, 89, 89],
            centered=centered,
        )

        articles.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), gray),
            ("SPAN", (0, -1), (3, -1)),
        ]))

        document.build([
            header,
            articles,
        ])

        return output.getvalue()

    # ==========================================
    # BESTANDSPRÜFUNG UND E-MAIL-VERSAND
    # ==========================================

    def check_stock(self):
        if not self.enabled("ENABLED"):
            return

        recipient = str(
            self.get_setting("RECIPIENT") or ""
        ).strip()

        if not recipient:
            return

        validate_email(recipient)

        part_id = int(
            self.get_setting("TEST_PART_ID") or 1
        )

        rounding = max(
            1,
            int(self.get_setting("ROUNDING") or 10)
        )

        with transaction.atomic():

            part = (
                Part.objects.select_for_update()
                .filter(pk=part_id, active=True)
                .first()
            )

            if part is None:
                return

            stock = self.number(part.total_stock)
            minimum = self.number(part.minimum_stock)
            maximum = self.number(part.maximum_stock)
            ordered = self.number(part.on_order)

            metadata = dict(part.metadata or {})
            key = "bedarfsmeldung_test_sent"

            if stock >= minimum:
                if metadata.pop(key, None):
                    part.metadata = metadata
                    part.save(update_fields=["metadata"])

                return

            if maximum <= 0 or maximum < minimum:
                return

            include_orders = self.enabled(
                "INCLUDE_OPEN_ORDERS"
            )

            considered = (
                ordered
                if include_orders
                else Decimal("0")
            )

            demand = maximum - stock - considered

            if demand <= 0:
                return

            if (
                self.enabled("PREVENT_DUPLICATES")
                and metadata.get(key)
            ):
                return

            quantity = int(
                (demand / Decimal(rounding))
                .to_integral_value(
                    rounding=ROUND_CEILING
                )
                * rounding
            )

            supplier_part = part.default_supplier
            supplier = "Nicht hinterlegt"

            unit_price = None
            total_price = None

            if supplier_part:
                supplier = str(
                    supplier_part.supplier.name
                )

                try:
                    price = supplier_part.get_price(
                        quantity,
                        multiples=False,
                        currency="EUR",
                    )

                    if price is not None:
                        total_price = self.number(price)

                        unit_price = (
                            total_price
                            / Decimal(quantity)
                        )

                except Exception:
                    logger.exception(
                        "Bedarfsmeldung: Preisabfrage "
                        "für Artikel %s fehlgeschlagen",
                        part.pk,
                    )

            data = {
                "company": str(
                    self.get_setting("COMPANY") or ""
                ),
                "applicant": str(
                    self.get_setting("APPLICANT") or ""
                ),
                "name": str(part.name),
                "ipn": str(part.IPN or ""),
                "supplier": supplier,
                "stock": self.fmt(stock),
                "minimum": self.fmt(minimum),
                "maximum": self.fmt(maximum),
                "ordered": self.fmt(ordered),
                "considered": self.fmt(considered),
                "demand": self.fmt(demand),
                "quantity": quantity,
                "unit_price": self.money(unit_price),
                "total_price": self.money(total_price),
            }

            logger.info(
                "Bedarfsmeldung 1.4.0: "
                "Artikel=%s, Bestand=%s, "
                "Minimum=%s, Maximum=%s, "
                "Offene Bestellungen=%s, "
                "Beruecksichtigen=%s, "
                "Rundung=%s, Bedarf=%s, Menge=%s, "
                "Gesamtpreis=%s",
                part.pk,
                stock,
                minimum,
                maximum,
                ordered,
                include_orders,
                rounding,
                demand,
                quantity,
                total_price,
            )

            subject = (
                "[InvenTree] Bedarfsmeldung - "
                + (data["ipn"] or data["name"])
            )

            # Textversion für E-Mail-Programme ohne HTML
            message = "\n".join(
                f"{label} {value}"
                for label, value in self.header_rows(data)
            )

            message += (
                f"\nMenge: {quantity}"
                f"\nBezeichnung: {data['name']}"
                f"\nArtikel Nr.: {data['ipn']}"
                f"\nLieferant: {supplier}"
                f"\nPreis netto €/ME: {data['unit_price']}"
                f"\nGesamtsumme: {data['total_price']}"
            )

            email = EmailMultiAlternatives(
                subject=subject,
                body=message,
                to=[recipient],
            )

            # Firmenformular direkt in der E-Mail
            email.attach_alternative(
                self.create_html(data),
                "text/html",
            )

            safe_ipn = "".join(
                c if c.isalnum() or c in "-_" else "_"
                for c in (
                    data["ipn"] or str(part.pk)
                )
            )

            filename = "Bedarfsmeldung_" + safe_ipn

            # Anhänge nur bei aktivierter Einstellung
            if self.enabled("ATTACH_WORD"):
                email.attach(
                    filename + ".docx",
                    self.create_word(data),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            if self.enabled("ATTACH_PDF"):
                email.attach(
                    filename + ".pdf",
                    self.create_pdf(data),
                    "application/pdf",
                )

            sent = email.send(
                fail_silently=False
            )

            if sent == 1:
                metadata[key] = True
                part.metadata = metadata
                part.save(
                    update_fields=["metadata"]
                )
