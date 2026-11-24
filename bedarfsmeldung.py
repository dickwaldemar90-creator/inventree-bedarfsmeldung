
from decimal import Decimal, ROUND_CEILING
from io import BytesIO
from xml.sax.saxutils import escape

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

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)


class BedarfsmeldungPlugin(
    SettingsMixin, ScheduleMixin, InvenTreePlugin
):
    NAME = "Bedarfsmeldung"
    SLUG = "bedarfsmeldung"
    TITLE = "Automatische Bedarfsmeldung"
    DESCRIPTION = "Bedarfsmeldungen mit Word und PDF"
    VERSION = "1.2.0"

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
            "default": True,
            "validator": bool,
        },
        "ATTACH_WORD": {
            "name": "Word-Datei anhängen",
            "default": True,
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
        ]

    def article_rows(self, data):
        rows = [
            [
                "Menge",
                "Bezeichnung/ArtNr.",
                "Lieferant/Hersteller",
                "Preis netto €/ME",
                "Preis total",
            ],
            [
                str(data["quantity"]),
                data["name"] + "\n" + data["ipn"],
                data["supplier"],
                data["unit_price"],
                data["total_price"],
            ],
        ]

        for _ in range(4):
            rows.append(["", "", "", "", ""])

        rows.append([
            "", "Gesamtsumme:", "", "",
            data["total_price"]
        ])

        return rows

    def reason(self, data):
        return (
            "Mindestbestand unterschritten.\n"
            f"Lagerbestand: {data['stock']} Stück\n"
            f"Mindestbestand: {data['minimum']} Stück\n"
            f"Maximalbestand: {data['maximum']} Stück\n"
            f"Offene Bestellungen: {data['ordered']} Stück\n"
            f"Berücksichtigt: {data['considered']} Stück\n"
            f"Rechnerischer Bedarf: {data['demand']} Stück\n"
            f"Bestellmenge: {data['quantity']} Stück"
        )

    def shade(self, cell):
        properties = cell._tc.get_or_add_tcPr()
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "D9D9D9")
        properties.append(shading)

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

                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(3)
                    paragraph.paragraph_format.space_before = Pt(3)

        return table

    def create_word(self, data):
        document = Document()

        section = document.sections[0]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

        normal = document.styles["Normal"]
        normal.font.name = "Arial"
        normal.font.size = Pt(10)

        header = self.word_table(
            document,
            self.header_rows(data),
            [5, 13],
        )

        for row in header.rows:
            self.shade(row.cells[0])

        document.add_paragraph()

        reason = self.word_table(
            document,
            [
                ["Ausführliche Begründung:"],
                [self.reason(data)],
            ],
            [18],
        )

        self.shade(reason.cell(0, 0))

        document.add_paragraph()

        articles = self.word_table(
            document,
            self.article_rows(data),
            [1.5, 7, 4, 2.75, 2.75],
        )

        for cell in articles.rows[0].cells:
            self.shade(cell)

        last = articles.rows[-1]
        last.cells[0].merge(last.cells[3])
        last.cells[0].text = "Gesamtsumme:"
        self.shade(last.cells[0])

        output = BytesIO()
        document.save(output)
        return output.getvalue()

    def pdf_table(self, rows, widths):
        styles = getSampleStyleSheet()
        style = styles["Normal"]
        style.fontName = "Helvetica"
        style.fontSize = 9
        style.leading = 12

        converted = []

        for row in rows:
            converted.append([
                Paragraph(
                    escape(str(value)).replace("\n", "<br/>"),
                    style,
                )
                for value in row
            ])

        table = Table(converted, colWidths=widths)

        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))

        return table

    def create_pdf(self, data):
        output = BytesIO()

        document = SimpleDocTemplate(
            output,
            pagesize=A4,
            leftMargin=42,
            rightMargin=42,
            topMargin=56,
            bottomMargin=56,
        )

        gray = colors.HexColor("#D9D9D9")

        header = self.pdf_table(
            self.header_rows(data),
            [142, 368],
        )

        header.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), gray),
        ]))

        reason = self.pdf_table(
            [
                ["Ausführliche Begründung:"],
                [self.reason(data)],
            ],
            [510],
        )

        reason.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), gray),
        ]))

        articles = self.pdf_table(
            self.article_rows(data),
            [43, 197, 114, 78, 78],
        )

        articles.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), gray),
            ("SPAN", (0, -1), (3, -1)),
            ("BACKGROUND", (0, -1), (3, -1), gray),
        ]))

        document.build([
            header,
            Spacer(1, 18),
            reason,
            Spacer(1, 18),
            articles,
        ])

        return output.getvalue()

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

            considered = (
                ordered
                if self.enabled("INCLUDE_OPEN_ORDERS")
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
                .to_integral_value(rounding=ROUND_CEILING)
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

                price = supplier_part.get_price(
                    quantity,
                    multiples=False,
                    currency="EUR",
                )

                if price is not None:
                    total_price = self.number(price)
                    unit_price = (
                        total_price / Decimal(quantity)
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

            subject = (
                "[InvenTree] Bedarfsmeldung - "
                + (data["ipn"] or data["name"])
            )

            message = (
                "AUTOMATISCHE BEDARFSMELDUNG\n\n"
                f"Artikel: {data['name']}\n"
                f"Artikelnummer: {data['ipn']}\n"
                f"Lieferant: {supplier}\n\n"
                + self.reason(data)
                + "\n\n"
                f"Einzelpreis: {data['unit_price']}\n"
                f"Gesamtpreis: {data['total_price']}\n\n"
                "Bitte Bedarfsmeldung prüfen.\n"
                "Kostenstelle manuell ergänzen.\n"
                "Es wurde keine Bestellung ausgelöst."
            )

            email = EmailMultiAlternatives(
                subject=subject,
                body=message,
                to=[recipient],
            )

            html_rows = [
                ("Artikelnummer", data["ipn"]),
                ("Bezeichnung", data["name"]),
                ("Lieferant", supplier),
                ("Lagerbestand", data["stock"]),
                ("Mindestbestand", data["minimum"]),
                ("Maximalbestand", data["maximum"]),
                ("Offene Bestellungen", data["ordered"]),
                ("Berücksichtigt", data["considered"]),
                ("Rechnerischer Bedarf", data["demand"]),
                ("Bestellmenge", str(quantity)),
                ("Einzelpreis", data["unit_price"]),
                ("Gesamtpreis", data["total_price"]),
            ]

            html = (
                '<h2>Automatische Bedarfsmeldung</h2>'
                '<p>Der Mindestbestand wurde unterschritten.</p>'
                '<table style="border-collapse:collapse;'
                'font-family:Arial;font-size:14px;">'
            )

            for label, value in html_rows:
                html += (
                    '<tr><td style="background:#eeeeee;'
                    'border:1px solid #888;padding:8px;">'
                    + escape(label)
                    + '</td><td style="border:1px solid #888;'
                    'padding:8px;">'
                    + escape(str(value))
                    + "</td></tr>"
                )

            html += (
                "</table>"
                "<p>Bitte Bedarfsmeldung prüfen.</p>"
                "<p>Es wurde keine Bestellung ausgelöst.</p>"
            )

            email.attach_alternative(html, "text/html")

            safe_ipn = "".join(
                c if c.isalnum() or c in "-_" else "_"
                for c in (data["ipn"] or str(part.pk))
            )

            filename = "Bedarfsmeldung_" + safe_ipn

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

            sent = email.send(fail_silently=False)

            if sent == 1:
                metadata[key] = True
                part.metadata = metadata
                part.save(update_fields=["metadata"])
