
from decimal import Decimal, ROUND_CEILING
from xml.sax.saxutils import escape
import logging

from django.core.mail import EmailMultiAlternatives
from django.core.validators import validate_email
from django.db import transaction

from plugin import InvenTreePlugin
from plugin.mixins import SettingsMixin, ScheduleMixin
from part.models import Part


logger = logging.getLogger(__name__)


class BedarfsmeldungPlugin(
    SettingsMixin, ScheduleMixin, InvenTreePlugin
):
    NAME = "Bedarfsmeldung"
    SLUG = "bedarfsmeldung"
    TITLE = "Automatische Bedarfsmeldung"
    DESCRIPTION = "Automatische Ersatzteil-Bedarfsüberwachung"
    VERSION = "2.0.0"

    STATE_KEY = "bedarfsmeldung_v2"

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
        "TEST_MODE": {
            "name": "Testmodus aktivieren",
            "description": (
                "An: Nur Testartikel überwachen. "
                "Aus: Alle aktiven Artikel überwachen."
            ),
            "default": True,
            "validator": bool,
        },
        "TEST_PART_ID": {
            "name": "Testartikel-ID",
            "description": (
                "Wird nur im Testmodus verwendet."
            ),
            "default": 1,
            "validator": int,
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
        "ROUNDING": {
            "name": "Bestellmengen-Rundung",
            "default": 10,
            "validator": int,
        },
        "INCLUDE_OPEN_ORDERS": {
            "name": "Offene Bestellungen berücksichtigen",
            "default": False,
            "validator": bool,
        },
        "PREVENT_DUPLICATES": {
            "name": "Bereits gemeldeten Bedarf berücksichtigen",
            "description": (
                "Verhindert wiederholte Meldungen "
                "für bereits gemeldete Mengen."
            ),
            "default": True,
            "validator": bool,
        },
        "RESET_PART_ID": {
            "name": "Bedarf zurücksetzen – Artikel-ID",
            "description": (
                "Interne ID des Artikels, dessen "
                "gemeldeter Bedarf zurückgesetzt werden soll."
            ),
            "default": 0,
            "validator": int,
        },
        "RESET_TOKEN": {
            "name": "Rücksetznummer",
            "description": (
                "Für jede Rücksetzung die Nummer erhöhen."
            ),
            "default": 0,
            "validator": int,
        },
    }

    # ==========================================
    # HILFSFUNKTIONEN
    # ==========================================

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
        return str(
            self.get_setting(key)
        ).strip().lower() in (
            "true", "1", "yes", "on"
        )

    def html_value(self, value):
        if value is None or value == "":
            return "&nbsp;"

        return escape(str(value)).replace(
            "\n", "<br>"
        )

    # ==========================================
    # FORMULARDATEN
    # ==========================================

    def header_rows(self, data):
        reason = (
            "Mindestbestand unterschritten.\n"
            f"Lagerbestand: {data['stock']} Stück\n"
            f"Maximalbestand: {data['maximum']} Stück"
        )

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
            ["Ausführliche Begründung:", reason],
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
        ]

    # ==========================================
    # HTML-FORMULAR DIREKT IN DER E-MAIL
    # ==========================================

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
                + "</td></tr>"
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
    # LIEFERANT UND STAFFELPREISE
    # ==========================================

    def get_supplier_price(self, part, quantity):
        supplier = "Nicht hinterlegt"
        unit_price = None
        total_price = None

        supplier_part = part.default_supplier

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
                    "Preisabfrage fehlgeschlagen: "
                    "Artikel %s",
                    part.pk,
                )

        return supplier, unit_price, total_price

    # ==========================================
    # E-MAIL VERSENDEN
    # ==========================================

    def send_notification(
        self,
        part,
        stock,
        minimum,
        maximum,
        quantity,
        recipient,
    ):
        supplier, unit_price, total_price = (
            self.get_supplier_price(
                part,
                quantity,
            )
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
            "quantity": quantity,
            "unit_price": self.money(unit_price),
            "total_price": self.money(total_price),
        }

        subject = (
            "[InvenTree] Bedarfsmeldung - "
            + (data["ipn"] or data["name"])
        )

        # Textversion für Programme ohne HTML

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

        # Formular direkt in der E-Mail.
        # Keine Word- oder PDF-Anhänge.

        email.attach_alternative(
            self.create_html(data),
            "text/html",
        )

        return email.send(
            fail_silently=False
        ) == 1

    # ==========================================
    # MANUELLE BEDARFSRÜCKSETZUNG
    # ==========================================

    def apply_reset(self):
        part_id = int(
            self.get_setting("RESET_PART_ID") or 0
        )

        token = int(
            self.get_setting("RESET_TOKEN") or 0
        )

        if part_id <= 0 or token <= 0:
            return

        with transaction.atomic():

            part = (
                Part.objects.select_for_update()
                .filter(pk=part_id)
                .first()
            )

            if part is None:
                return

            metadata = dict(part.metadata or {})

            state = dict(
                metadata.get(self.STATE_KEY) or {}
            )

            if token <= int(
                state.get("reset_token") or 0
            ):
                return

            state["pending"] = "0"

            state["last_stock"] = str(
                self.number(part.total_stock)
            )

            state["last_ordered"] = str(
                self.number(part.on_order)
            )

            state["reset_token"] = token

            metadata[self.STATE_KEY] = state

            metadata.pop(
                "bedarfsmeldung_test_sent",
                None,
            )

            part.metadata = metadata

            part.save(
                update_fields=["metadata"]
            )

            logger.info(
                "Bedarf zurückgesetzt: Artikel %s",
                part_id,
            )

    # ==========================================
    # EINZELNEN ARTIKEL ÜBERPRÜFEN
    # ==========================================

    def check_part(
        self,
        part_id,
        recipient,
        rounding,
    ):
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

            ordered = max(
                Decimal("0"),
                self.number(part.on_order),
            )

            metadata = dict(part.metadata or {})

            state = dict(
                metadata.get(self.STATE_KEY) or {}
            )

            previous_stock = self.number(
                state.get("last_stock", stock)
            )

            previous_ordered = self.number(
                state.get("last_ordered", ordered)
            )

            pending = max(
                Decimal("0"),
                self.number(
                    state.get("pending", 0)
                ),
            )

            # Jede beobachtete Bestandserhöhung
            # reduziert den gemeldeten Bedarf.

            increase = max(
                Decimal("0"),
                stock - previous_stock,
            )

            pending = max(
                Decimal("0"),
                pending - increase,
            )

            include_orders = self.enabled(
                "INCLUDE_OPEN_ORDERS"
            )

            # Optional offene Bestellungen
            # aus InvenTree berücksichtigen.

            if include_orders:
                new_orders = max(
                    Decimal("0"),
                    ordered - previous_ordered,
                )

                pending = max(
                    Decimal("0"),
                    pending - new_orders,
                )

            considered_orders = (
                ordered
                if include_orders
                else Decimal("0")
            )

            # Alte Testmeldungen übernehmen.
            # Die damalige Menge ist unbekannt.

            if (
                not state
                and metadata.get(
                    "bedarfsmeldung_test_sent"
                )
                and stock < minimum
                and maximum >= minimum
            ):
                pending = max(
                    pending,
                    maximum
                    - stock
                    - considered_orders,
                    Decimal("0"),
                )

            metadata.pop(
                "bedarfsmeldung_test_sent",
                None,
            )

            # Nur gültige Mindest- und
            # Maximalbestände berücksichtigen.

            valid_limits = (
                minimum > 0
                and maximum >= minimum
            )

            if valid_limits and stock < minimum:

                considered_pending = (
                    pending
                    if self.enabled(
                        "PREVENT_DUPLICATES"
                    )
                    else Decimal("0")
                )

                demand = (
                    maximum
                    - stock
                    - considered_orders
                    - considered_pending
                )

                if demand > 0:

                    # Bedarf auf die eingestellte
                    # Bestellmenge aufrunden.

                    quantity = int(
                        (
                            demand
                            / Decimal(rounding)
                        ).to_integral_value(
                            rounding=ROUND_CEILING
                        )
                        * rounding
                    )

                    logger.info(
                        "Bedarf berechnet: "
                        "Artikel=%s Bestand=%s "
                        "Minimum=%s Maximum=%s "
                        "Gemeldet=%s Bestellt=%s "
                        "Bedarf=%s Menge=%s",
                        part.pk,
                        stock,
                        minimum,
                        maximum,
                        pending,
                        considered_orders,
                        demand,
                        quantity,
                    )

                    sent = self.send_notification(
                        part,
                        stock,
                        minimum,
                        maximum,
                        quantity,
                        recipient,
                    )

                    # Nur erfolgreich versendete
                    # Meldungen speichern.

                    if sent:
                        pending += Decimal(quantity)

                        logger.info(
                            "Bedarfsmeldung versendet: "
                            "Artikel=%s Menge=%s "
                            "Offener Bedarf=%s",
                            part.pk,
                            quantity,
                            pending,
                        )

            # Bedarf und letzten Bestand
            # dauerhaft speichern.

            state["pending"] = str(pending)
            state["last_stock"] = str(stock)
            state["last_ordered"] = str(ordered)

            metadata[self.STATE_KEY] = state

            if metadata != part.metadata:
                part.metadata = metadata

                part.save(
                    update_fields=["metadata"]
                )

    # ==========================================
    # AUTOMATISCHE BESTANDSÜBERWACHUNG
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

        rounding = max(
            1,
            int(
                self.get_setting("ROUNDING") or 10
            ),
        )

        self.apply_reset()

        # Alle aktiven Artikel auswählen.

        parts = Part.objects.filter(
            active=True
        )

        # Nur im Testmodus auf eine ID begrenzen.

        if self.enabled("TEST_MODE"):
            part_id = int(
                self.get_setting("TEST_PART_ID") or 1
            )

            parts = parts.filter(pk=part_id)

        part_ids = list(
            parts.values_list("pk", flat=True)
        )

        # Jeder Artikel wird separat geprüft.
        # Ein Fehler stoppt nicht alle Artikel.

        for part_id in part_ids:

            try:
                self.check_part(
                    part_id,
                    recipient,
                    rounding,
                )

            except Exception:
                logger.exception(
                    "Bedarfsmeldung: Fehler bei "
                    "Artikel %s",
                    part_id,
                )
