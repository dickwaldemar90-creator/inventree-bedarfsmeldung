from decimal import Decimal, ROUND_CEILING
from django.core.mail import send_mail
from django.core.validators import validate_email
from django.db import transaction

from plugin import InvenTreePlugin
from plugin.mixins import SettingsMixin, ScheduleMixin
from part.models import Part


class BedarfsmeldungPlugin(
    SettingsMixin, ScheduleMixin, InvenTreePlugin
):

    NAME = "Bedarfsmeldung"
    SLUG = "bedarfsmeldung"
    TITLE = "Automatische Bedarfsmeldung"
    DESCRIPTION = "Automatische Bedarfsmeldungen bei Mindestbestandsunterschreitung"
    VERSION = "1.1.0"

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
        "TEST_PART_ID": {
            "name": "Testartikel-ID",
            "description": "Nur diesen Artikel überwachen",
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
            "description": "In dieser Testversion noch nicht verfügbar",
            "default": True,
            "validator": bool,
        },
        "ATTACH_WORD": {
            "name": "Word-Datei anhängen",
            "description": "In dieser Testversion noch nicht verfügbar",
            "default": True,
            "validator": bool,
        },
    }

    def number(self, value):
        if callable(value):
            value = value()

        return Decimal(str(value or 0))

    def check_stock(self):

        if not self.get_setting("ENABLED"):
            return

        recipient = str(
            self.get_setting("RECIPIENT") or ""
        ).strip()

        if not recipient:
            return

        validate_email(recipient)

        part_id = int(
            self.get_setting("TEST_PART_ID")
        )

        rounding = max(
            1, int(self.get_setting("ROUNDING"))
        )

        with transaction.atomic():

            part = Part.objects.select_for_update().filter(
                pk=part_id,
                active=True
            ).first()

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

            considered_orders = (
                ordered
                if self.get_setting("INCLUDE_OPEN_ORDERS")
                else Decimal("0")
            )

            demand = maximum - stock - considered_orders

            if demand <= 0:
                return

            if (
                self.get_setting("PREVENT_DUPLICATES")
                and metadata.get(key)
            ):
                return

            quantity = int(
                (demand / Decimal(rounding))
                .to_integral_value(rounding=ROUND_CEILING)
                * rounding
            )

            supplier_part = part.default_supplier

            supplier = (
                str(supplier_part.supplier.name)
                if supplier_part
                else "Nicht hinterlegt"
            )

            subject = (
                f"[InvenTree] Bedarfsmeldung - "
                f"{part.IPN or part.name}"
            )

            message = (
                "AUTOMATISCHE BEDARFSMELDUNG\n"
                "Testversion 1.1.0\n\n"
                f"Artikel: {part.name}\n"
                f"Artikelnummer: {part.IPN}\n"
                f"Lieferant: {supplier}\n\n"
                f"Lagerbestand: {stock:f}\n"
                f"Mindestbestand: {minimum:f}\n"
                f"Maximalbestand: {maximum:f}\n"
                f"Bereits bestellt: {ordered:f}\n"
                f"Berücksichtigte Bestellungen: "
                f"{considered_orders:f}\n"
                f"Rechnerischer Bedarf: {demand:f}\n\n"
                f"Empfohlene Bestellmenge: "
                f"{quantity} Stück\n\n"
                "Preis: Noch nicht berücksichtigt\n"
                "Kostenstelle: Manuell ergänzen\n\n"
                "Bitte Bedarf prüfen.\n"
                "Es wurde keine Bestellung ausgelöst."
            )

            sent = send_mail(
                subject,
                message,
                None,
                [recipient],
                fail_silently=False,
            )

            if sent == 1:
                metadata[key] = True
                part.metadata = metadata
                part.save(update_fields=["metadata"])
