from plugin import InvenTreePlugin
from plugin.mixins import SettingsMixin, EventMixin


class BedarfsmeldungPlugin(
    SettingsMixin, EventMixin, InvenTreePlugin
):

    NAME = "Bedarfsmeldung"
    SLUG = "bedarfsmeldung"
    TITLE = "Automatische Bedarfsmeldung"
    DESCRIPTION = "Automatische Bedarfsmeldungen bei Mindestbestandsunterschreitung"
    VERSION = "1.0.0"

    SETTINGS = {
        "ENABLED": {
            "name": "Automatischen Versand aktivieren",
            "description": "Bedarfsmeldungen automatisch versenden",
            "default": False,
            "validator": bool,
        },
        "RECIPIENT": {
            "name": "E-Mail-Empfänger",
            "description": "Empfänger der Bedarfsmeldung",
            "default": "",
        },
        "ROUNDING": {
            "name": "Bestellmengen-Rundung",
            "description": "Auf dieses Vielfache aufrunden",
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

    def wants_process_event(self, event):
        return False

    def process_event(self, event, *args, **kwargs):
        pass
