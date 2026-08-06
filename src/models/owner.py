import json

class Owner:
    def __init__(self, taxpayer_name, taxpayer_addr1, taxpayer_addr2, taxpayer_addr3):
        self.taxpayer_name = taxpayer_name
        self.taxpayer_addr1 = taxpayer_addr1
        self.taxpayer_addr2 = taxpayer_addr2
        self.taxpayer_addr3 = taxpayer_addr3

    def __str__(self):
        """Returns a clean, human-readable multiline string layout."""
        return (
            f"Taxpayer Name: {self.taxpayer_name}\n"
            f"Address Line 1: {self.taxpayer_addr1}\n"
            f"Address Line 2: {self.taxpayer_addr2}\n"
            f"Address Line 3: {self.taxpayer_addr3}"
        )

    def to_dict(self):
        """Converts the object attributes into a standard Python dictionary."""
        return {
            "taxpayer_name": self.taxpayer_name,
            "taxpayer_addr1": self.taxpayer_addr1,
            "taxpayer_addr2": self.taxpayer_addr2,
            "taxpayer_addr3": self.taxpayer_addr3
        }

    def to_json(self, indent=4):
        """Converts the object directly into a structured JSON string."""
        return json.dumps(self.to_dict(), indent=indent)
