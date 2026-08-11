import os
from enum import Enum
from dataclasses import dataclass
from typing import Final

class PropertyType(Enum):
    PERSONAL = "pp"
    REAL = "real"

@dataclass(frozen=True)
class JobConfiguration:
    template_path: str
    folder_suffix: str
    display_name: str  # 👈 Added this missing declaration
    file_prefix: str   # 👈 Added this missing declaration

# Centralize template references and naming conventions
JOB_REGISTRY: Final[dict[PropertyType, JobConfiguration]] = {
    PropertyType.PERSONAL: JobConfiguration(
        template_path=os.path.abspath("../assets/templates/pp.docx"),
        folder_suffix="pp_generated_letters",
        display_name="Personal_Property",  # 👈 Maps to display_name field above
        file_prefix="PP"                   # 👈 Maps to file_prefix field above
    ),
    PropertyType.REAL: JobConfiguration(
        template_path=os.path.abspath("../assets/templates/real.docx"),
        folder_suffix="real_generated_letters",
        display_name="Real_Property",      # 👈 Maps to display_name field above
        file_prefix="REAL"                 # 👈 Maps to file_prefix field above
    )
}
