import os
from enum import Enum
from dataclasses import dataclass
from typing import Final

class PropertyType(Enum):
    PERSONAL = "PP"
    REAL = "RE"

# Define what each job type requires
@dataclass(frozen=True)
class JobConfiguration:
    template_path: str
    expected_extension: str
    output_prefix: str

# Centralize ALL your changing configurations in one clean map
JOB_REGISTRY: Final[dict[PropertyType, JobConfiguration]] = {
    PropertyType.PERSONAL: JobConfiguration(
        template_path=os.path.abspath("../assets/templates/pp.docx"),
        expected_extension=".xls" ,
        output_prefix="pp_master"
    ),
    PropertyType.REAL: JobConfiguration(
        template_path=os.path.abspath("../assets/templates/re.docx"),
        expected_extension=".json",
        output_prefix="re_master"
    )
}
