from dataclasses import dataclass
from enum import Enum
import os
import sys
from typing import Final


class PropertyType(Enum):
  PERSONAL = "pp"
  REAL = "real"


@dataclass(frozen=True)
class JobConfiguration:
  template_path: str
  folder_suffix: str
  display_name: str
  file_prefix: str


def get_asset_path(relative_path: str) -> str:
  """Resolves the absolute path for both development (PyCharm) and PyInstaller builds."""
  if hasattr(sys, "_MEIPASS"):
    # Inside PyInstaller bundle: assets are directly in the temp root
    return os.path.join(sys._MEIPASS, relative_path)

  # Inside PyCharm: Step up out of the 'src' directory to find the 'assets' root
  base_dir = os.path.dirname(os.path.abspath(__file__))  # Points to 'src'
  project_root = os.path.abspath(
      os.path.join(base_dir, "..")
  )  # Steps out to project root
  return os.path.join(project_root, relative_path)


# Centralize template references using the runtime resolver
JOB_REGISTRY: Final[dict[PropertyType, JobConfiguration]] = {
    PropertyType.PERSONAL: JobConfiguration(
        template_path=get_asset_path("assets/templates/pp.docx"),
        folder_suffix="pp_generated_letters",
        display_name="Personal_Property",
        file_prefix="PP",
    ),
    PropertyType.REAL: JobConfiguration(
        template_path=get_asset_path("assets/templates/real.docx"),
        folder_suffix="real_generated_letters",
        display_name="Real_Property",
        file_prefix="REAL",
    ),
}




