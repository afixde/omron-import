from pathlib import Path
from datetime import datetime
import shutil


def archive_file(file: Path, archive_dir: Path) -> Path:
    """
    Verschiebt eine Datei in das Archiv.

    Existiert bereits eine Datei mit gleichem Namen,
    wird ein Zeitstempel angehängt.
    """

    archive_dir.mkdir(parents=True, exist_ok=True)

    target = archive_dir / file.name

    if target.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = archive_dir / (
            f"{file.stem}_{timestamp}{file.suffix}"
        )

    shutil.move(str(file), str(target))

    return target