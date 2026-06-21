import logging
import os
import re
from pathlib import Path
from email.utils import parseaddr

from config import LOG_FILE, LOG_LEVEL

EMAIL_REGEX = re.compile(
    r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
)


def setup_logging(name="offer_letters"):
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    level = getattr(logging, LOG_LEVEL, logging.INFO)
    logger.setLevel(level)

    formatter = logging.Formatter("%(asctime)s %(levelname)-8s %(message)s")

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


def safe_filename(value: str) -> str:
    name = re.sub(r"[\\/*:?\"<>|]+", "", value)
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        name = "output"
    return name


def unique_filepath(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 1
    while True:
        candidate = path.with_name(f"{path.stem}-{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def validate_email(address: str) -> bool:
    address = address.strip()
    if not address:
        return False

    parsed = parseaddr(address)[1]
    if not parsed:
        return False

    return bool(EMAIL_REGEX.fullmatch(parsed))


def build_safe_attachment_name(name: str, email: str = None) -> str:
    base = name
    if email and validate_email(email):
        local_part = email.split("@", 1)[0]
        base = f"{name} - {local_part}"
    return safe_filename(base)


def find_matching_file(folder: str, base_name: str, extension: str):
    search_path = Path(folder)
    exact = search_path / f"{base_name}{extension}"
    if exact.exists():
        return exact

    glob_pattern = f"{base_name}*{extension}"
    matches = sorted(search_path.glob(glob_pattern))
    return matches[0] if matches else None
