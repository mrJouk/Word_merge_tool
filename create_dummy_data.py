from pathlib import Path

import config as cfg
from app import create_dummy_data


if __name__ == "__main__":
    create_dummy_data(Path(cfg.DEFAULT_KYRGYZ_DIR), Path(cfg.DEFAULT_RUSSIAN_DIR))
    print(f"Created sample files in {cfg.DEFAULT_KYRGYZ_DIR}/ and {cfg.DEFAULT_RUSSIAN_DIR}/")
