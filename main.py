"""
Combined run: sync + analyse.
For manual execution or future "vad är min status" Telegram bot trigger.
"""
from __future__ import annotations

import sync
import analyse


def main() -> None:
    sync.main()
    analyse.main()


if __name__ == "__main__":
    main()
