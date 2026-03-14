"""
Orchestrator: run sync, then run analyse and send Telegram only if new activities were added.
For use in hourly (or similar) workflows instead of separate sync + conditional analyse steps.
"""
from __future__ import annotations

import sync
import analyse


def main() -> None:
    new_count = sync.main(only_notify_on_new=True)
    if new_count > 0:
        analyse.main()


if __name__ == "__main__":
    main()
