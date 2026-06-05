"""
utils/state_manager.py — Persists daily send count to state.json.
Automatically resets when the date changes.
"""

import json
import os
from datetime import date


class StateManager:

    def __init__(self, filepath: str):
        self.filepath = filepath
        os.makedirs(os.path.dirname(filepath), exist_ok=True)

    def load(self) -> dict:
        """
        Load state. Auto-resets if stored date != today.
        Returns: { "date": "YYYY-MM-DD", "count": int }
        """
        today = str(date.today())

        try:
            with open(self.filepath, "r") as f:
                state = json.load(f)

            # New day → reset counter
            if state.get("date") != today:
                state = self._fresh(today)
                self.save(state)

            return state

        except (FileNotFoundError, json.JSONDecodeError):
            state = self._fresh(today)
            self.save(state)
            return state

    def save(self, state: dict):
        """Persist state to disk."""
        with open(self.filepath, "w") as f:
            json.dump(state, f, indent=2)

    def reset(self):
        """Manually reset counter to 0 for today."""
        state = self._fresh(str(date.today()))
        self.save(state)

    def increment(self):
        """Increment counter by 1 and save."""
        state = self.load()
        state["count"] += 1
        self.save(state)
        return state["count"]

    @staticmethod
    def _fresh(today: str) -> dict:
        return {"date": today, "count": 0}
