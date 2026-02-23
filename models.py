"""Dataclasses for the printer key tracker."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Final


EVENT_COLUMNS: Final[tuple[str, ...]] = (
    "EventId",
    "Timestamp",
    "UserId",
    "Action",
    "KeyId",
    "FromLocation",
    "ToLocation",
    "PrinterOrDestination",
    "ReturnedToLocation",
    "Notes",
)


@dataclass(frozen=True)
class KeyEvent:
    EventId: str
    Timestamp: str
    UserId: str
    Action: str
    KeyId: str
    FromLocation: str
    ToLocation: str
    PrinterOrDestination: str
    ReturnedToLocation: str
    Notes: str

    def to_csv_row(self) -> dict[str, str]:
        return {
            "EventId": self.EventId,
            "Timestamp": self.Timestamp,
            "UserId": self.UserId,
            "Action": self.Action,
            "KeyId": self.KeyId,
            "FromLocation": self.FromLocation,
            "ToLocation": self.ToLocation,
            "PrinterOrDestination": self.PrinterOrDestination,
            "ReturnedToLocation": self.ReturnedToLocation,
            "Notes": self.Notes,
        }


@dataclass(frozen=True)
class OutstandingKey:
    KeyId: str
    CheckedOutBy: str
    TimeOut: str
    ToLocation: str
    PrinterOrDestination: str

