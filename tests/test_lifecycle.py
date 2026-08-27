"""孵化期 / 长存筛选单元测试。"""
from __future__ import annotations

import os
import unittest
from datetime import datetime, timedelta
from unittest import mock

from outlook_api_reg import lifecycle


class TestIncubationMath(unittest.TestCase):
    def test_default_hours(self):
        with mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop("OUTLOOK_INCUBATION_HOURS", None)
            self.assertEqual(lifecycle.incubation_hours(), 48.0)

    def test_env_hours(self):
        with mock.patch.dict(os.environ, {"OUTLOOK_INCUBATION_HOURS": "24"}):
            self.assertEqual(lifecycle.incubation_hours(), 24.0)

    def test_env_invalid_falls_back(self):
        with mock.patch.dict(os.environ, {"OUTLOOK_INCUBATION_HOURS": "abc"}):
            self.assertEqual(lifecycle.incubation_hours(), 48.0)

    def test_parse_created_at_iso(self):
        dt = lifecycle.parse_created_at("2026-08-20T12:00:00")
        self.assertEqual(dt, datetime(2026, 8, 20, 12, 0, 0))

    def test_is_incubating_inside_window(self):
        created = datetime(2026, 8, 25, 10, 0, 0)
        now = datetime(2026, 8, 26, 10, 0, 0)  # +24h
        self.assertTrue(lifecycle.is_incubating(created, now=now, hours=48))

    def test_is_incubating_expired(self):
        created = datetime(2026, 8, 20, 10, 0, 0)
        now = datetime(2026, 8, 26, 10, 0, 0)  # +6d
        self.assertFalse(lifecycle.is_incubating(created, now=now, hours=48))

    def test_incubation_until(self):
        created = "2026-08-25T10:00:00"
        until = lifecycle.incubation_until(created, hours=48)
        self.assertEqual(until, datetime(2026, 8, 27, 10, 0, 0))

    def test_enrich_lifecycle_fields(self):
        created = (datetime.now() - timedelta(hours=1)).isoformat(timespec="seconds")
        row = {"created_at": created, "tags": []}
        lifecycle.enrich_lifecycle_fields(row, hours=48)
        self.assertTrue(row["incubating"])
        self.assertTrue(row["incubation_until"])
        self.assertIn("incubating", row["tags"])

    def test_enrich_removes_tag_when_expired(self):
        created = (datetime.now() - timedelta(days=5)).isoformat(timespec="seconds")
        row = {"created_at": created, "tags": ["incubating", "keep"]}
        lifecycle.enrich_lifecycle_fields(row, hours=48)
        self.assertFalse(row["incubating"])
        self.assertNotIn("incubating", row["tags"])
        self.assertIn("keep", row["tags"])

    def test_account_age_days(self):
        created = datetime.now() - timedelta(days=10, hours=12)
        age = lifecycle.account_age_days(created)
        self.assertIsNotNone(age)
        self.assertAlmostEqual(age or 0, 10.5, delta=0.02)

    def test_is_long_lived(self):
        old = (datetime.now() - timedelta(days=10)).isoformat(timespec="seconds")
        row = {
            "created_at": old,
            "combo_recovery": "a@b.com----pwd----cid----rt----rec@x.com----recpwd",
            "refresh_token": "rt",
            "recovery_email": "rec@x.com",
        }
        self.assertTrue(lifecycle.is_long_lived(row, min_days=7))

    def test_not_long_lived_when_incubating(self):
        recent = (datetime.now() - timedelta(hours=2)).isoformat(timespec="seconds")
        row = {
            "created_at": recent,
            "combo_recovery": "a@b.com----pwd----cid----rt----rec@x.com----recpwd",
        }
        self.assertFalse(lifecycle.is_long_lived(row, min_days=0))

    def test_not_long_lived_without_recovery(self):
        old = (datetime.now() - timedelta(days=10)).isoformat(timespec="seconds")
        row = {"created_at": old, "combo": "a@b.com----pwd----cid----rt", "refresh_token": "rt"}
        self.assertFalse(lifecycle.is_long_lived(row, min_days=7))


if __name__ == "__main__":
    unittest.main()
