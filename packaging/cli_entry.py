#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import sys

import daily_wecom_report


def main() -> int:
    return int(daily_wecom_report.main())


if __name__ == "__main__":
    raise SystemExit(main())

