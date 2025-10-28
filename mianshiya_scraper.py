"""Utility to download interview questions from https://www.mianshiya.com/.

This module provides a simple command line interface that pulls every
question exposed by the public Mianshiya question bank API and writes the
result into an Excel workbook.  Only the question text is persisted (answers
and other metadata are ignored).

The default configuration targets the undocumented but publicly available
`/api/question/list/page` endpoint that backs the Mianshiya web client.  The
script keeps requesting pages until the service stops returning new records.

Example usage
-------------

```bash
python mianshiya_scraper.py --output questions.xlsx
```

To turn the script into a one-click executable on Windows, install the
dependencies (``pip install -r requirements.txt``) and then run::

    pyinstaller --onefile mianshiya_scraper.py

The generated ``dist/mianshiya_scraper.exe`` can be double-clicked in
Explorer.  Optional command line flags can be stored in a ``.bat`` file next
to the executable if required.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from typing import Iterable, List, Optional

import requests
from openpyxl import Workbook


API_URL = "https://www.mianshiya.com/api/question/list/page"


@dataclass
class Question:
    """Represents a question extracted from the Mianshiya API."""

    text: str

    @classmethod
    def from_api(cls, raw: dict) -> Optional["Question"]:
        """Best-effort conversion from a raw API payload to :class:`Question`.

        The API is undocumented and subject to change, so we check a handful
        of key names that have appeared historically in the Mianshiya JSON
        responses.  If none of them contain usable text, ``None`` is returned
        and the record is skipped.
        """

        candidate_fields: Iterable[str] = (
            "title",
            "content",
            "questionContent",
            "questionTitle",
            "questionText",
        )

        for field in candidate_fields:
            value = raw.get(field)
            if isinstance(value, str) and value.strip():
                return cls(text=value.strip())

        return None


class MianshiyaClient:
    """Lightweight client that pages through the Mianshiya question API."""

    def __init__(self, page_size: int = 100, timeout: int = 30) -> None:
        self.page_size = page_size
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": "Mozilla/5.0 (compatible; MianshiyaScraper/1.0)",
                "Accept": "application/json, text/plain, */*",
                "Referer": "https://www.mianshiya.com/",
            }
        )

    def fetch_all_questions(self) -> List[Question]:
        """Retrieve every question exposed by the API."""

        questions: List[Question] = []
        current_page = 1

        while True:
            payload = {
                "current": current_page,
                "pageSize": self.page_size,
                "sortField": "createTime",
                "sortOrder": "descend",
            }

            response = self.session.post(API_URL, json=payload, timeout=self.timeout)
            response.raise_for_status()
            body = response.json()

            records = self._extract_records(body)
            if not records:
                break

            for item in records:
                question = Question.from_api(item)
                if question is not None:
                    questions.append(question)

            if len(records) < self.page_size:
                break

            current_page += 1

        return questions

    @staticmethod
    def _extract_records(body: dict) -> List[dict]:
        """Pull the record list from a Mianshiya API response."""

        if not isinstance(body, dict):
            return []

        data = body.get("data")
        if isinstance(data, dict):
            records = data.get("records") or data.get("list")
        else:
            records = None

        if isinstance(records, list):
            return [item for item in records if isinstance(item, dict)]

        return []


def write_questions_to_excel(questions: Iterable[Question], output_path: str) -> None:
    """Persist questions to a single-column Excel file."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Questions"
    sheet.append(["Question"])

    for question in questions:
        sheet.append([question.text])

    workbook.save(output_path)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Download Mianshiya questions")
    parser.add_argument(
        "--output",
        default="mianshiya_questions.xlsx",
        help="Path to the Excel file to create (default: %(default)s)",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=100,
        help="Number of items to request per page (default: %(default)s)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="HTTP timeout in seconds (default: %(default)s)",
    )

    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    client = MianshiyaClient(page_size=args.page_size, timeout=args.timeout)

    try:
        questions = client.fetch_all_questions()
    except requests.HTTPError as exc:  # pragma: no cover - network errors are runtime-only
        print(f"Request failed: {exc}", file=sys.stderr)
        return 1
    except requests.RequestException as exc:  # pragma: no cover - runtime-only
        print(f"Network error: {exc}", file=sys.stderr)
        return 1

    if not questions:
        print("No questions were retrieved.  Check your network connection or the API.")
        return 1

    write_questions_to_excel(questions, args.output)

    print(f"Exported {len(questions)} questions to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
