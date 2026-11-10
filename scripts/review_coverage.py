"""Validate review coverage without creating or approving review records."""


def complete_page_review(records, page_count, page_key, required_fields):
    """Require one fully checked record for each actual source page."""
    if type(page_count) is not int or page_count < 1:
        return False
    if not isinstance(records, list) or len(records) != page_count:
        return False
    seen = set()
    for record in records:
        if not isinstance(record, dict):
            return False
        page = record.get(page_key)
        if type(page) is not int or not 1 <= page <= page_count or page in seen:
            return False
        if any(record.get(field) is not True for field in required_fields):
            return False
        seen.add(page)
    return len(seen) == page_count
