from rapidfuzz import fuzz
from doccompare.models import DiffElement, DiffSegment, DiffType, ComparisonResult

MOVE_THRESHOLD = 85.0


class MoveDetector:
    def __init__(self, threshold: float = MOVE_THRESHOLD):
        self.threshold = threshold

    def detect(self, result: ComparisonResult) -> ComparisonResult:
        # Collect all deleted and added segments with their parent elements
        deleted_segs = []
        added_segs = []

        for elem in result.diff_elements:
            for seg in elem.segments:
                if seg.diff_type == DiffType.DELETED and len(seg.text.strip()) > 20:
                    deleted_segs.append(seg)
                elif seg.diff_type == DiffType.ADDED and len(seg.text.strip()) > 20:
                    added_segs.append(seg)

        # Find candidate moves
        candidates = []
        for d in deleted_segs:
            for a in added_segs:
                score = fuzz.ratio(d.text.strip(), a.text.strip())
                if score >= self.threshold:
                    candidates.append((d, a, score))

        candidates.sort(key=lambda x: x[2], reverse=True)

        used_deleted: set = set()
        used_added: set = set()
        move_counter = 0

        for d, a, score in candidates:
            d_id = id(d)
            a_id = id(a)
            if d_id not in used_deleted and a_id not in used_added:
                move_id = f"move_{move_counter}"
                d.diff_type = DiffType.MOVED_FROM
                d.move_id = move_id
                a.diff_type = DiffType.MOVED_TO
                a.move_id = move_id
                used_deleted.add(d_id)
                used_added.add(a_id)
                move_counter += 1

        # Recompute summary
        from .differ import Differ
        result.summary = Differ()._compute_summary(result.diff_elements)
        return result
