# -*- coding: UTF-8 -*-
import re
import time
import appModuleHandler
import ui
import api
from logHandler import log
from controlTypes import Role

class AppModule(appModuleHandler.AppModule):
    """
    Duxbury (dbtw.exe) AppModule

    Shortcuts:
      • Alt+-        -> Read full status bar (NVDA built-in first, then fallback).
      • Alt+.        -> Speak only the line number ("Line 12").
      • Alt+,        -> Speak only the page number ("Page 5").
      • Alt+Shift+.  -> Debug: list candidate UI texts and say best-found status.
      • Alt+Shift+,  -> Debug: force UI-scan summary (Page/Line/Col).

    All UI strings and comments are in English.
    """

    __gestures = {
        # Full status line
        "kb:alt+-": "reportDuxburyStatus",
        "kb(laptop):alt+-": "reportDuxburyStatus",
        

        # Line only
        "kb:alt+.": "reportDuxburyLine",
        "kb(laptop):alt+.": "reportDuxburyLine",
        

        # Page only
        "kb:alt+,": "reportDuxburyPage",
        "kb(laptop):alt+,": "reportDuxburyPage",
        

        # Debug
        "kb:alt+shift+.": "debugListCandidates",
        "kb(laptop):alt+shift+.": "debugListCandidates",
        "kb:alt+shift+,": "debugScanSummary",
        "kb(laptop):alt+shift+,": "debugScanSummary",
    }

    # ---------------------------
    # Core retrieval
    # ---------------------------
    def _get_status_text_api(self):
        """Try NVDA API (same source as Insert+End)."""
        for _ in range(3):
            try:
                t = api.getStatusBarText()
                if t and t.strip():
                    return re.sub(r"\s+", " ", t).strip()
            except Exception as e:
                log.debug(f"dbtw: api.getStatusBarText failed: {e!r}")
            time.sleep(0.03)
        return None

    def _iter_children(self, node, max_depth=6, max_nodes=800):
        """Depth-first scan of UI tree with safety limits."""
        count = 0
        def _walk(n, depth):
            nonlocal count
            if not n or depth > max_depth or count >= max_nodes:
                return
            try:
                kids = list(getattr(n, "children", []) or [])
            except Exception:
                kids = []
            for k in kids:
                count += 1
                yield k
                yield from _walk(k, depth+1)
        yield from _walk(node, 0)

    def _collect_candidate_texts(self):
        """
        Collect individual strings from UI nodes that might represent status parts.
        Returns list of (priority, text) — lower priority number means stronger candidate.
        """
        candidates = []

        # 1) Foreground object
        try:
            fg = api.getForegroundObject()
        except Exception as e:
            log.debug(f"dbtw: getForegroundObject failed: {e!r}")
            fg = None

        def _add(txt, prio):
            if not txt:
                return
            s = re.sub(r"\s+", " ", txt).strip()
            if not s:
                return
            candidates.append((prio, s))

        if not fg:
            return candidates

        # 2) Walk subtree
        for n in self._iter_children(fg, max_depth=6, max_nodes=800):
            try:
                role = getattr(n, "role", None)
                # Prefer actual STATUSBAR nodes
                if role == Role.STATUSBAR:
                    for attr in ("name", "value", "windowText", "description"):
                        _add(getattr(n, attr, None), 0)
                else:
                    # Other text-like nodes (panes, static texts) can hold status chunks
                    basePrio = 2
                    # If the window class looks like a status bar, give it higher priority
                    wc = getattr(n, "windowClassName", "") or ""
                    if isinstance(wc, str) and wc.lower() in ("msctls_statusbar32", "statusbarwindow32", "tstatusbar"):
                        basePrio = 1

                    for attr in ("name", "value", "windowText", "description"):
                        _add(getattr(n, attr, None), basePrio)
            except Exception:
                continue

        # Sort by priority, keep unique order
        seen = set()
        unique = []
        for pr, s in sorted(candidates, key=lambda x: x[0]):
            if s not in seen:
                unique.append((pr, s))
                seen.add(s)
        return unique

    # ---------------------------
    # Parsing
    # ---------------------------
    _PAGE_PATTERNS = [
        r"\b(Page|Pg|P|Stranica|Str)\b\s*[:#.=]?\s*(\d+)\b",
        r"\b(Page)\s+(\d+)\s+of\b",
        r"\bP\s*=\s*(\d+)\b", r"\bP(\d+)\b",
        r"\bStr\s*=\s*(\d+)\b", r"\bStr(\d+)\b",
    ]
    _LINE_PATTERNS = [
        r"\b(Line|Ln|Row|Linija|Redak|Retka)\b\s*[:#.=]?\s*(\d+)\b",
        r"\b(Line)\s+(\d+)\s+of\b",
        r"\b(L|R)\b\s*[:#.=]?\s*(\d+)\b",   # L:12 / R:12 (Croatian "Redak")
        r"\bL\s*=\s*(\d+)\b", r"\bL(\d+)\b", r"\bLn\.?\s*[:#.=]?\s*(\d+)\b",
        r"\b(Red(?:ak)?)\b\s*[:#.=]?\s*(\d+)\b",
    ]
    _COL_PATTERNS = [
        r"\b(Column|Col|Cell|Kolona|Stupac|Kol|Stu)\b\s*[:#.=]?\s*(\d+)\b",
        r"\b(C)\s*[:#.=]?\s*(\d+)\b", r"\bC\s*=\s*(\d+)\b", r"\bC(\d+)\b",
    ]

    def _match_first_number(self, text, patterns):
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                for g in reversed(m.groups()):
                    if g and str(g).isdigit():
                        return g
        return None

    def _parse_from_any(self, texts):
        """
        Try to parse Page/Line/Col from a sequence of UI texts.
        We first try each text individually (stronger signal), then try the concatenated string.
        """
        out = {}
        # 1) Individual texts
        for _, t in texts:
            if "page" not in out:
                n = self._match_first_number(t, self._PAGE_PATTERNS)
                if n: out["page"] = n
            if "line" not in out:
                n = self._match_first_number(t, self._LINE_PATTERNS)
                if n: out["line"] = n
            if "col" not in out:
                n = self._match_first_number(t, self._COL_PATTERNS)
                if n: out["col"] = n
            if "page" in out and "line" in out and "col" in out:
                break

        if "page" in out and "line" in out and "col" in out:
            return out

        # 2) Concatenated fallback (helps when pieces are in one label)
        big = " | ".join(t for _, t in texts)
        if "page" not in out:
            n = self._match_first_number(big, self._PAGE_PATTERNS)
            if n: out["page"] = n
        if "line" not in out:
            n = self._match_first_number(big, self._LINE_PATTERNS)
            if n: out["line"] = n
        if "col" not in out:
            n = self._match_first_number(big, self._COL_PATTERNS)
            if n: out["col"] = n

        # 3) Positional heuristic (Page, Line, Col) if still missing
        nums = re.findall(r"\d+", big)
        if len(nums) >= 3:
            out.setdefault("page", nums[0])
            out.setdefault("line", nums[1])
            out.setdefault("col",  nums[2])

        return out

    # ---------------------------
    # Scripts
    # ---------------------------
    def script_reportDuxburyStatus(self, gesture):
        """Speak the entire status bar, like Insert+End."""
        # Try NVDA's built-in
        try:
            import globalCommands
            globalCommands.commands.script_reportStatusLine(None)
            return
        except Exception as e:
            log.debug(f"dbtw Alt+-: built-in reportStatusLine failed: {e!r}")

        # Fallback to API text
        text = self._get_status_text_api()
        if text:
            ui.message(text)
            return

        # As a last resort, compose from UI scan
        cand = self._collect_candidate_texts()
        parsed = self._parse_from_any(cand)
        if parsed:
            parts = []
            if "page" in parsed: parts.append(f"Page {parsed['page']}")
            if "line" in parsed: parts.append(f"Line {parsed['line']}")
            if "col"  in parsed: parts.append(f"Column {parsed['col']}")
            if parts:
                ui.message(", ".join(parts))
                return

        ui.message("Status bar not available.")

    script_reportDuxburyStatus.__doc__ = "Read status bar (Alt+-)."

    def script_reportDuxburyLine(self, gesture):
        """Speak only the current line number ("Line 12")."""
        # 1) Try API
        text = self._get_status_text_api()
        if text:
            n = self._match_first_number(text, self._LINE_PATTERNS)
            if not n:
                # fallback positional when API returns combined text
                nums = re.findall(r"\d+", text)
                if len(nums) == 3: n = nums[1]
            if n:
                ui.message(f"Line {n}")
                return

        # 2) UI-scan
        cand = self._collect_candidate_texts()
        parsed = self._parse_from_any(cand)
        if "line" in parsed:
            ui.message(f"Line {parsed['line']}")
            return

        ui.message("Line number not available.")

    script_reportDuxburyLine.__doc__ = "Speak only the line number (Alt+.)."

    def script_reportDuxburyPage(self, gesture):
        """Speak only the current page number ("Page 5")."""
        # 1) Try API
        text = self._get_status_text_api()
        if text:
            n = self._match_first_number(text, self._PAGE_PATTERNS)
            if not n:
                nums = re.findall(r"\d+", text)
                if len(nums) >= 1: n = nums[0]
            if n:
                ui.message(f"Page {n}")
                return

        # 2) UI-scan
        cand = self._collect_candidate_texts()
        parsed = self._parse_from_any(cand)
        if "page" in parsed:
            ui.message(f"Page {parsed['page']}")
            return

        ui.message("Page number not available.")

    script_reportDuxburyPage.__doc__ = "Speak only the page number (Alt+,)."

    # ---------------------------
    # Debug helpers
    # ---------------------------
    def debugListCandidates(self, gesture):
        """Speak & log the first found status summary and list up to 30 raw candidates in the log."""
        cand = self._collect_candidate_texts()
        parsed = self._parse_from_any(cand)
        log.debug("dbtw debug: ----- candidate texts begin -----")
        for i, (_, s) in enumerate(cand[:30], 1):
            log.debug(f"dbtw cand {i:02d}: {s!r}")
        log.debug("dbtw debug: ----- candidate texts end -----")
        if parsed:
            msg = []
            if "page" in parsed: msg.append(f"Page {parsed['page']}")
            if "line" in parsed: msg.append(f"Line {parsed['line']}")
            if "col"  in parsed: msg.append(f"Column {parsed['col']}")
            ui.message(", ".join(msg))
        else:
            ui.message("No status candidates found.")

    def debugScanSummary(self, gesture):
        """Force a UI-scan and speak the summary Page/Line/Col if any."""
        cand = self._collect_candidate_texts()
        parsed = self._parse_from_any(cand)
        if parsed:
            parts = []
            if "page" in parsed: parts.append(f"Page {parsed['page']}")
            if "line" in parsed: parts.append(f"Line {parsed['line']}")
            if "col"  in parsed: parts.append(f"Column {parsed['col']}")
            ui.message(", ".join(parts))
        else:
            ui.message("UI scan did not find status information.")
