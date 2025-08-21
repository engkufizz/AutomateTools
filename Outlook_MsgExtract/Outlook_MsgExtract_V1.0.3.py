import os
import re
import glob
from datetime import datetime
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    from backports.zoneinfo import ZoneInfo  # pip install backports.zoneinfo
from email.utils import parseaddr

import extract_msg
from bs4 import BeautifulSoup
from dateutil import parser as dtparser

# Configuration
INCLUDE_SUBJECT = True
KL_TZ = ZoneInfo("Asia/Kuala_Lumpur")
HEADER_KEYS = ("From", "Sent", "To", "Cc", "CC", "Subject")

# ---------- Utilities ----------
def html_to_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(separator="\n")
    lines = [ln.rstrip() for ln in text.splitlines()]
    # Limit to max two consecutive blank lines
    out, blanks = [], 0
    for ln in lines:
        if ln.strip():
            blanks = 0
            out.append(ln)
        else:
            blanks += 1
            if blanks <= 2:
                out.append("")
    return "\n".join(out).strip()

def choose_body(msg_obj) -> str:
    try:
        html = getattr(msg_obj, "htmlBody", None) or getattr(msg_obj, "bodyHTML", None)
    except Exception:
        html = None
    if html:
        return html_to_text(html)
    body = getattr(msg_obj, "body", "") or ""
    return body.replace("\r\n", "\n").strip()

def clean_dt_string(s: str) -> str:
    if not s:
        return s
    s = re.sub(r"(\d+)\s*(st|nd|rd|th)\b", r"\1", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalise_dt(dt) -> datetime | None:
    if not dt:
        return None
    if isinstance(dt, datetime):
        d = dt
    else:
        try:
            d = dtparser.parse(clean_dt_string(str(dt)))
        except Exception:
            return None
    if d.tzinfo is None:
        d = d.replace(tzinfo=KL_TZ)
    try:
        d = d.astimezone(KL_TZ)
    except Exception:
        pass
    return d

def clean_name(from_text: str) -> str:
    val = re.sub(r"\s+", " ", (from_text or "")).strip()
    val = val.replace("< ", "<").replace(" >", ">")
    name, email = parseaddr(val)
    if name:
        return name.strip()
    if email:
        return email.strip()
    return re.sub(r"[<>]", "", val).strip() or "Unknown"

def get_line_positions(text: str):
    lines = text.splitlines(True)  # keep endings
    starts, pos = [], 0
    for ln in lines:
        starts.append(pos)
        pos += len(ln)
    return lines, starts

def collect_multiline_header_value(lines, start_idx):
    m = re.match(r"^\s*([A-Za-z]+)\s*:\s*(.*)$", lines[start_idx])
    if not m:
        return "", start_idx + 1
    first_val = m.group(2).rstrip("\r\n")
    buf = [first_val] if first_val else []
    i = start_idx + 1
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
            break
        if re.match(r"^\s*(From|Sent|To|Cc|CC|Subject)\s*:", line, re.IGNORECASE):
            break
        if line.startswith((" ", "\t")):
            buf.append(line.strip())
            i += 1
        else:
            break
    return " ".join(x for x in buf if x).strip(), i

def parse_header_block(block_text: str) -> dict:
    lines = block_text.splitlines()
    info = {"From": None, "Sent": None}
    i = 0
    while i < len(lines):
        if re.match(r"^\s*(From|Sent|To|Cc|CC|Subject)\s*:", lines[i], re.IGNORECASE):
            key_m = re.match(r"^\s*([A-Za-z]+)\s*:", lines[i])
            key = key_m.group(1).title() if key_m else None
            val, j = collect_multiline_header_value(lines, i)
            if key in info and val and not info.get(key):
                info[key] = val
            i = j
        else:
            i += 1
    return info

def find_header_boundaries(text: str):
    lines, starts = get_line_positions(text)
    boundaries = []
    n = len(lines)
    for i, line in enumerate(lines):
        if re.match(r"^\s*From\s*:", line, re.IGNORECASE):
            subj_idx = None
            for k in range(i, min(n, i + 60)):
                if re.match(r"^\s*Subject\s*:", lines[k], re.IGNORECASE):
                    subj_idx = k
                    break
            if subj_idx is not None:
                start_pos = starts[i]
                end_k = subj_idx
                k2 = subj_idx + 1
                while k2 < n and lines[k2].startswith((" ", "\t")):
                    end_k = k2
                    k2 += 1
                end_pos = starts[end_k] + len(lines[end_k])
                header_block_text = "".join(lines[i:end_k + 1])
                boundaries.append((start_pos, end_pos, header_block_text))
    for m in re.finditer(r"\n-+\s*Original Message\s*-+\s*\n", text, flags=re.IGNORECASE):
        boundaries.append((m.start(), m.start(), ""))
    boundaries.sort(key=lambda x: x[0])
    return boundaries

def strip_leading_headers(text: str) -> str:
    lines = text.splitlines()
    if not lines:
        return text
    if any(re.match(rf"^\s*{k}\s*:", lines[0], re.IGNORECASE) for k in HEADER_KEYS):
        i = 0
        while i < len(lines):
            if not lines[i].strip():
                i += 1
                break
            if re.match(r"^\s*(From|Sent|To|Cc|CC|Subject)\s*:", lines[i], re.IGNORECASE):
                _, j = collect_multiline_header_value(lines, i)
                i = j
            else:
                break
        return "\n".join(lines[i:]).strip()
    return text.strip()

def remove_header_blocks_anywhere(text: str) -> str:
    lines = text.splitlines()
    out, i = [], 0
    while i < len(lines):
        if re.match(r"^\s*(From|Sent|To|Cc|CC|Subject)\s*:", lines[i], re.IGNORECASE):
            _, j = collect_multiline_header_value(lines, i)
            while j < len(lines):
                if not lines[j].strip():
                    j += 1
                    break
                if re.match(r"^\s*(From|Sent|To|Cc|CC|Subject)\s*:", lines[j], re.IGNORECASE):
                    _, j = collect_multiline_header_value(lines, j)
                elif lines[j].startswith((" ", "\t")):
                    j += 1
                else:
                    break
            i = j
        else:
            out.append(lines[i].rstrip())
            i += 1
    text2 = "\n".join(out).strip()
    return re.sub(r"\n{3,}", "\n\n", text2)

# ---------- Reflow helpers ----------
def is_bullet_line(line: str) -> bool:
    return bool(re.match(r"^\s*(?:[-*•–—]\s+|\d+[.)]\s+|[A-Za-z]\)\s+)", line))

def is_heading_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if s.startswith("[") and s.endswith("]"):
        return True
    # Whitelisted headings (add more if needed)
    if re.fullmatch(r"(Summary|Symptoms|Impact|Actions?|Action Plan|Root Cause|Resolution|Findings|Next Steps|Notes|Conclusion|Update)", s, re.IGNORECASE):
        return True
    # Explicitly NOT headings
    if re.match(r"^\s*(dear|hi|hello)\b", s, re.IGNORECASE):
        return False
    if re.match(r"^\s*(SR|RE|FW)\b", s, re.IGNORECASE):
        return False
    # Avoid treating any line that contains digits as a heading
    if re.search(r"\d", s):
        return False
    return False  # be conservative

def remove_banners_and_disclaimers(text: str) -> str:
    lines = text.splitlines()
    out, i = [], 0
    while i < len(lines):
        ln = lines[i]
        if re.search(r"^\s*CAUTION\s*:", ln, re.IGNORECASE) or \
           "This email originated from outside the organization" in ln:
            i += 1
            while i < len(lines) and lines[i].strip():
                i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1
            continue
        out.append(ln)
        i += 1
    return "\n".join(out).strip()

def should_join_through_blank(prev: str, nxt: str) -> bool:
    if is_heading_line(prev) or is_heading_line(nxt):
        return False
    # Build up the salutation only (Dear/Hi/Hello -> mention or comma)
    if re.match(r"^\s*(dear|hi|hello)\b", prev, re.IGNORECASE):
        if re.match(r"^@", nxt) or re.fullmatch(r"[,.;:!?]", nxt):
            return True
        return False
    # Don’t join if previous ends a sentence
    if re.search(r"[.!?]\s*$", prev):
        return False
    # Join if next is purely punctuation or starts with punctuation/mention/&, (, [, digits, SR id etc.
    if re.fullmatch(r"[,.;:!?]", nxt) or re.match(r"^[,.;:!?@&(\[\d]|^SR\b", nxt, re.IGNORECASE):
        return True
    # Join if previous ends with conjunction/preposition or dangling & or /
    if re.search(r"\b(of|for|to|and|or|with|in|on|at|by|from|as|that|than|but|because|so)\s*$", prev, re.IGNORECASE):
        return True
    if re.search(r"[&/]\s*$", prev):
        return True
    # Otherwise keep the blank as a paragraph break
    return False

def reflow_paragraphs(text: str) -> str:
    raw = [l.rstrip() for l in (text or "").splitlines()]
    # Remove “soft” blanks that actually belong to the same sentence
    soft, i = [], 0
    while i < len(raw):
        cur = raw[i]
        if cur.strip() == "":
            prev = next((soft[j] for j in range(len(soft) - 1, -1, -1) if soft[j] != ""), "")
            j, nxt = i + 1, ""
            while j < len(raw):
                if raw[j].strip():
                    nxt = raw[j].strip()
                    break
                j += 1
            if prev and nxt and should_join_through_blank(prev.strip(), nxt):
                i += 1
                continue
            else:
                soft.append("")
                i += 1
        else:
            soft.append(cur)
            i += 1
    # Join consecutive non-empty lines into paragraphs; preserve bullets/headings
    paras, buf = [], ""
    def flush():
        nonlocal buf
        if buf:
            p = re.sub(r"\s+([,.;:!?])", r"\1", buf)
            p = re.sub(r"\(\s+", "(", p)
            p = re.sub(r"\s+\)", ")", p)
            p = re.sub(r"\s+/\s+", "/", p)
            p = re.sub(r"\s+", " ", p).strip()
            paras.append(p)
            buf = ""
    for ln in soft:
        s = ln.strip()
        if not s:
            flush()
            continue
        if is_heading_line(s) or is_bullet_line(s):
            flush()
            paras.append(s)
            continue
        if re.fullmatch(r"[,.;:!?]", s) or re.match(r"^[,.;:!?]", s):
            buf = (buf.rstrip() + s) if buf else s
            continue
        if not buf:
            buf = s
        else:
            buf += s if buf.endswith(("-", "/")) else " " + s
    flush()
    return "\n\n".join(paras).strip()

# ---------- Thread splitting and extraction ----------
def split_thread(text: str):
    t = (text or "").replace("\r\n", "\n")
    boundaries = find_header_boundaries(t)
    segments = []
    if boundaries:
        first_start, _, _ = boundaries[0]
        top_body = t[:first_start].strip()
        if top_body:
            segments.append({"from_raw": None, "sent_raw": None, "body": strip_leading_headers(top_body)})
        for i, (beg, end, header_text) in enumerate(boundaries):
            nxt_beg = boundaries[i + 1][0] if i + 1 < len(boundaries) else len(t)
            body = t[end:nxt_beg].strip()
            body = strip_leading_headers(body)
            hdr = parse_header_block(header_text) if header_text else {}
            segments.append({"from_raw": hdr.get("From"), "sent_raw": hdr.get("Sent"), "body": body})
    else:
        segments.append({"from_raw": None, "sent_raw": None, "body": strip_leading_headers(t)})
    return [s for s in segments if s["body"]]

def extract_messages_from_msg(path: str):
    msg = extract_msg.Message(path)
    subject = (msg.subject or "").strip()
    sender = (getattr(msg, "sender", None) or getattr(msg, "sender_name", None) or "").strip()
    sent_dt = getattr(msg, "date", None)
    top_dt = normalise_dt(sent_dt)
    full_text = choose_body(msg)
    segments = split_thread(full_text)
    messages = []
    for idx, seg in enumerate(segments):
        author, dt_val = None, None
        if seg["from_raw"]:
            author = clean_name(seg["from_raw"])
        if seg["sent_raw"]:
            dt_val = normalise_dt(seg["sent_raw"])
        if idx == 0:
            if not author and sender:
                author = clean_name(sender)
            if not dt_val and top_dt:
                dt_val = top_dt
        if not author:
            m = re.search(r"^\s*From\s*:\s*(.+)$", seg["body"], re.IGNORECASE | re.MULTILINE)
            if m:
                author = clean_name(m.group(1))
        if not dt_val:
            m = re.search(r"^\s*Sent\s*:\s*(.+)$", seg["body"], re.IGNORECASE | re.MULTILINE)
            if m:
                dt_val = normalise_dt(m.group(1))
        body_clean = remove_header_blocks_anywhere(seg["body"])
        body_clean = remove_banners_and_disclaimers(body_clean)
        body_clean = reflow_paragraphs(body_clean)
        author = author or "Unknown"
        time_str = dt_val.strftime("%Y-%m-%d %H:%M:%S") if isinstance(dt_val, datetime) else "Unknown"
        messages.append({"author": author, "time": time_str, "dt_sort": dt_val, "body": body_clean})
    messages.sort(key=lambda m: (m["dt_sort"] is None, m["dt_sort"]), reverse=True)
    for m in messages:
        m.pop("dt_sort", None)
    return subject, messages

def format_output(subject: str, messages):
    lines = []
    if INCLUDE_SUBJECT:
        lines.append(f"Subject: {subject}")
        lines.append("----------------------------------------------------------------------")
    for i, m in enumerate(messages, start=1):
        lines.append("")
        lines.append(f"{i:03d}. Author: {m['author']}")
        lines.append(f"     Time:   {m['time']}")
        lines.append("     Message:")
        for ln in (m["body"] or "").splitlines():
            lines.append(f"       {ln.rstrip()}")
    return "\n".join(lines).rstrip() + "\n"

# ---------- CLI selection ----------
def prompt_select(files):
    print("Select .msg file(s) to process:")
    for i, f in enumerate(files, 1):
        print(f"  {i:03d}. {f}")
    while True:
        choice = input("Enter number(s) (e.g., 1 or 1,3). Press Enter to cancel: ").strip()
        if not choice:
            print("No selection. Exiting.")
            return []
        tokens = re.split(r"[,\s]+", choice)
        idxs, ok = [], True
        for tok in tokens:
            if not tok:
                continue
            if not tok.isdigit():
                ok = False; break
            n = int(tok)
            if not (1 <= n <= len(files)):
                ok = False; break
            idxs.append(n - 1)
        if not ok or not idxs:
            print("Invalid input. Please enter valid number(s) from the list.")
            continue
        seen, selected = set(), []
        for i in idxs:
            if i not in seen:
                seen.add(i); selected.append(files[i])
        return selected

def main():
    files = sorted(glob.glob("*.msg"))
    if not files:
        print("No .msg files found in the current directory.")
        return
    selected_files = prompt_select(files)
    if not selected_files:
        return
    for msg_path in selected_files:
        try:
            subject, messages = extract_messages_from_msg(msg_path)
            text_out = format_output(subject, messages)
            base = os.path.splitext(os.path.basename(msg_path))[0]
            out_path = f"{base}_parsed.txt"
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(text_out)
            print(f"Wrote: {out_path}")
        except Exception as e:
            print(f"Failed to parse {msg_path}: {e}")

if __name__ == "__main__":
    main()
