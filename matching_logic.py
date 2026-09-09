"""
Shared match-evidence scoring logic.

Used by:
  - matching_pipeline.py     (offline pass over a finished results.jsonl)
  - orbis_candidate_scraper.py (live, in-loop, to decide when to stop early)

Keeping this in one module means the "requires >=2 independent references"
rule can never drift between the live scraper and the offline audit pass.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Any

COUNTRY_TO_ISO2 = {
    'united kingdom': 'GB', 'france': 'FR', 'spain': 'ES', 'germany': 'DE',
    'sweden': 'SE', 'netherlands': 'NL', 'switzerland': 'CH', 'portugal': 'PT',
    'finland': 'FI', 'italy': 'IT', 'norway': 'NO', 'poland': 'PL',
    'austria': 'AT', 'belgium': 'BE', 'estonia': 'EE', 'ireland': 'IE',
    'denmark': 'DK', 'united states': 'US', 'luxembourg': 'LU',
    'czechia': 'CZ', 'lithuania': 'LT',
}

STRONG_NAME_SIM = 0.85
WEAK_NAME_SIM = 0.60
MIN_WITNESSES = 2


def norm(s: Any) -> str:
    if s is None:
        return ''
    s = str(s).strip().lower()
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def norm_id(s: Any) -> str:
    if s is None:
        return ''
    return re.sub(r'[^A-Za-z0-9]', '', str(s)).upper()


def name_similarity(a: str, b: str) -> float:
    a, b = norm(a), norm(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def split_person_names(value: Any) -> list[str]:
    cleaned = re.sub(r'\+\s*\d+\s+more\b', '', str(value or ''), flags=re.IGNORECASE)
    people: list[str] = []
    for part in re.split(r'[;,/]', cleaned):
        name = re.sub(r'\(.*?\)', '', part)
        name = re.sub(r'^\s*(mr|mrs|ms|dr|prof)\.?\s+', '', name, flags=re.IGNORECASE)
        name = re.sub(r'\s+', ' ', name).strip()
        if name and len(norm(name)) >= 3:
            people.append(name)
    return people


def person_token_subset_match(a: Any, b: Any) -> bool:
    a_tokens = {token for token in norm(a).split() if len(token) >= 2}
    b_tokens = {token for token in norm(b).split() if len(token) >= 2}
    if len(a_tokens) < 2 or len(b_tokens) < 2:
        return False
    smaller, larger = (a_tokens, b_tokens) if len(a_tokens) <= len(b_tokens) else (b_tokens, a_tokens)
    return smaller.issubset(larger)


def normalized_containment_match(a: Any, b: Any, min_length: int = 4) -> bool:
    a_norm = norm(a)
    b_norm = norm(b)
    if len(a_norm) < min_length or len(b_norm) < min_length:
        return False
    return a_norm in b_norm or b_norm in a_norm


def reference_year(value: Any) -> int | None:
    if value in (None, ''):
        return None
    year = getattr(value, 'year', None)
    if isinstance(year, int) and 1800 <= year <= 2100:
        return year
    match = re.search(r'\b(18|19|20|21)\d{2}\b', str(value))
    return int(match.group(0)) if match else None


def incorporation_year(cand: dict) -> int | None:
    value = cand.get('date_of_incorporation')
    if not value:
        report_fields = cand.get('report_fields') or {}
        if isinstance(report_fields, dict):
            value = next(
                (
                    field_value
                    for field_name, field_value in report_fields.items()
                    if 'date of incorporation' in str(field_name).casefold()
                ),
                '',
            )
    match = re.search(r'\b\d{1,2}/\d{1,2}/((?:18|19|20|21)\d{2})\b', str(value or ''))
    return int(match.group(1)) if match else None


@dataclass
class Evidence:
    witnesses: list[str] = field(default_factory=list)
    weak_notes: list[str] = field(default_factory=list)

    @property
    def count(self) -> int:
        return len(self.witnesses)

    @property
    def decided(self) -> bool:
        return self.count >= MIN_WITNESSES


@dataclass
class MatchDecision:
    selected: dict | None
    evidence: Evidence | None
    decided: bool
    matched_via: str
    reason_1: str
    reason_2: str
    comments: str
    best_candidate_if_no_match: dict | None = None
    tied_candidates: list[dict] | None = None


def score_candidate(ref_row: dict, cand: dict) -> Evidence:
    """
    ref_row: dict-like with keys Company name, City, HQ city, ISO2, Website,
             Zipcode, Cand_BvD_guess, NationalID_norm, Founders, Launch year
    cand:    dict with keys bvdid, city, country, candidate_name, snapshot_text
             (optional: website, management - only present once report-page
             scraping has run for that candidate)
    """
    ev = Evidence()

    cand_bvdid = norm_id(cand.get('bvdid'))
    cand_city = norm(cand.get('city'))
    cand_country = norm(cand.get('country'))
    snapshot = ' '.join(
        str(cand.get(key) or '')
        for key in ('snapshot_text', 'report_text', 'full_address', 'management', 'domain', 'website')
    )
    report_fields = cand.get('report_fields') or {}
    if isinstance(report_fields, dict):
        snapshot += ' ' + ' '.join(str(value or '') for value in report_fields.values())

    ref_bvd_guess = norm_id(ref_row.get('Cand_BvD_guess'))
    ref_city = norm(ref_row.get('City') or ref_row.get('HQ city'))
    ref_iso2 = norm(ref_row.get('ISO2'))
    ref_website = str(ref_row.get('Website') or '')
    ref_zip = norm_id(ref_row.get('Zipcode'))

    # 1. Registry / BvD ID exact match. NOTE: Orbis's candidate-row "NationalId"
    # field uses a different numbering scheme than our reference trade-register
    # number (confirmed during validation) - only the derived BvD ID
    # equality is a reliable witness, do not compare NationalId directly.
    if ref_bvd_guess and cand_bvdid and ref_bvd_guess == cand_bvdid:
        ev.witnesses.append('BvD ID (from national register number)')

    # 2. Geographic match. City, country, and postal code are not independent
    # evidence, so they contribute at most one witness in order of specificity.
    ref_zip_found = ref_zip and ref_zip in norm_id(snapshot)
    city_matches = ref_city and cand_city and ref_city == cand_city
    country_matches = False
    if ref_iso2 and cand_country:
        country_matches = cand_country == ref_iso2 or COUNTRY_TO_ISO2.get(cand_country) == ref_iso2.upper()
    if ref_zip_found:
        ev.witnesses.append('Geographic match (postal code)')
    elif city_matches:
        ev.witnesses.append('Geographic match (city)')
    elif country_matches:
        ev.witnesses.append('Geographic match (country)')

    # 3. Name similarity - check both the casual/display name AND the legal
    # trade register name (often much closer to Orbis's legal candidate name).
    sim_company = name_similarity(ref_row.get('Company name'), cand.get('candidate_name'))
    sim_register = name_similarity(ref_row.get('Trade register name'), cand.get('candidate_name'))
    sim, sim_source = (sim_register, 'trade register name') if sim_register > sim_company else (sim_company, 'company name')
    if sim >= STRONG_NAME_SIM:
        ev.witnesses.append(f'Name similarity via {sim_source} ({sim:.2f})')
    elif sim >= WEAK_NAME_SIM:
        ev.weak_notes.append(f'weak name similarity via {sim_source} ({sim:.2f})')

    # 4. Website - direct field (report-page scrape) preferred, else substring
    if ref_website:
        domain = re.sub(r'^https?://(www\.)?', '', ref_website).split('/')[0].lower()
        cand_website = str(cand.get('website') or '').lower()
        if domain and cand_website and domain in cand_website:
            ev.witnesses.append('Website (direct field match)')
        elif domain and domain in snapshot.lower():
            ev.witnesses.append('Website domain (found in snapshot text)')

    # 5. Management / founder name - only usable once report-page scraping
    # supplies cand['management']
    cand_management = str(cand.get('management') or '')
    ref_founders = str(ref_row.get('Founders') or '')
    if cand_management and ref_founders:
        management_people = split_person_names(cand_management)
        for founder_core in split_person_names(ref_founders):
            for manager in management_people:
                if person_token_subset_match(founder_core, manager):
                    ev.witnesses.append(f'Management name match ({founder_core})')
                    break
            if has_witness(ev, 'management name match'):
                break

    owner_name = str(cand.get('owner_name') or '').strip()
    if owner_name:
        owner_company_sim = name_similarity(owner_name, ref_row.get('Company name'))
        owner_register_sim = name_similarity(owner_name, ref_row.get('Trade register name'))
        if (
            max(owner_company_sim, owner_register_sim) >= STRONG_NAME_SIM
            or normalized_containment_match(owner_name, ref_row.get('Company name'))
            or normalized_containment_match(owner_name, ref_row.get('Trade register name'))
        ):
            ev.witnesses.append(f'Global Ultimate Owner name match ({owner_name})')

    # 6. Industry / business purpose - opportunistic, only usable when we have
    # snapshot text with a NACE-style line (only present for escalated
    # candidates that had their popup opened). Weak signal on its own -
    # word-overlap between our Industries/Sub industries tags and Orbis's
    # activity description - so it only counts as a witness on a decent overlap.
    ref_industries = ' '.join(
        str(ref_row.get(c) or '') for c in ('Industries', 'Sub industries')
    ).replace(';', ' ')
    if ref_industries.strip() and snapshot:
        industry_words = {w for w in norm(ref_industries).split() if len(w) > 3}
        snapshot_words = set(norm(snapshot).split())
        overlap = industry_words & snapshot_words
        if len(overlap) >= 2:
            ev.witnesses.append(f'Industry/activity overlap ({", ".join(sorted(overlap))})')

    ref_launch_year = reference_year(ref_row.get('Launch year'))
    cand_incorporation_year = incorporation_year(cand)
    if (
        ref_launch_year is not None
        and cand_incorporation_year is not None
        and abs(ref_launch_year - cand_incorporation_year) <= 1
    ):
        ev.witnesses.append(f'Founding year match ({cand_incorporation_year})')

    add_project_text_witnesses(ev, ref_row, snapshot)

    return ev


def add_project_text_witnesses(ev: Evidence, ref_row: dict, text: str) -> None:
    haystack = norm(text)
    if not haystack:
        return
    useful_header_terms = [
        'legal', 'impressum', 'crunchbase', 'ownership', 'owner',
        'shareholder', 'parent', 'founder', 'manager', 'director', 'board',
        'management', 'incorporation', 'founded', 'industry', 'business',
        'purpose', 'activity', 'trade register', 'registered name',
    ]
    for column, value in ref_row.items():
        if value in (None, ''):
            continue
        column_norm = norm(column)
        if not any(term in column_norm for term in useful_header_terms):
            continue
        value_norm = norm(value)
        if len(value_norm) < 4:
            continue
        if value_norm in haystack:
            witness = f'{column} (found in Orbis text)'
            if witness not in ev.witnesses:
                ev.witnesses.append(witness)


def has_witness(ev: Evidence, *needles: str) -> bool:
    lowered = [w.casefold() for w in ev.witnesses]
    return any(any(needle.casefold() in witness for needle in needles) for witness in lowered)


def location_only(ev: Evidence) -> bool:
    return bool(ev.witnesses) and all(w.casefold().startswith('geographic match') for w in ev.witnesses)


def is_convincing(ev: Evidence) -> bool:
    if has_witness(ev, 'website', 'domain') and (
        ev.count >= 2 or has_witness(ev, 'name similarity', 'bvd', 'management', 'founder', 'manager', 'director', 'city', 'country')
    ):
        return True
    if ev.count < MIN_WITNESSES:
        return False
    if location_only(ev):
        return False
    if ev.count == 1 and has_witness(ev, 'name similarity'):
        return False
    if ev.count == 1 and has_witness(ev, 'country'):
        return False
    return True


def geographic_specificity(ev: Evidence) -> int:
    rank = 0
    for witness in ev.witnesses:
        lowered = witness.casefold()
        if 'geographic match (postal code)' in lowered:
            rank = max(rank, 3)
        elif 'geographic match (city)' in lowered:
            rank = max(rank, 2)
        elif 'geographic match (country)' in lowered:
            rank = max(rank, 1)
    return rank


def match_priority(pair: tuple[dict, Evidence]) -> tuple[int, int, int, int, int, int, int, int]:
    candidate, ev = pair
    return (
        1 if has_witness(ev, 'website', 'domain') else 0,
        ev.count,
        1 if has_witness(ev, 'bvd') else 0,
        1 if has_witness(ev, 'legal', 'impressum', 'trade register', 'registered name') else 0,
        1 if has_witness(ev, 'management', 'founder', 'manager', 'director', 'board', 'shareholder', 'owner', 'parent') else 0,
        1 if has_witness(ev, 'industry', 'activity', 'business', 'purpose', 'founding year') else 0,
        geographic_specificity(ev),
        1 if has_witness(ev, 'name similarity') else 0,
    )


def explain_decision(ev: Evidence) -> tuple[str, str]:
    primary = strongest_reason_witness(ev.witnesses, allow_geographic=False)
    secondary = strongest_reason_witness(ev.witnesses, allow_geographic=False, exclude=primary)
    if secondary is None:
        secondary = strongest_reason_witness(ev.witnesses, allow_geographic=True, exclude=primary)

    reason_1 = witness_to_reason(primary, primary=True)
    reason_2 = witness_to_reason(secondary, primary=False)
    return reason_1, reason_2


def strongest_reason_witness(witnesses: list[str], allow_geographic: bool, exclude: str | None = None) -> str | None:
    priority_groups = [
        ('website', 'domain'),
        ('bvd',),
        ('trade register', 'registered name'),
        ('management name match', 'founder', 'manager', 'director', 'board'),
        ('name similarity',),
        ('industry', 'activity', 'business purpose', 'founding year'),
        ('found in orbis text', 'legal', 'impressum', 'crunchbase', 'ownership', 'shareholder', 'parent', 'incorporation'),
    ]
    if allow_geographic:
        priority_groups.append(('geographic match',))

    for group in priority_groups:
        for witness in witnesses:
            if witness == exclude:
                continue
            lowered = witness.casefold()
            if not allow_geographic and lowered.startswith('geographic match'):
                continue
            if any(term in lowered for term in group):
                return witness
    for witness in witnesses:
        if witness == exclude:
            continue
        if not allow_geographic and witness.casefold().startswith('geographic match'):
            continue
        return witness
    return None


def witness_to_reason(witness: str | None, primary: bool) -> str:
    if not witness:
        return 'No additional distinct corroborating witness was available.'
    lowered = witness.casefold()
    prefix = '' if primary else 'Additionally corroborated by '
    safe_to_lowercase_reason_starts = {
        'Website',
        'Company',
        'Industry',
        'Trade',
        'Management',
        'Geographic',
    }
    if 'website' in lowered or 'domain' in lowered:
        text = 'Website matches the reference domain.'
    elif 'bvd' in lowered:
        text = 'BvD ID derived from national register number matches.'
    elif 'trade register' in lowered or 'registered name' in lowered or 'name similarity via trade register name' in lowered:
        text = 'Trade register name matches the Orbis candidate name.'
    elif 'management name match' in lowered or 'founder' in lowered or 'manager' in lowered or 'director' in lowered or 'board' in lowered:
        text = 'Management name matches a founder or manager listed in the reference data.'
    elif 'name similarity' in lowered:
        text = 'Company name similarity supports the candidate.'
    elif 'industry' in lowered or 'activity' in lowered or 'business purpose' in lowered:
        text = 'Industry/activity evidence is consistent with the reference company.'
    elif 'founding year' in lowered:
        text = 'Founding year is consistent with the Orbis incorporation year.'
    elif lowered.startswith('geographic match'):
        text = witness[0].upper() + witness[1:] + '.'
    else:
        text = f'{witness}.'
    if primary:
        return text
    first_word = text.split(maxsplit=1)[0].rstrip('.,:;')
    if first_word in safe_to_lowercase_reason_starts:
        text = text[0].lower() + text[1:]
    return prefix + text


def summarize_tied_candidate(candidate: dict) -> dict:
    return {
        'bvdid': candidate.get('bvdid', ''),
        'candidate_name': candidate.get('candidate_name', ''),
    }


def pick_convincing_candidate(ref_row: dict | None, candidates: list[dict]) -> tuple[dict | None, Evidence | None, dict | None, Evidence | None, str, list[dict] | None]:
    if ref_row is None or not candidates:
        return None, None, None, None, 'none', None

    scored = [(candidate, score_candidate(ref_row, candidate)) for candidate in candidates]
    convincing = [(candidate, ev) for candidate, ev in scored if is_convincing(ev)]
    if not convincing:
        best_candidate, best_ev = max(scored, key=match_priority)
        return None, None, best_candidate, best_ev, 'no_convincing', None

    convincing.sort(key=match_priority, reverse=True)
    best_candidate, best_ev = convincing[0]
    if len(convincing) > 1 and match_priority(convincing[0]) == match_priority(convincing[1]):
        tied = [summarize_tied_candidate(candidate) for candidate, _ev in convincing[:2]]
        return None, None, best_candidate, best_ev, 'ambiguous', tied

    return best_candidate, best_ev, best_candidate, best_ev, 'selected', None


def has_convincing_candidate(candidates: list[dict], ref_row: dict | None) -> bool:
    selected, _ev, _best_candidate, _best_ev, status, _tied = pick_convincing_candidate(ref_row, candidates)
    return bool(status == 'selected' and selected)


def dissolved_best_candidate_comment(candidate: dict | None) -> str:
    if not candidate:
        return ''
    status = str(candidate.get('company_status') or '')
    if 'dissolved' not in status.casefold():
        return ''
    name = candidate.get('candidate_name') or ''
    bvdid = candidate.get('bvdid') or ''
    return f' Best available candidate ({name}, {bvdid}) is marked Dissolved on Orbis - likely the correct entity but defunct; no further action expected.'


def select_best_match(ref_row: dict | None, candidates: list[dict], incomplete_candidate_count: int = 0) -> MatchDecision:
    if ref_row is None:
        return MatchDecision(None, None, False, 'No match', 'No reference row was available.', 'Manual review required.', 'No convincing match - manual review required')
    if not candidates:
        return MatchDecision(None, None, False, 'No match', 'No Orbis candidates were collected.', 'Manual review required.', 'No convincing match - manual review required')

    selected, ev, best_candidate, best_ev, status, tied_candidates = pick_convincing_candidate(ref_row, candidates)
    if status == 'no_convincing':
        comments = 'No convincing match - manual review required' + dissolved_best_candidate_comment(best_candidate)
        return MatchDecision(
            None,
            best_ev,
            False,
            'No match',
            f'Best candidate had {best_ev.count} supporting evidence signal(s), below the project threshold.',
            'Missing a strong independent signal such as website/domain, register ID, legal notice, management, ownership, or business-purpose evidence.',
            comments,
            best_candidate,
        )

    if status == 'ambiguous':
        return MatchDecision(
            None,
            best_ev,
            False,
            'Undecided (tie)',
            'Undecided - manual review: multiple candidates had equally strong supporting evidence.',
            'Manual review required to choose between the tied candidates.',
            'Undecided - manual review',
            best_candidate,
            tied_candidates,
        )

    reason_1, reason_2 = explain_decision(ev)
    if incomplete_candidate_count > 0:
        total_candidate_count = len(candidates) + incomplete_candidate_count
        comments = (
            f'Auto-matched after collecting {len(candidates)} of {total_candidate_count} candidates for this company; '
            f'{incomplete_candidate_count} candidate(s) could not be collected due to a scraping error.'
        )
    else:
        comments = 'Auto-matched after collecting all Orbis candidates for this company.'
    return MatchDecision(
        selected,
        ev,
        True,
        '; '.join(ev.witnesses),
        reason_1,
        reason_2,
        comments,
    )
