"""
Data Enrichment Module for Excel Forum Scraper.

Features:
- VBA/Macro code extraction
- Excel formula detection and parsing
- Duplicate detection using text similarity
- Answer quality scoring
- Attachment URL extraction
"""

import re
import zlib
import hashlib
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple, Set
from difflib import SequenceMatcher
from collections import Counter
import json


# Excel function list for formula detection
EXCEL_FUNCTIONS = {
    # Lookup & Reference
    'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'INDIRECT', 'OFFSET',
    'CHOOSE', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'LOOKUP', 'ADDRESS',
    'HYPERLINK', 'TRANSPOSE', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE',

    # Text
    'CONCATENATE', 'CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN',
    'FIND', 'SEARCH', 'REPLACE', 'SUBSTITUTE', 'TRIM', 'CLEAN', 'UPPER',
    'LOWER', 'PROPER', 'TEXT', 'VALUE', 'CHAR', 'CODE', 'REPT',

    # Math & Trig
    'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'COUNT', 'COUNTIF', 'COUNTIFS',
    'COUNTA', 'COUNTBLANK', 'AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'MAX',
    'MAXIFS', 'MIN', 'MINIFS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT',
    'MOD', 'ABS', 'SQRT', 'POWER', 'LOG', 'LN', 'EXP', 'RAND', 'RANDBETWEEN',

    # Logical
    'IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'TRUE', 'FALSE',
    'SWITCH', 'LET', 'LAMBDA',

    # Date & Time
    'DATE', 'DATEVALUE', 'DAY', 'MONTH', 'YEAR', 'TODAY', 'NOW', 'WEEKDAY',
    'WEEKNUM', 'EOMONTH', 'EDATE', 'DATEDIF', 'NETWORKDAYS', 'WORKDAY',
    'HOUR', 'MINUTE', 'SECOND', 'TIME', 'TIMEVALUE',

    # Statistical
    'STDEV', 'STDEVP', 'VAR', 'VARP', 'MEDIAN', 'MODE', 'PERCENTILE',
    'QUARTILE', 'LARGE', 'SMALL', 'RANK', 'FREQUENCY',

    # Financial
    'PMT', 'PV', 'FV', 'NPV', 'IRR', 'RATE', 'NPER',

    # Information
    'ISBLANK', 'ISERROR', 'ISNA', 'ISNUMBER', 'ISTEXT', 'ISLOGICAL',
    'ISREF', 'ISFORMULA', 'TYPE', 'N', 'NA', 'ERROR.TYPE', 'INFO', 'CELL',

    # Array/Dynamic
    'SEQUENCE', 'RANDARRAY', 'MAKEARRAY', 'MAP', 'REDUCE', 'SCAN',
    'BYROW', 'BYCOL', 'WRAPCOLS', 'WRAPROWS', 'TOCOL', 'TOROW',
    'EXPAND', 'DROP', 'TAKE', 'VSTACK', 'HSTACK', 'TEXTSPLIT', 'TEXTBEFORE', 'TEXTAFTER',
}


@dataclass
class ExtractedCode:
    """Represents extracted code from a post."""
    code_type: str  # 'vba', 'formula', 'sql', 'other'
    content: str
    language: str
    start_pos: int
    end_pos: int
    line_count: int
    functions_used: List[str] = field(default_factory=list)


@dataclass
class ExtractedFormula:
    """Represents an extracted Excel formula."""
    formula: str
    functions: List[str]
    cell_references: List[str]
    named_ranges: List[str]
    is_array_formula: bool = False
    complexity_score: int = 0


@dataclass
class QualityScore:
    """Quality score for a post/answer."""
    total_score: float
    has_code: bool
    has_formula: bool
    has_explanation: bool
    code_quality: float
    length_score: float
    author_reputation: float
    is_solution: bool
    upvotes: int = 0


class CodeExtractor:
    """Extract VBA code and formulas from forum posts."""

    # Patterns for VBA code blocks
    VBA_PATTERNS = [
        # XenForo code blocks
        r'\[CODE[^\]]*\](.*?)\[/CODE\]',
        # HTML pre/code tags
        r'<pre[^>]*><code[^>]*>(.*?)</code></pre>',
        r'<code[^>]*>(.*?)</code>',
        # Rich text code blocks
        r'```(?:vb|vba|basic)?\s*(.*?)```',
        # Indented code (4+ spaces at start of lines)
        r'(?:^|\n)((?:[ ]{4,}|\t).+(?:\n(?:[ ]{4,}|\t).+)*)',
    ]

    # VBA keywords for detection
    VBA_KEYWORDS = {
        'Sub', 'Function', 'End Sub', 'End Function', 'Dim', 'Set', 'Let',
        'If', 'Then', 'Else', 'ElseIf', 'End If', 'For', 'Next', 'Do', 'Loop',
        'While', 'Wend', 'Select Case', 'Case', 'End Select', 'With', 'End With',
        'Public', 'Private', 'Static', 'Const', 'ByVal', 'ByRef', 'Optional',
        'Variant', 'String', 'Integer', 'Long', 'Double', 'Boolean', 'Object',
        'Range', 'Cells', 'Worksheets', 'Workbooks', 'ActiveSheet', 'ActiveCell',
        'Application', 'ThisWorkbook', 'Selection', 'MsgBox', 'InputBox',
    }

    # Formula pattern
    FORMULA_PATTERN = r'=\s*([A-Z]+\s*\([^)]+\)(?:\s*[+\-*/&]\s*[A-Z]+\s*\([^)]+\))*|[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?(?:\s*[+\-*/&]\s*[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)*)'

    def __init__(self):
        self.vba_regex = [re.compile(p, re.DOTALL | re.IGNORECASE) for p in self.VBA_PATTERNS]
        self.formula_regex = re.compile(self.FORMULA_PATTERN, re.IGNORECASE)

    def extract_all(self, content: str) -> Dict[str, Any]:
        """Extract all code and formulas from content."""
        return {
            'vba_code': self.extract_vba(content),
            'formulas': self.extract_formulas(content),
            'sql_code': self.extract_sql(content),
            'attachments': self.extract_attachment_urls(content),
        }

    def extract_vba(self, content: str) -> List[ExtractedCode]:
        """Extract VBA/macro code from content."""
        extracted = []
        seen_code = set()

        for pattern in self.vba_regex:
            for match in pattern.finditer(content):
                code = match.group(1) if match.lastindex else match.group(0)
                code = self._clean_code(code)

                # Skip if too short or already seen
                code_hash = hashlib.md5(code.encode()).hexdigest()
                if len(code) < 20 or code_hash in seen_code:
                    continue

                # Check if it looks like VBA
                if self._is_vba_code(code):
                    seen_code.add(code_hash)
                    extracted.append(ExtractedCode(
                        code_type='vba',
                        content=code,
                        language='vba',
                        start_pos=match.start(),
                        end_pos=match.end(),
                        line_count=code.count('\n') + 1,
                        functions_used=self._extract_vba_functions(code)
                    ))

        return extracted

    def extract_formulas(self, content: str) -> List[ExtractedFormula]:
        """Extract Excel formulas from content."""
        formulas = []
        seen = set()

        # Pattern for formulas (starting with =)
        formula_patterns = [
            r'=\s*([A-Z]+\s*\([^=]*?\))',  # Function calls
            r'=\s*([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?(?:\s*[+\-*/&]\s*[^\s=]+)*)',  # Cell references
            r'=\s*((?:[A-Z]+\s*\([^)]*\)\s*[+\-*/&,]\s*)*[A-Z]+\s*\([^)]*\))',  # Nested functions
        ]

        # Also look for formulas in quotes or code blocks
        extended_patterns = [
            r'"(=[^"]+)"',
            r"'(=[^']+)'",
            r'\[CODE[^\]]*\](=[^\[]+)\[/CODE\]',
        ]

        all_patterns = formula_patterns + extended_patterns

        for pattern in all_patterns:
            for match in re.finditer(pattern, content, re.IGNORECASE):
                formula = match.group(1) if match.lastindex else match.group(0)
                formula = formula.strip()

                # Ensure it starts with =
                if not formula.startswith('='):
                    formula = '=' + formula

                # Skip duplicates and very short formulas
                if formula in seen or len(formula) < 3:
                    continue

                parsed = self._parse_formula(formula)
                if parsed and parsed.functions:
                    seen.add(formula)
                    formulas.append(parsed)

        return formulas

    def extract_sql(self, content: str) -> List[ExtractedCode]:
        """Extract SQL code from content."""
        extracted = []

        sql_patterns = [
            r'\[CODE[^\]]*sql[^\]]*\](.*?)\[/CODE\]',
            r'```sql\s*(.*?)```',
            r'(SELECT\s+.+?\s+FROM\s+.+?(?:WHERE|ORDER|GROUP|HAVING|;|\n\n))',
        ]

        sql_keywords = {'SELECT', 'INSERT', 'UPDATE', 'DELETE', 'FROM', 'WHERE',
                        'JOIN', 'LEFT', 'RIGHT', 'INNER', 'OUTER', 'GROUP BY', 'ORDER BY'}

        for pattern in sql_patterns:
            for match in re.finditer(pattern, content, re.DOTALL | re.IGNORECASE):
                code = match.group(1) if match.lastindex else match.group(0)
                code = self._clean_code(code)

                # Verify it's SQL
                upper_code = code.upper()
                sql_count = sum(1 for kw in sql_keywords if kw in upper_code)

                if sql_count >= 2 and len(code) > 20:
                    extracted.append(ExtractedCode(
                        code_type='sql',
                        content=code,
                        language='sql',
                        start_pos=match.start(),
                        end_pos=match.end(),
                        line_count=code.count('\n') + 1
                    ))

        return extracted

    def extract_attachment_urls(self, content: str) -> List[Dict[str, str]]:
        """Extract attachment/file URLs from content."""
        attachments = []

        patterns = [
            # Direct file links
            r'href=["\']([^"\']+\.(?:xlsx?|xlsm|xlsb|csv|txt))["\']',
            # Attachment references
            r'attachments?/([^"\'\s]+\.(?:xlsx?|xlsm|xlsb|csv))',
            # Download links
            r'download[^"\']*["\']([^"\']+)["\']',
        ]

        for pattern in patterns:
            for match in re.finditer(pattern, content, re.IGNORECASE):
                url = match.group(1)
                file_type = url.split('.')[-1].lower()
                attachments.append({
                    'url': url,
                    'file_type': file_type,
                    'is_excel': file_type in ('xls', 'xlsx', 'xlsm', 'xlsb')
                })

        return attachments

    def _clean_code(self, code: str) -> str:
        """Clean extracted code."""
        # Remove HTML entities
        code = re.sub(r'&lt;', '<', code)
        code = re.sub(r'&gt;', '>', code)
        code = re.sub(r'&amp;', '&', code)
        code = re.sub(r'&quot;', '"', code)
        code = re.sub(r'&#39;', "'", code)
        code = re.sub(r'&nbsp;', ' ', code)

        # Remove HTML tags
        code = re.sub(r'<[^>]+>', '', code)

        # Normalize whitespace
        lines = code.split('\n')
        lines = [line.rstrip() for line in lines]

        # Remove leading/trailing empty lines
        while lines and not lines[0].strip():
            lines.pop(0)
        while lines and not lines[-1].strip():
            lines.pop()

        return '\n'.join(lines)

    def _is_vba_code(self, code: str) -> bool:
        """Check if code looks like VBA."""
        upper_code = code.upper()

        # Count VBA keywords
        keyword_count = sum(1 for kw in self.VBA_KEYWORDS if kw.upper() in upper_code)

        # Check for definite VBA patterns
        definite_patterns = [
            r'\bSub\s+\w+\s*\(',
            r'\bFunction\s+\w+\s*\(',
            r'\bDim\s+\w+\s+As\s+',
            r'\bSet\s+\w+\s*=',
            r'\.Range\s*\(',
            r'\.Cells\s*\(',
        ]

        for pattern in definite_patterns:
            if re.search(pattern, code, re.IGNORECASE):
                return True

        return keyword_count >= 3

    def _extract_vba_functions(self, code: str) -> List[str]:
        """Extract function/sub names from VBA code."""
        functions = []

        patterns = [
            r'\b(?:Sub|Function)\s+(\w+)\s*\(',
            r'\bCall\s+(\w+)\s*\(',
        ]

        for pattern in patterns:
            for match in re.finditer(pattern, code, re.IGNORECASE):
                functions.append(match.group(1))

        return list(set(functions))

    def _parse_formula(self, formula: str) -> Optional[ExtractedFormula]:
        """Parse an Excel formula and extract components."""
        if not formula or not formula.startswith('='):
            return None

        # Extract functions used
        functions = []
        for func in EXCEL_FUNCTIONS:
            pattern = r'\b' + func + r'\s*\('
            if re.search(pattern, formula, re.IGNORECASE):
                functions.append(func)

        if not functions:
            return None

        # Extract cell references (A1, B2:C10, etc.)
        cell_refs = re.findall(r'\b([A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?)\b',
                               formula, re.IGNORECASE)

        # Extract named ranges (words that aren't functions or cell refs)
        words = re.findall(r'\b([A-Za-z_][A-Za-z0-9_]*)\b', formula)
        named_ranges = [w for w in words
                        if w.upper() not in EXCEL_FUNCTIONS
                        and not re.match(r'^[A-Z]{1,3}[0-9]+$', w, re.IGNORECASE)]

        # Check if array formula
        is_array = formula.startswith('={') or 'CTRL+SHIFT+ENTER' in formula.upper()

        # Calculate complexity
        complexity = (
            len(functions) * 2 +
            formula.count('(') +
            formula.count(',') +
            len(cell_refs) +
            (5 if is_array else 0)
        )

        return ExtractedFormula(
            formula=formula,
            functions=functions,
            cell_references=list(set(cell_refs)),
            named_ranges=list(set(named_ranges)),
            is_array_formula=is_array,
            complexity_score=complexity
        )


class DuplicateDetector:
    """Detect duplicate or similar discussions."""

    def __init__(self, similarity_threshold: float = 0.7):
        self.threshold = similarity_threshold
        self._cache: Dict[str, str] = {}

    def compute_similarity(self, text1: str, text2: str) -> float:
        """Compute similarity between two texts."""
        # Normalize texts
        text1 = self._normalize(text1)
        text2 = self._normalize(text2)

        # Use SequenceMatcher for similarity
        return SequenceMatcher(None, text1, text2).ratio()

    def compute_fingerprint(self, text: str) -> str:
        """Compute a fingerprint for quick duplicate detection."""
        text = self._normalize(text)

        # Create shingles (n-grams)
        shingles = set()
        words = text.split()
        for i in range(len(words) - 2):
            shingle = ' '.join(words[i:i+3])
            shingles.add(shingle)

        # Hash the shingles
        shingle_hashes = sorted([hash(s) % 10000 for s in shingles])

        # Take min-hash signature
        if shingle_hashes:
            signature = '_'.join(str(h) for h in shingle_hashes[:20])
        else:
            signature = hashlib.md5(text.encode()).hexdigest()

        return signature

    def find_duplicates(self, discussions: List[Dict[str, Any]],
                        field: str = 'title') -> List[Tuple[int, int, float]]:
        """Find duplicate discussions based on a field."""
        duplicates = []
        n = len(discussions)

        # Compute fingerprints
        fingerprints = {}
        for i, d in enumerate(discussions):
            text = d.get(field, '')
            fingerprints[i] = self.compute_fingerprint(text)

        # Compare fingerprints and compute similarity for potential matches
        for i in range(n):
            for j in range(i + 1, n):
                # Quick fingerprint comparison
                fp1, fp2 = fingerprints[i], fingerprints[j]
                fp_sim = self._fingerprint_similarity(fp1, fp2)

                if fp_sim > 0.3:  # Potential match
                    text1 = discussions[i].get(field, '')
                    text2 = discussions[j].get(field, '')
                    similarity = self.compute_similarity(text1, text2)

                    if similarity >= self.threshold:
                        duplicates.append((i, j, similarity))

        return duplicates

    def is_duplicate(self, new_text: str, existing_texts: List[str]) -> Tuple[bool, Optional[int], float]:
        """Check if new text is duplicate of any existing text."""
        new_text = self._normalize(new_text)

        best_match = -1
        best_score = 0.0

        for i, existing in enumerate(existing_texts):
            existing = self._normalize(existing)
            score = self.compute_similarity(new_text, existing)

            if score > best_score:
                best_score = score
                best_match = i

        is_dup = best_score >= self.threshold
        return is_dup, best_match if is_dup else None, best_score

    def _normalize(self, text: str) -> str:
        """Normalize text for comparison."""
        # Lowercase
        text = text.lower()

        # Remove punctuation and extra whitespace
        text = re.sub(r'[^\w\s]', ' ', text)
        text = re.sub(r'\s+', ' ', text)

        # Remove common stop words
        stop_words = {'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been',
                      'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will',
                      'would', 'could', 'should', 'may', 'might', 'must', 'shall',
                      'can', 'need', 'to', 'of', 'in', 'for', 'on', 'with', 'at',
                      'by', 'from', 'as', 'into', 'through', 'during', 'before',
                      'after', 'above', 'below', 'between', 'under', 'again',
                      'further', 'then', 'once', 'here', 'there', 'when', 'where',
                      'why', 'how', 'all', 'each', 'few', 'more', 'most', 'other',
                      'some', 'such', 'no', 'nor', 'not', 'only', 'own', 'same',
                      'so', 'than', 'too', 'very', 'just', 'i', 'me', 'my', 'we',
                      'our', 'you', 'your', 'he', 'him', 'his', 'she', 'her', 'it',
                      'its', 'they', 'them', 'their', 'this', 'that', 'these', 'those'}

        words = [w for w in text.split() if w not in stop_words and len(w) > 2]
        return ' '.join(words)

    def _fingerprint_similarity(self, fp1: str, fp2: str) -> float:
        """Quick similarity check between fingerprints."""
        parts1 = set(fp1.split('_'))
        parts2 = set(fp2.split('_'))

        if not parts1 or not parts2:
            return 0.0

        intersection = len(parts1 & parts2)
        union = len(parts1 | parts2)

        return intersection / union if union > 0 else 0.0


class QualityScorer:
    """Score the quality of posts and answers."""

    def __init__(self):
        self.code_extractor = CodeExtractor()

    def score_post(self, post: Dict[str, Any], author_stats: Dict[str, Any] = None) -> QualityScore:
        """Calculate quality score for a post."""
        content = post.get('content', '')

        # Extract code and formulas
        extracted = self.code_extractor.extract_all(content)
        has_code = bool(extracted['vba_code'])
        has_formula = bool(extracted['formulas'])

        # Check for explanation
        explanation_indicators = [
            len(content) > 200,
            bool(re.search(r'\b(because|since|therefore|this means|note that|remember)\b',
                          content, re.IGNORECASE)),
            content.count('.') > 2,  # Multiple sentences
        ]
        has_explanation = sum(explanation_indicators) >= 2

        # Code quality score (0-1)
        code_quality = 0.0
        if has_code:
            vba_code = extracted['vba_code']
            total_lines = sum(c.line_count for c in vba_code)
            has_comments = any("'" in c.content or 'rem ' in c.content.lower() for c in vba_code)
            has_error_handling = any('on error' in c.content.lower() for c in vba_code)

            code_quality = min(1.0, (
                0.3 * min(1.0, total_lines / 20) +
                0.3 * (1 if has_comments else 0) +
                0.2 * (1 if has_error_handling else 0) +
                0.2 * (1 if len(vba_code) > 0 else 0)
            ))
        elif has_formula:
            formulas = extracted['formulas']
            avg_complexity = sum(f.complexity_score for f in formulas) / len(formulas)
            code_quality = min(1.0, avg_complexity / 20)

        # Length score (0-1)
        word_count = len(content.split())
        length_score = min(1.0, word_count / 150)  # Optimal around 150 words

        # Author reputation score (0-1)
        author_reputation = 0.5  # Default
        if author_stats:
            post_count = author_stats.get('post_count', 0)
            reputation = author_stats.get('reputation', 0)

            author_reputation = min(1.0, (
                0.5 * min(1.0, post_count / 1000) +
                0.5 * min(1.0, reputation / 500)
            ))

        # Is solution marker
        is_solution = post.get('is_solution', False)

        # Calculate total score
        total_score = (
            0.15 * (1 if has_code else 0) +
            0.10 * (1 if has_formula else 0) +
            0.15 * (1 if has_explanation else 0) +
            0.15 * code_quality +
            0.10 * length_score +
            0.15 * author_reputation +
            0.20 * (1 if is_solution else 0)
        )

        return QualityScore(
            total_score=round(total_score, 3),
            has_code=has_code,
            has_formula=has_formula,
            has_explanation=has_explanation,
            code_quality=round(code_quality, 3),
            length_score=round(length_score, 3),
            author_reputation=round(author_reputation, 3),
            is_solution=is_solution,
            upvotes=post.get('upvotes', 0)
        )

    def rank_answers(self, posts: List[Dict[str, Any]],
                     author_stats: Dict[str, Dict[str, Any]] = None) -> List[Tuple[int, QualityScore]]:
        """Rank answers by quality score."""
        if author_stats is None:
            author_stats = {}

        scored = []
        for i, post in enumerate(posts):
            author = post.get('author', '')
            stats = author_stats.get(author, {})
            score = self.score_post(post, stats)
            scored.append((i, score))

        # Sort by total score descending
        scored.sort(key=lambda x: x[1].total_score, reverse=True)
        return scored

    def get_best_answer(self, posts: List[Dict[str, Any]],
                        author_stats: Dict[str, Dict[str, Any]] = None) -> Optional[int]:
        """Get the index of the best answer."""
        if not posts:
            return None

        ranked = self.rank_answers(posts, author_stats)

        # Return the best one if score is above threshold
        if ranked and ranked[0][1].total_score > 0.3:
            return ranked[0][0]

        return None


class ContentCompressor:
    """Compress and decompress content for storage."""

    @staticmethod
    def compress(content: str) -> bytes:
        """Compress string content."""
        return zlib.compress(content.encode('utf-8'), level=6)

    @staticmethod
    def decompress(data: bytes) -> str:
        """Decompress content."""
        return zlib.decompress(data).decode('utf-8')

    @staticmethod
    def compression_ratio(original: str, compressed: bytes) -> float:
        """Calculate compression ratio."""
        original_size = len(original.encode('utf-8'))
        compressed_size = len(compressed)

        if original_size == 0:
            return 1.0

        return compressed_size / original_size


class EnrichmentPipeline:
    """Pipeline for enriching forum discussions."""

    def __init__(self):
        self.code_extractor = CodeExtractor()
        self.duplicate_detector = DuplicateDetector()
        self.quality_scorer = QualityScorer()
        self.compressor = ContentCompressor()

    def enrich_discussion(self, discussion: Dict[str, Any],
                          author_stats: Dict[str, Dict[str, Any]] = None) -> Dict[str, Any]:
        """Enrich a single discussion with extracted data."""
        enriched = dict(discussion)

        # Extract code and formulas from all posts
        all_vba = []
        all_formulas = []
        all_sql = []
        all_attachments = []

        posts = discussion.get('posts', [])
        scored_posts = []

        for post in posts:
            content = post.get('content', '')
            extracted = self.code_extractor.extract_all(content)

            all_vba.extend(extracted['vba_code'])
            all_formulas.extend(extracted['formulas'])
            all_sql.extend(extracted['sql_code'])
            all_attachments.extend(extracted['attachments'])

            # Score the post
            author = post.get('author', '')
            stats = (author_stats or {}).get(author, {})
            score = self.quality_scorer.score_post(post, stats)
            scored_posts.append({
                'post_number': post.get('post_number', 0),
                'quality_score': score.total_score,
                'has_code': score.has_code,
                'has_formula': score.has_formula,
            })

        # Add enrichment data
        enriched['enrichment'] = {
            'vba_code_count': len(all_vba),
            'formula_count': len(all_formulas),
            'sql_code_count': len(all_sql),
            'attachment_count': len(all_attachments),
            'vba_code': [{'content': c.content, 'functions': c.functions_used} for c in all_vba],
            'formulas': [{'formula': f.formula, 'functions': f.functions,
                          'complexity': f.complexity_score} for f in all_formulas],
            'attachments': all_attachments,
            'post_scores': scored_posts,
            'best_answer_index': self.quality_scorer.get_best_answer(posts, author_stats),
            'functions_mentioned': list(set(
                func for f in all_formulas for func in f.functions
            )),
        }

        return enriched

    def enrich_batch(self, discussions: List[Dict[str, Any]],
                     author_stats: Dict[str, Dict[str, Any]] = None,
                     detect_duplicates: bool = True) -> List[Dict[str, Any]]:
        """Enrich a batch of discussions."""
        enriched = []

        for discussion in discussions:
            enriched_disc = self.enrich_discussion(discussion, author_stats)
            enriched.append(enriched_disc)

        # Detect duplicates
        if detect_duplicates and len(enriched) > 1:
            duplicates = self.duplicate_detector.find_duplicates(enriched, 'title')

            # Mark duplicates
            for i, j, similarity in duplicates:
                if 'duplicates' not in enriched[i]:
                    enriched[i]['duplicates'] = []
                if 'duplicates' not in enriched[j]:
                    enriched[j]['duplicates'] = []

                enriched[i]['duplicates'].append({
                    'discussion_index': j,
                    'similarity': similarity
                })
                enriched[j]['duplicates'].append({
                    'discussion_index': i,
                    'similarity': similarity
                })

        return enriched
