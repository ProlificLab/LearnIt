#!/usr/bin/env python3
"""
Community Scrapers for Excel and Power BI content.

Supported sources:
- Stack Overflow (excel, vba, powerbi tags)
- Reddit (r/excel, r/vba, r/PowerBI)
- Microsoft Tech Community
- Power BI Community (community.powerbi.com)

These sources help understand what people want to learn about Excel/Power BI.
"""

import re
import json
import time
import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import List, Dict, Any, Optional
from urllib.parse import urljoin, urlencode, quote

import requests
from bs4 import BeautifulSoup
from ratelimit import limits, sleep_and_retry

from database import ForumDatabase

logger = logging.getLogger(__name__)


@dataclass
class Question:
    """Represents a Q&A question from any source."""
    title: str
    url: str
    body: str
    author: str
    created_at: str
    tags: List[str] = field(default_factory=list)
    score: int = 0
    views: int = 0
    answers_count: int = 0
    is_answered: bool = False
    accepted_answer: Optional[str] = None
    source: str = ""
    source_id: str = ""
    category: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class Answer:
    """Represents an answer to a question."""
    body: str
    author: str
    created_at: str
    score: int = 0
    is_accepted: bool = False
    source_id: str = ""


class BaseCommunityScraper(ABC):
    """Base class for community scrapers."""

    def __init__(self, db: ForumDatabase = None, delay: float = 1.0):
        self.db = db
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json, text/html',
        })
        self.questions: List[Question] = []

    @property
    @abstractmethod
    def source_name(self) -> str:
        pass

    @abstractmethod
    def scrape(self, **kwargs) -> List[Question]:
        pass

    @sleep_and_retry
    @limits(calls=30, period=60)
    def fetch(self, url: str, params: Dict = None) -> Optional[requests.Response]:
        """Fetch URL with rate limiting."""
        try:
            time.sleep(self.delay)
            response = self.session.get(url, params=params, timeout=30)
            response.raise_for_status()
            return response
        except requests.RequestException as e:
            logger.error(f"Failed to fetch {url}: {e}")
            return None

    def save_to_db(self, question: Question):
        """Save question to database as a discussion."""
        if not self.db:
            return

        discussion = {
            'url': question.url,
            'title': question.title,
            'author': question.author,
            'date_created': question.created_at,
            'category': question.category or ','.join(question.tags[:3]),
            'tags': question.tags,
            'views': question.views,
            'replies': question.answers_count,
            'forum_source': question.source,
            'is_solved': question.is_answered,
            'posts': [{
                'author': question.author,
                'content': question.body,
                'date': question.created_at,
                'post_number': 1,
                'is_solution': False
            }]
        }

        if question.accepted_answer:
            discussion['posts'].append({
                'author': 'Accepted Answer',
                'content': question.accepted_answer,
                'date': question.created_at,
                'post_number': 2,
                'is_solution': True
            })

        self.db.upsert_discussion(discussion)


# ============== Stack Overflow ==============

class StackOverflowScraper(BaseCommunityScraper):
    """
    Scraper for Stack Overflow using the public API.

    API Docs: https://api.stackexchange.com/docs
    No API key required for basic usage (limited to 300 requests/day)
    """

    API_BASE = "https://api.stackexchange.com/2.3"

    EXCEL_TAGS = ['excel', 'vba', 'excel-formula', 'excel-vba', 'ms-excel',
                  'excel-2016', 'excel-2019', 'excel-365', 'xlookup', 'pivot-table']

    POWERBI_TAGS = ['powerbi', 'power-bi', 'dax', 'power-query', 'm',
                    'powerbi-desktop', 'power-bi-service', 'power-bi-embedded']

    @property
    def source_name(self) -> str:
        return "StackOverflow"

    def scrape(self, tags: List[str] = None, pages: int = 5,
               sort: str = 'votes', include_body: bool = True) -> List[Question]:
        """
        Scrape questions from Stack Overflow.

        Args:
            tags: Tags to search for (default: Excel + Power BI tags)
            pages: Number of pages to fetch (100 questions per page)
            sort: Sort order - 'votes', 'activity', 'creation', 'hot'
            include_body: Whether to fetch full question body
        """
        if tags is None:
            tags = self.EXCEL_TAGS + self.POWERBI_TAGS

        logger.info(f"Scraping Stack Overflow for tags: {tags}")

        for tag in tags:
            logger.info(f"Fetching questions for tag: {tag}")

            for page in range(1, pages + 1):
                questions = self._fetch_questions(tag, page, sort, include_body)

                if not questions:
                    break

                for q in questions:
                    self.questions.append(q)
                    self.save_to_db(q)

                logger.info(f"  Page {page}: {len(questions)} questions")

        logger.info(f"Total scraped: {len(self.questions)} questions")
        return self.questions

    def _fetch_questions(self, tag: str, page: int, sort: str,
                         include_body: bool) -> List[Question]:
        """Fetch questions for a tag."""
        params = {
            'order': 'desc',
            'sort': sort,
            'tagged': tag,
            'site': 'stackoverflow',
            'page': page,
            'pagesize': 100,
            'filter': 'withbody' if include_body else 'default'
        }

        response = self.fetch(f"{self.API_BASE}/questions", params)
        if not response:
            return []

        try:
            data = response.json()
        except json.JSONDecodeError:
            return []

        questions = []
        for item in data.get('items', []):
            q = Question(
                title=item.get('title', ''),
                url=item.get('link', ''),
                body=item.get('body', ''),
                author=item.get('owner', {}).get('display_name', 'Unknown'),
                created_at=datetime.fromtimestamp(
                    item.get('creation_date', 0)
                ).isoformat(),
                tags=item.get('tags', []),
                score=item.get('score', 0),
                views=item.get('view_count', 0),
                answers_count=item.get('answer_count', 0),
                is_answered=item.get('is_answered', False),
                source=self.source_name,
                source_id=str(item.get('question_id', '')),
                category=tag
            )
            questions.append(q)

        # Check quota
        remaining = data.get('quota_remaining', 0)
        if remaining < 10:
            logger.warning(f"Stack Overflow API quota low: {remaining} remaining")

        return questions

    def get_trending_tags(self, count: int = 50) -> List[Dict[str, Any]]:
        """Get trending Excel/Power BI related tags."""
        all_trending = []

        for base_tag in ['excel', 'powerbi', 'vba', 'dax']:
            params = {
                'order': 'desc',
                'sort': 'popular',
                'inname': base_tag,
                'site': 'stackoverflow',
                'pagesize': min(count, 100)
            }

            response = self.fetch(f"{self.API_BASE}/tags", params)
            if response:
                try:
                    data = response.json()
                    for item in data.get('items', []):
                        all_trending.append({
                            'name': item.get('name'),
                            'count': item.get('count', 0)
                        })
                except json.JSONDecodeError:
                    pass

        # Sort by count and deduplicate
        seen = set()
        unique = []
        for t in sorted(all_trending, key=lambda x: x['count'], reverse=True):
            if t['name'] not in seen:
                seen.add(t['name'])
                unique.append(t)

        return unique[:count]


# ============== Reddit ==============

class RedditScraper(BaseCommunityScraper):
    """
    Scraper for Reddit using the public JSON API.

    No authentication required for public subreddits.
    Append .json to any Reddit URL to get JSON data.
    """

    SUBREDDITS = {
        'excel': {'name': 'excel', 'category': 'Excel'},
        'vba': {'name': 'vba', 'category': 'VBA'},
        'powerbi': {'name': 'PowerBI', 'category': 'Power BI'},
        'msexcel': {'name': 'msexcel', 'category': 'Excel'},
        'excelevator': {'name': 'excelevator', 'category': 'Excel Advanced'},
    }

    @property
    def source_name(self) -> str:
        return "Reddit"

    def scrape(self, subreddits: List[str] = None, posts_per_sub: int = 100,
               sort: str = 'top', time_filter: str = 'month') -> List[Question]:
        """
        Scrape posts from Excel/Power BI subreddits.

        Args:
            subreddits: List of subreddit names (default: all Excel/PBI subs)
            posts_per_sub: Number of posts to fetch per subreddit
            sort: Sort order - 'hot', 'new', 'top', 'rising'
            time_filter: For 'top' sort - 'hour', 'day', 'week', 'month', 'year', 'all'
        """
        if subreddits is None:
            subreddits = list(self.SUBREDDITS.keys())

        for sub_key in subreddits:
            sub_info = self.SUBREDDITS.get(sub_key, {'name': sub_key, 'category': sub_key})
            sub_name = sub_info['name']
            category = sub_info['category']

            logger.info(f"Scraping r/{sub_name}")

            posts = self._fetch_subreddit(sub_name, posts_per_sub, sort, time_filter)

            for post in posts:
                post.category = category
                self.questions.append(post)
                self.save_to_db(post)

            logger.info(f"  Fetched {len(posts)} posts from r/{sub_name}")

        logger.info(f"Total scraped: {len(self.questions)} posts")
        return self.questions

    def _fetch_subreddit(self, subreddit: str, limit: int, sort: str,
                         time_filter: str) -> List[Question]:
        """Fetch posts from a subreddit."""
        posts = []
        after = None
        fetched = 0

        while fetched < limit:
            batch_size = min(100, limit - fetched)
            url = f"https://www.reddit.com/r/{subreddit}/{sort}.json"

            params = {'limit': batch_size, 't': time_filter}
            if after:
                params['after'] = after

            response = self.fetch(url, params)
            if not response:
                break

            try:
                data = response.json()
            except json.JSONDecodeError:
                break

            items = data.get('data', {}).get('children', [])
            if not items:
                break

            for item in items:
                post_data = item.get('data', {})

                # Skip removed/deleted posts
                if post_data.get('removed_by_category') or post_data.get('selftext') == '[removed]':
                    continue

                q = Question(
                    title=post_data.get('title', ''),
                    url=f"https://reddit.com{post_data.get('permalink', '')}",
                    body=post_data.get('selftext', ''),
                    author=post_data.get('author', 'Unknown'),
                    created_at=datetime.fromtimestamp(
                        post_data.get('created_utc', 0)
                    ).isoformat(),
                    tags=[post_data.get('link_flair_text', '')] if post_data.get('link_flair_text') else [],
                    score=post_data.get('score', 0),
                    views=0,  # Reddit doesn't expose view counts
                    answers_count=post_data.get('num_comments', 0),
                    is_answered=bool(post_data.get('link_flair_text', '').lower() in
                                    ['solved', 'answered', 'solution verified']),
                    source=f"Reddit/r/{subreddit}",
                    source_id=post_data.get('id', ''),
                )
                posts.append(q)

            after = data.get('data', {}).get('after')
            fetched += len(items)

            if not after:
                break

        return posts

    def get_trending_topics(self, subreddit: str = 'excel', limit: int = 20) -> List[Dict[str, Any]]:
        """Get trending/hot topics from a subreddit."""
        url = f"https://www.reddit.com/r/{subreddit}/hot.json"
        response = self.fetch(url, {'limit': limit})

        if not response:
            return []

        try:
            data = response.json()
            topics = []

            for item in data.get('data', {}).get('children', []):
                post = item.get('data', {})
                topics.append({
                    'title': post.get('title'),
                    'score': post.get('score'),
                    'comments': post.get('num_comments'),
                    'url': f"https://reddit.com{post.get('permalink')}",
                    'flair': post.get('link_flair_text')
                })

            return topics
        except json.JSONDecodeError:
            return []


# ============== Microsoft Tech Community ==============

class MSTechCommunityScraper(BaseCommunityScraper):
    """
    Scraper for Microsoft Tech Community.

    Covers Excel, Power BI, and other Microsoft products.
    """

    BASE_URL = "https://techcommunity.microsoft.com"

    BOARDS = {
        'excel': '/t5/excel/bd-p/ExcelGeneral',
        'powerbi': '/t5/power-bi/bd-p/powerbi',
        'powerquery': '/t5/power-query/bd-p/powerquery',
        'office365': '/t5/microsoft-365/bd-p/microsoft365',
    }

    @property
    def source_name(self) -> str:
        return "MSTechCommunity"

    def scrape(self, boards: List[str] = None, pages: int = 5) -> List[Question]:
        """
        Scrape discussions from Microsoft Tech Community.

        Args:
            boards: Board names to scrape (default: all)
            pages: Pages to scrape per board
        """
        if boards is None:
            boards = list(self.BOARDS.keys())

        for board in boards:
            board_path = self.BOARDS.get(board)
            if not board_path:
                continue

            logger.info(f"Scraping MS Tech Community: {board}")

            for page in range(1, pages + 1):
                url = f"{self.BASE_URL}{board_path}/page/{page}"
                response = self.fetch(url)

                if not response:
                    break

                questions = self._parse_board_page(response.text, board)
                for q in questions:
                    self.questions.append(q)
                    self.save_to_db(q)

                logger.info(f"  Page {page}: {len(questions)} discussions")

                if len(questions) < 10:
                    break

        logger.info(f"Total scraped: {len(self.questions)} discussions")
        return self.questions

    def _parse_board_page(self, html: str, category: str) -> List[Question]:
        """Parse a board page for discussions."""
        soup = BeautifulSoup(html, 'lxml')
        questions = []

        # Find discussion items
        items = soup.select('.lia-component-messages-column-message-body, .message-subject')

        for item in items:
            try:
                title_elem = item.select_one('a.page-link, .message-subject a')
                if not title_elem:
                    continue

                title = title_elem.get_text(strip=True)
                url = urljoin(self.BASE_URL, title_elem.get('href', ''))

                author_elem = item.select_one('.lia-user-name-link, .author-name')
                author = author_elem.get_text(strip=True) if author_elem else 'Unknown'

                date_elem = item.select_one('.local-date, time')
                created_at = date_elem.get('datetime', '') if date_elem else ''

                # Get reply count
                reply_elem = item.select_one('.lia-message-stats-count, .reply-count')
                replies = 0
                if reply_elem:
                    try:
                        replies = int(re.sub(r'\D', '', reply_elem.get_text()))
                    except ValueError:
                        pass

                # Check if solved
                is_solved = bool(item.select_one('.lia-component-accepted-solution, .solved-icon'))

                q = Question(
                    title=title,
                    url=url,
                    body='',  # Would need to fetch individual pages
                    author=author,
                    created_at=created_at,
                    tags=[category],
                    answers_count=replies,
                    is_answered=is_solved,
                    source=self.source_name,
                    category=category
                )
                questions.append(q)

            except Exception as e:
                logger.debug(f"Error parsing item: {e}")
                continue

        return questions


# ============== Power BI Community ==============

class PowerBICommunityScraper(BaseCommunityScraper):
    """
    Scraper for Power BI Community (community.powerbi.com).

    Official Power BI forum with categories for Desktop, Service, DAX, etc.
    """

    BASE_URL = "https://community.powerbi.com"

    CATEGORIES = {
        'desktop': '/t5/Desktop/bd-p/PBI_Comm_Desktop',
        'service': '/t5/Service/bd-p/pbi_comm_service',
        'dax': '/t5/DAX-Commands-and-Tips/bd-p/DAXCommands',
        'dataflows': '/t5/Dataflows/bd-p/Dataflows',
        'paginated': '/t5/Report-Server/bd-p/pbi_comm_report_server',
        'mobile': '/t5/Mobile-Apps/bd-p/PBI_Comm_MobileApps',
        'developer': '/t5/Developer/bd-p/Community_Comm_Development',
    }

    @property
    def source_name(self) -> str:
        return "PowerBICommunity"

    def scrape(self, categories: List[str] = None, pages: int = 5) -> List[Question]:
        """
        Scrape discussions from Power BI Community.

        Args:
            categories: Categories to scrape (default: all)
            pages: Pages to scrape per category
        """
        if categories is None:
            categories = list(self.CATEGORIES.keys())

        for cat in categories:
            cat_path = self.CATEGORIES.get(cat)
            if not cat_path:
                continue

            logger.info(f"Scraping Power BI Community: {cat}")

            for page in range(1, pages + 1):
                url = f"{self.BASE_URL}{cat_path}/page/{page}"
                response = self.fetch(url)

                if not response:
                    break

                questions = self._parse_category_page(response.text, cat)
                for q in questions:
                    self.questions.append(q)
                    self.save_to_db(q)

                logger.info(f"  Page {page}: {len(questions)} discussions")

                if len(questions) < 10:
                    break

        logger.info(f"Total scraped: {len(self.questions)} discussions")
        return self.questions

    def _parse_category_page(self, html: str, category: str) -> List[Question]:
        """Parse a category page for discussions."""
        soup = BeautifulSoup(html, 'lxml')
        questions = []

        items = soup.select('.lia-quilt-row, .MessageSubject')

        for item in items:
            try:
                title_elem = item.select_one('.message-subject a, .page-link')
                if not title_elem:
                    continue

                title = title_elem.get_text(strip=True)
                url = urljoin(self.BASE_URL, title_elem.get('href', ''))

                author_elem = item.select_one('.lia-user-name-link')
                author = author_elem.get_text(strip=True) if author_elem else 'Unknown'

                date_elem = item.select_one('.local-date, .DateTime')
                created_at = date_elem.get('datetime', '') if date_elem else ''

                # Views
                views_elem = item.select_one('.lia-message-stats-view-count')
                views = 0
                if views_elem:
                    try:
                        views = int(re.sub(r'\D', '', views_elem.get_text()))
                    except ValueError:
                        pass

                # Replies
                reply_elem = item.select_one('.lia-message-stats-reply-count')
                replies = 0
                if reply_elem:
                    try:
                        replies = int(re.sub(r'\D', '', reply_elem.get_text()))
                    except ValueError:
                        pass

                # Check if solved
                is_solved = bool(item.select_one('.lia-component-accepted-solution'))

                # Tags/labels
                tags = [category]
                label_elems = item.select('.lia-label, .tag')
                for label in label_elems:
                    tags.append(label.get_text(strip=True))

                q = Question(
                    title=title,
                    url=url,
                    body='',
                    author=author,
                    created_at=created_at,
                    tags=tags,
                    views=views,
                    answers_count=replies,
                    is_answered=is_solved,
                    source=self.source_name,
                    category=category
                )
                questions.append(q)

            except Exception as e:
                logger.debug(f"Error parsing item: {e}")
                continue

        return questions

    def get_ideas(self, pages: int = 3) -> List[Dict[str, Any]]:
        """Get feature ideas/requests from the community."""
        ideas = []
        ideas_url = f"{self.BASE_URL}/t5/Ideas/idb-p/Ideas"

        for page in range(1, pages + 1):
            response = self.fetch(f"{ideas_url}/page/{page}")
            if not response:
                break

            soup = BeautifulSoup(response.text, 'lxml')
            items = soup.select('.lia-idea-row, .idea-item')

            for item in items:
                try:
                    title_elem = item.select_one('.idea-title a, .page-link')
                    if not title_elem:
                        continue

                    votes_elem = item.select_one('.kudos-count, .idea-votes')
                    votes = 0
                    if votes_elem:
                        try:
                            votes = int(re.sub(r'\D', '', votes_elem.get_text()))
                        except ValueError:
                            pass

                    status_elem = item.select_one('.idea-status, .lia-idea-status')
                    status = status_elem.get_text(strip=True) if status_elem else 'New'

                    ideas.append({
                        'title': title_elem.get_text(strip=True),
                        'url': urljoin(self.BASE_URL, title_elem.get('href', '')),
                        'votes': votes,
                        'status': status
                    })
                except Exception:
                    continue

        return ideas


# ============== Topic Analyzer ==============

class TopicAnalyzer:
    """Analyze scraped questions to find trending topics and common issues."""

    def __init__(self, questions: List[Question]):
        self.questions = questions

    def get_common_words(self, top_n: int = 50) -> Dict[str, int]:
        """Extract most common words from question titles."""
        # Stop words to exclude
        stop_words = {
            'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
            'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could',
            'should', 'may', 'might', 'must', 'shall', 'can', 'need', 'to', 'of',
            'in', 'for', 'on', 'with', 'at', 'by', 'from', 'as', 'into', 'through',
            'and', 'or', 'but', 'not', 'so', 'than', 'too', 'very', 'just', 'i',
            'me', 'my', 'we', 'our', 'you', 'your', 'he', 'him', 'his', 'she', 'her',
            'it', 'its', 'they', 'them', 'their', 'this', 'that', 'these', 'those',
            'how', 'what', 'when', 'where', 'why', 'which', 'who', 'whom', 'whose',
            'if', 'then', 'else', 'all', 'each', 'every', 'both', 'few', 'more',
            'most', 'other', 'some', 'such', 'no', 'nor', 'only', 'own', 'same',
            'using', 'use', 'get', 'getting', 'want', 'need', 'help', 'please',
            'excel', 'vba', 'powerbi', 'power', 'bi', 'question', 'issue', 'problem',
        }

        word_counts = {}

        for q in self.questions:
            # Tokenize title
            words = re.findall(r'\b[a-zA-Z]{3,}\b', q.title.lower())

            for word in words:
                if word not in stop_words:
                    word_counts[word] = word_counts.get(word, 0) + 1

        return dict(sorted(word_counts.items(), key=lambda x: x[1], reverse=True)[:top_n])

    def get_tag_distribution(self) -> Dict[str, int]:
        """Get distribution of tags across questions."""
        tag_counts = {}

        for q in self.questions:
            for tag in q.tags:
                if tag:
                    tag_counts[tag.lower()] = tag_counts.get(tag.lower(), 0) + 1

        return dict(sorted(tag_counts.items(), key=lambda x: x[1], reverse=True))

    def get_unanswered_topics(self, top_n: int = 20) -> List[Dict[str, Any]]:
        """Get common topics in unanswered questions."""
        unanswered = [q for q in self.questions if not q.is_answered]

        word_counts = {}
        for q in unanswered:
            words = re.findall(r'\b[a-zA-Z]{4,}\b', q.title.lower())
            for word in words:
                word_counts[word] = word_counts.get(word, 0) + 1

        return [
            {'topic': word, 'count': count}
            for word, count in sorted(word_counts.items(), key=lambda x: x[1], reverse=True)[:top_n]
        ]

    def get_source_summary(self) -> Dict[str, Dict[str, Any]]:
        """Get summary statistics by source."""
        sources = {}

        for q in self.questions:
            source = q.source
            if source not in sources:
                sources[source] = {
                    'count': 0,
                    'answered': 0,
                    'total_score': 0,
                    'total_views': 0
                }

            sources[source]['count'] += 1
            if q.is_answered:
                sources[source]['answered'] += 1
            sources[source]['total_score'] += q.score
            sources[source]['total_views'] += q.views

        # Calculate percentages
        for source in sources:
            count = sources[source]['count']
            sources[source]['answer_rate'] = round(
                sources[source]['answered'] / count * 100, 1
            ) if count > 0 else 0
            sources[source]['avg_score'] = round(
                sources[source]['total_score'] / count, 1
            ) if count > 0 else 0

        return sources

    def export_insights(self, filepath: str):
        """Export analysis to JSON file."""
        insights = {
            'generated_at': datetime.now().isoformat(),
            'total_questions': len(self.questions),
            'sources': self.get_source_summary(),
            'common_words': self.get_common_words(30),
            'tag_distribution': self.get_tag_distribution(),
            'unanswered_topics': self.get_unanswered_topics(20)
        }

        with open(filepath, 'w') as f:
            json.dump(insights, f, indent=2)

        return insights


# ============== CLI ==============

def main():
    import argparse

    parser = argparse.ArgumentParser(description='Scrape Excel/Power BI communities')
    parser.add_argument('source', choices=['stackoverflow', 'reddit', 'mstech', 'powerbi', 'all'],
                        help='Source to scrape')
    parser.add_argument('--pages', type=int, default=5, help='Pages to scrape')
    parser.add_argument('--database', default='./data/forum_data.db', help='Database path')
    parser.add_argument('--output', default='./data', help='Output directory')
    parser.add_argument('--analyze', action='store_true', help='Run topic analysis')

    args = parser.parse_args()

    db = ForumDatabase(args.database)
    all_questions = []

    sources = {
        'stackoverflow': StackOverflowScraper,
        'reddit': RedditScraper,
        'mstech': MSTechCommunityScraper,
        'powerbi': PowerBICommunityScraper,
    }

    if args.source == 'all':
        scrapers_to_run = list(sources.keys())
    else:
        scrapers_to_run = [args.source]

    for source in scrapers_to_run:
        scraper_class = sources[source]
        scraper = scraper_class(db)

        print(f"\nScraping {source}...")
        questions = scraper.scrape(pages=args.pages)
        all_questions.extend(questions)
        print(f"  Scraped {len(questions)} questions")

    print(f"\nTotal scraped: {len(all_questions)} questions")

    if args.analyze and all_questions:
        print("\nAnalyzing topics...")
        analyzer = TopicAnalyzer(all_questions)

        insights_path = f"{args.output}/topic_insights.json"
        insights = analyzer.export_insights(insights_path)

        print(f"\nTop Topics People Ask About:")
        for word, count in list(insights['common_words'].items())[:15]:
            print(f"  {word}: {count}")

        print(f"\nUnanswered Topics (opportunities to help):")
        for topic in insights['unanswered_topics'][:10]:
            print(f"  {topic['topic']}: {topic['count']} unanswered")

        print(f"\nInsights saved to: {insights_path}")


if __name__ == '__main__':
    main()
