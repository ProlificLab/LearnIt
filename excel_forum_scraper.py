#!/usr/bin/env python3
"""
Excel Forum Scraper - Download and organize discussions from major Excel forums.

Features:
- Incremental updates (only fetch new/changed content)
- Persistent SQLite database storage
- User profile scraping
- Thread tracking and watchlists
- Export to JSON/CSV

Supported forums:
- MrExcel.com
- ExcelForum.com
- Chandoo.org

Usage:
    python excel_forum_scraper.py scrape --forum all --pages 5
    python excel_forum_scraper.py update --incremental
    python excel_forum_scraper.py watch --url "https://..."
    python excel_forum_scraper.py users --forum mrexcel
    python excel_forum_scraper.py stats
"""

import argparse
import csv
import json
import logging
import os
import re
import sys
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, asdict, field
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any, Set, Tuple
from urllib.parse import urljoin, urlparse

import functools

import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

# Try to import ratelimit, fall back to simple implementation
try:
    from ratelimit import limits, sleep_and_retry
except ImportError:
    def limits(calls, period):
        """Simple rate limiting decorator."""
        def decorator(func):
            last_called = [0.0]
            min_interval = period / calls

            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                elapsed = time.time() - last_called[0]
                if elapsed < min_interval:
                    time.sleep(min_interval - elapsed)
                last_called[0] = time.time()
                return func(*args, **kwargs)
            return wrapper
        return decorator

    def sleep_and_retry(func):
        """Simple retry decorator."""
        return func

from database import ForumDatabase, UserProfile

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@dataclass
class Post:
    """Represents a single post/reply in a discussion."""
    author: str
    content: str
    date: str
    post_number: int
    author_id: str = ""
    is_solution: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class Discussion:
    """Represents a forum discussion/thread."""
    title: str
    url: str
    author: str
    date_created: str
    category: str
    tags: List[str] = field(default_factory=list)
    views: int = 0
    replies: int = 0
    posts: List[Post] = field(default_factory=list)
    forum_source: str = ""
    is_solved: bool = False
    last_activity: str = ""
    author_id: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        data = asdict(self)
        data['posts'] = [p.to_dict() for p in self.posts]
        return data


class BaseForumScraper(ABC):
    """Abstract base class for forum scrapers."""

    def __init__(self, db: ForumDatabase, output_dir: str = "./data", delay: float = 1.0):
        self.db = db
        self.output_dir = output_dir
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
        })
        self.discussions: List[Discussion] = []
        self.scraped_users: Set[str] = set()
        self.stats = {
            'threads_found': 0,
            'threads_new': 0,
            'threads_updated': 0,
            'users_found': 0
        }

    @property
    @abstractmethod
    def forum_name(self) -> str:
        """Return the forum name."""
        pass

    @property
    @abstractmethod
    def base_url(self) -> str:
        """Return the base URL of the forum."""
        pass

    @abstractmethod
    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Return list of forum sections with name and URL."""
        pass

    @abstractmethod
    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from a section page."""
        pass

    @abstractmethod
    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single thread and return Discussion object."""
        pass

    @abstractmethod
    def scrape_user_profile(self, username: str, profile_url: str = "") -> Optional[Dict[str, Any]]:
        """Scrape a user's profile page."""
        pass

    @sleep_and_retry
    @limits(calls=30, period=60)
    def fetch_page(self, url: str) -> Optional[BeautifulSoup]:
        """Fetch a page with rate limiting and error handling."""
        try:
            time.sleep(self.delay)
            response = self.session.get(url, timeout=30)
            response.raise_for_status()
            return BeautifulSoup(response.text, 'lxml')
        except requests.RequestException as e:
            logger.error(f"Failed to fetch {url}: {e}")
            return None

    def is_thread_new_or_updated(self, thread_url: str, last_activity: str = "") -> Tuple[bool, bool]:
        """
        Check if a thread is new or has been updated.
        Returns (should_scrape, is_new).
        """
        existing = self.db.get_discussion_by_url(thread_url)

        if not existing:
            return True, True

        # Check if thread has new activity
        if last_activity and existing.get('last_activity'):
            try:
                existing_time = datetime.fromisoformat(existing['last_activity'].replace('Z', '+00:00'))
                new_time = datetime.fromisoformat(last_activity.replace('Z', '+00:00'))
                if new_time > existing_time:
                    return True, False
            except (ValueError, TypeError):
                pass

        # Check if it's been a while since last update
        last_updated = existing.get('last_updated', '')
        if last_updated:
            try:
                updated_time = datetime.fromisoformat(last_updated)
                if datetime.now() - updated_time > timedelta(days=7):
                    return True, False
            except (ValueError, TypeError):
                pass

        return False, False

    def save_discussion_to_db(self, discussion: Discussion) -> Tuple[int, bool]:
        """Save a discussion and its posts to the database."""
        discussion_dict = discussion.to_dict()
        discussion_id, is_new = self.db.upsert_discussion(discussion_dict)

        # Save posts
        for post in discussion.posts:
            post_dict = post.to_dict()
            self.db.upsert_post(discussion_id, post_dict)

        # Track user for later profile scraping
        if discussion.author and discussion.author != "Unknown":
            self.scraped_users.add(discussion.author)

        for post in discussion.posts:
            if post.author and post.author != "Unknown":
                self.scraped_users.add(post.author)

        return discussion_id, is_new

    def scrape(self, max_pages: int = 5, max_threads_per_section: int = 50,
               incremental: bool = True, scrape_profiles: bool = False) -> List[Discussion]:
        """
        Main scraping method.

        Args:
            max_pages: Maximum pages to scrape per section
            max_threads_per_section: Maximum threads per section
            incremental: Only scrape new/updated threads
            scrape_profiles: Also scrape user profiles
        """
        logger.info(f"Starting {'incremental ' if incremental else ''}scrape of {self.forum_name}")

        session_id = self.db.start_scrape_session(self.forum_name)

        try:
            sections = self.get_forum_sections()
            logger.info(f"Found {len(sections)} sections to scrape")

            for section in tqdm(sections, desc=f"Scraping {self.forum_name} sections"):
                section_name = section.get('name', 'Unknown')
                section_url = section.get('url', '')

                if not section_url:
                    continue

                logger.info(f"Scraping section: {section_name}")
                threads_scraped = 0

                for page in range(1, max_pages + 1):
                    if threads_scraped >= max_threads_per_section:
                        break

                    threads = self.get_thread_list(section_url, page)

                    if not threads:
                        break

                    for thread_info in threads:
                        if threads_scraped >= max_threads_per_section:
                            break

                        thread_url = thread_info.get('url', '')
                        if not thread_url:
                            continue

                        self.stats['threads_found'] += 1

                        # Check if we need to scrape this thread
                        if incremental:
                            last_activity = thread_info.get('last_activity', '')
                            should_scrape, is_new = self.is_thread_new_or_updated(
                                thread_url, last_activity
                            )
                            if not should_scrape:
                                continue

                        discussion = self.scrape_thread(thread_url, section_name)
                        if discussion:
                            self.discussions.append(discussion)
                            discussion_id, is_new = self.save_discussion_to_db(discussion)

                            if is_new:
                                self.stats['threads_new'] += 1
                            else:
                                self.stats['threads_updated'] += 1

                            threads_scraped += 1

            # Scrape user profiles if requested
            if scrape_profiles and self.scraped_users:
                logger.info(f"Scraping {len(self.scraped_users)} user profiles...")
                self._scrape_user_profiles()

            self.db.end_scrape_session(
                session_id,
                threads_found=self.stats['threads_found'],
                threads_new=self.stats['threads_new'],
                threads_updated=self.stats['threads_updated'],
                users_found=self.stats['users_found']
            )

        except Exception as e:
            self.db.end_scrape_session(session_id, error=str(e))
            raise

        logger.info(f"Scraped {len(self.discussions)} discussions from {self.forum_name}")
        logger.info(f"  New: {self.stats['threads_new']}, Updated: {self.stats['threads_updated']}")

        return self.discussions

    def _scrape_user_profiles(self):
        """Scrape profiles for all users encountered during scraping."""
        for username in tqdm(self.scraped_users, desc="Scraping user profiles"):
            # Check if we already have recent profile data
            existing = self.db.get_user(username, self.forum_name)
            if existing:
                last_updated = existing.get('last_updated', '')
                if last_updated:
                    try:
                        updated_time = datetime.fromisoformat(last_updated)
                        if datetime.now() - updated_time < timedelta(days=30):
                            continue  # Skip recently updated profiles
                    except (ValueError, TypeError):
                        pass

            profile = self.scrape_user_profile(username)
            if profile:
                profile['forum_source'] = self.forum_name
                self.db.upsert_user(profile)
                self.stats['users_found'] += 1

    def update_watched_threads(self) -> List[Dict[str, Any]]:
        """Update all watched threads and return changed ones."""
        watched = self.db.get_watched_threads()
        updated = []

        for thread in tqdm(watched, desc="Updating watched threads"):
            url = thread.get('url', '')
            if not url:
                continue

            old_replies = thread.get('replies', 0)
            old_solved = thread.get('is_solved', False)

            discussion = self.scrape_thread(url, thread.get('category', ''))
            if discussion:
                self.save_discussion_to_db(discussion)

                # Check if there are changes
                if discussion.replies != old_replies or discussion.is_solved != old_solved:
                    updated.append({
                        'url': url,
                        'title': discussion.title,
                        'old_replies': old_replies,
                        'new_replies': discussion.replies,
                        'newly_solved': discussion.is_solved and not old_solved
                    })

        return updated

    def save_to_json(self, filename: Optional[str] = None) -> str:
        """Save discussions to JSON file."""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.forum_name.lower().replace(' ', '_')}_{timestamp}.json"

        filepath = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)

        data = {
            'forum': self.forum_name,
            'scraped_at': datetime.now().isoformat(),
            'total_discussions': len(self.discussions),
            'statistics': self.stats,
            'discussions': [d.to_dict() for d in self.discussions]
        }

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        logger.info(f"Saved {len(self.discussions)} discussions to {filepath}")
        return filepath

    def save_to_csv(self, filename: Optional[str] = None) -> str:
        """Save discussions to CSV file (summary without full posts)."""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{self.forum_name.lower().replace(' ', '_')}_{timestamp}.csv"

        filepath = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Title', 'URL', 'Author', 'Date Created', 'Category',
                'Tags', 'Views', 'Replies', 'Is Solved', 'Forum Source'
            ])

            for d in self.discussions:
                writer.writerow([
                    d.title, d.url, d.author, d.date_created, d.category,
                    ';'.join(d.tags), d.views, d.replies, d.is_solved, d.forum_source
                ])

        logger.info(f"Saved {len(self.discussions)} discussions to {filepath}")
        return filepath


class MrExcelScraper(BaseForumScraper):
    """Scraper for MrExcel.com forum."""

    @property
    def forum_name(self) -> str:
        return "MrExcel"

    @property
    def base_url(self) -> str:
        return "https://www.mrexcel.com/board/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Get MrExcel forum sections."""
        sections = [
            {'name': 'Excel Questions', 'url': urljoin(self.base_url, 'forums/excel-questions.teknedia/')},
            {'name': 'VBA Macros', 'url': urljoin(self.base_url, 'forums/excel-questions.teknedia/')},
            {'name': 'Power Query', 'url': urljoin(self.base_url, 'forums/power-query.teknedia/')},
            {'name': 'Power Pivot', 'url': urljoin(self.base_url, 'forums/power-pivot.teknedia/')},
            {'name': 'Charts and Graphs', 'url': urljoin(self.base_url, 'forums/charts-and-charting-in-excel.teknedia/')},
        ]
        return sections

    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from a MrExcel section page."""
        url = section_url if page == 1 else f"{section_url}page-{page}"
        soup = self.fetch_page(url)

        if not soup:
            return []

        threads = []
        thread_elements = soup.select('.structItem--thread')

        for elem in thread_elements:
            title_elem = elem.select_one('.structItem-title a')
            time_elem = elem.select_one('.structItem-latestDate time')

            if title_elem:
                thread_url = urljoin(self.base_url, title_elem.get('href', ''))
                last_activity = time_elem.get('datetime', '') if time_elem else ''

                threads.append({
                    'title': title_elem.get_text(strip=True),
                    'url': thread_url,
                    'last_activity': last_activity
                })

        return threads

    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single MrExcel thread."""
        soup = self.fetch_page(thread_url)

        if not soup:
            return None

        try:
            title_elem = soup.select_one('.p-title-value')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            # Get author info
            author_elem = soup.select_one('.message--post .message-userDetails a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"
            author_id = ""
            if author_elem:
                href = author_elem.get('href', '')
                match = re.search(r'/members/[^.]+\.(\d+)', href)
                if match:
                    author_id = match.group(1)

            date_elem = soup.select_one('.message--post time')
            date_created = date_elem.get('datetime', '') if date_elem else ""

            # Get last activity time
            last_time_elem = soup.select_one('.message--post:last-child time')
            last_activity = last_time_elem.get('datetime', '') if last_time_elem else date_created

            # Get reply count
            reply_elem = soup.select_one('.p-description li:contains("Replies")')
            replies = 0
            if reply_elem:
                reply_text = reply_elem.get_text()
                match = re.search(r'(\d+)', reply_text)
                if match:
                    replies = int(match.group(1))

            # Get view count
            view_elem = soup.select_one('.p-description li:contains("Views")')
            views = 0
            if view_elem:
                view_text = view_elem.get_text()
                match = re.search(r'([\d,]+)', view_text)
                if match:
                    views = int(match.group(1).replace(',', ''))

            # Get tags
            tag_elems = soup.select('.tagItem')
            tags = [t.get_text(strip=True) for t in tag_elems]

            # Check if solved
            is_solved = bool(soup.select_one('.label--solved, .label--answered'))

            # Get posts
            posts = []
            post_elems = soup.select('.message--post')

            for idx, post_elem in enumerate(post_elems[:50], 1):
                post_author_elem = post_elem.select_one('.message-userDetails a')
                post_author = post_author_elem.get_text(strip=True) if post_author_elem else "Unknown"

                post_author_id = ""
                if post_author_elem:
                    href = post_author_elem.get('href', '')
                    match = re.search(r'/members/[^.]+\.(\d+)', href)
                    if match:
                        post_author_id = match.group(1)

                post_content = post_elem.select_one('.message-body')
                post_date = post_elem.select_one('time')

                posts.append(Post(
                    author=post_author,
                    author_id=post_author_id,
                    content=post_content.get_text(strip=True)[:10000] if post_content else "",
                    date=post_date.get('datetime', '') if post_date else "",
                    post_number=idx,
                    is_solution=bool(post_elem.select_one('.label--solution'))
                ))

            return Discussion(
                title=title,
                url=thread_url,
                author=author,
                author_id=author_id,
                date_created=date_created,
                category=category,
                tags=tags,
                views=views,
                replies=replies,
                posts=posts,
                forum_source=self.forum_name,
                is_solved=is_solved,
                last_activity=last_activity
            )

        except Exception as e:
            logger.error(f"Error parsing thread {thread_url}: {e}")
            return None

    def scrape_user_profile(self, username: str, profile_url: str = "") -> Optional[Dict[str, Any]]:
        """Scrape a MrExcel user profile."""
        if not profile_url:
            # Try to construct profile URL
            safe_username = username.lower().replace(' ', '-')
            profile_url = urljoin(self.base_url, f'members/{safe_username}/')

        soup = self.fetch_page(profile_url)
        if not soup:
            return None

        try:
            profile = {
                'username': username,
                'profile_url': profile_url,
                'forum_source': self.forum_name
            }

            # Get user ID from URL or page
            user_id_match = re.search(r'/members/[^.]+\.(\d+)', profile_url)
            if user_id_match:
                profile['user_id'] = user_id_match.group(1)

            # Get join date
            join_elem = soup.select_one('dl.pairs:contains("Joined") dd')
            if join_elem:
                profile['join_date'] = join_elem.get_text(strip=True)

            # Get post count
            posts_elem = soup.select_one('dl.pairs:contains("Messages") dd')
            if posts_elem:
                posts_text = posts_elem.get_text(strip=True).replace(',', '')
                try:
                    profile['post_count'] = int(posts_text)
                except ValueError:
                    pass

            # Get reputation/reaction score
            reaction_elem = soup.select_one('dl.pairs:contains("Reaction") dd')
            if reaction_elem:
                reaction_text = reaction_elem.get_text(strip=True).replace(',', '')
                try:
                    profile['reputation'] = int(reaction_text)
                except ValueError:
                    pass

            # Get location
            location_elem = soup.select_one('dl.pairs:contains("Location") dd')
            if location_elem:
                profile['location'] = location_elem.get_text(strip=True)

            # Get title/rank
            title_elem = soup.select_one('.userTitle')
            if title_elem:
                profile['title'] = title_elem.get_text(strip=True)

            # Get avatar
            avatar_elem = soup.select_one('.avatar img')
            if avatar_elem:
                profile['avatar_url'] = avatar_elem.get('src', '')

            # Get last seen
            seen_elem = soup.select_one('dl.pairs:contains("Last seen") dd')
            if seen_elem:
                profile['last_seen'] = seen_elem.get_text(strip=True)

            # Get about/bio
            about_elem = soup.select_one('.aboutInfo')
            if about_elem:
                profile['bio'] = about_elem.get_text(strip=True)[:2000]

            # Get website
            website_elem = soup.select_one('dl.pairs:contains("Website") dd a')
            if website_elem:
                profile['website'] = website_elem.get('href', '')

            return profile

        except Exception as e:
            logger.error(f"Error parsing user profile {username}: {e}")
            return None


class ExcelForumScraper(BaseForumScraper):
    """Scraper for ExcelForum.com."""

    @property
    def forum_name(self) -> str:
        return "ExcelForum"

    @property
    def base_url(self) -> str:
        return "https://www.excelforum.com/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Get ExcelForum sections."""
        sections = [
            {'name': 'Excel General', 'url': urljoin(self.base_url, 'excel-general/')},
            {'name': 'Excel Formulas', 'url': urljoin(self.base_url, 'excel-formulas-functions/')},
            {'name': 'Excel VBA', 'url': urljoin(self.base_url, 'excel-programming-vba-macros/')},
            {'name': 'Excel Charts', 'url': urljoin(self.base_url, 'excel-charting-pivots/')},
            {'name': 'Excel New Users', 'url': urljoin(self.base_url, 'excel-new-users-basics/')},
        ]
        return sections

    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from an ExcelForum section page."""
        url = section_url if page == 1 else f"{section_url}page{page}/"
        soup = self.fetch_page(url)

        if not soup:
            return []

        threads = []
        thread_elements = soup.select('.threadtitle a, .topic-title a, .structItem-title a')

        for elem in thread_elements:
            href = elem.get('href', '')
            if href and ('/threads/' in href or '/topic/' in href):
                thread_url = urljoin(self.base_url, href)
                threads.append({
                    'title': elem.get_text(strip=True),
                    'url': thread_url
                })

        return threads

    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single ExcelForum thread."""
        soup = self.fetch_page(thread_url)

        if not soup:
            return None

        try:
            title_elem = soup.select_one('h1, .thread-title, .topic-title, .p-title-value')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            author_elem = soup.select_one('.postauthor a, .author a, .message-name a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"

            date_elem = soup.select_one('.postdate, .post-date, time')
            date_created = ""
            if date_elem:
                date_created = date_elem.get('datetime', '') or date_elem.get_text(strip=True)

            stats_text = soup.get_text()

            views = 0
            view_match = re.search(r'Views?:?\s*([\d,]+)', stats_text, re.I)
            if view_match:
                views = int(view_match.group(1).replace(',', ''))

            replies = 0
            reply_match = re.search(r'Replies?:?\s*(\d+)', stats_text, re.I)
            if reply_match:
                replies = int(reply_match.group(1))

            posts = []
            post_elems = soup.select('.post, .message, .postcontainer')[:50]

            for idx, post_elem in enumerate(post_elems, 1):
                post_author = post_elem.select_one('.username, .postauthor a, .message-name a')
                post_content = post_elem.select_one('.postcontent, .post-body, .message-body')
                post_date = post_elem.select_one('.postdate, time')

                if post_content:
                    posts.append(Post(
                        author=post_author.get_text(strip=True) if post_author else "Unknown",
                        content=post_content.get_text(strip=True)[:10000],
                        date=post_date.get_text(strip=True) if post_date else "",
                        post_number=idx
                    ))

            return Discussion(
                title=title,
                url=thread_url,
                author=author,
                date_created=date_created,
                category=category,
                views=views,
                replies=replies,
                posts=posts,
                forum_source=self.forum_name
            )

        except Exception as e:
            logger.error(f"Error parsing thread {thread_url}: {e}")
            return None

    def scrape_user_profile(self, username: str, profile_url: str = "") -> Optional[Dict[str, Any]]:
        """Scrape an ExcelForum user profile."""
        if not profile_url:
            safe_username = username.lower().replace(' ', '-')
            profile_url = urljoin(self.base_url, f'members/{safe_username}/')

        soup = self.fetch_page(profile_url)
        if not soup:
            return None

        try:
            profile = {
                'username': username,
                'profile_url': profile_url,
                'forum_source': self.forum_name
            }

            # Get basic info
            info_items = soup.select('.profilefield_content, dl.pairs dd')
            for item in info_items:
                text = item.get_text(strip=True)
                label = item.find_previous_sibling()
                if label:
                    label_text = label.get_text(strip=True).lower()
                    if 'join' in label_text:
                        profile['join_date'] = text
                    elif 'post' in label_text or 'message' in label_text:
                        try:
                            profile['post_count'] = int(text.replace(',', ''))
                        except ValueError:
                            pass
                    elif 'location' in label_text:
                        profile['location'] = text

            return profile

        except Exception as e:
            logger.error(f"Error parsing user profile {username}: {e}")
            return None


class ChandooScraper(BaseForumScraper):
    """Scraper for Chandoo.org forum."""

    @property
    def forum_name(self) -> str:
        return "Chandoo"

    @property
    def base_url(self) -> str:
        return "https://forum.chandoo.org/"

    def get_forum_sections(self) -> List[Dict[str, str]]:
        """Get Chandoo forum sections."""
        sections = [
            {'name': 'Excel Help', 'url': urljoin(self.base_url, 'forums/excel-help.5/')},
            {'name': 'VBA Macros', 'url': urljoin(self.base_url, 'forums/vba-macros.6/')},
            {'name': 'Power Query', 'url': urljoin(self.base_url, 'forums/power-query.32/')},
            {'name': 'Charts', 'url': urljoin(self.base_url, 'forums/excel-charts.7/')},
            {'name': 'Dashboards', 'url': urljoin(self.base_url, 'forums/excel-dashboards.8/')},
        ]
        return sections

    def get_thread_list(self, section_url: str, page: int = 1) -> List[Dict[str, str]]:
        """Get list of threads from a Chandoo section page."""
        url = section_url if page == 1 else f"{section_url}page-{page}"
        soup = self.fetch_page(url)

        if not soup:
            return []

        threads = []
        thread_elements = soup.select('.structItem--thread')

        for elem in thread_elements:
            title_elem = elem.select_one('.structItem-title a')
            time_elem = elem.select_one('.structItem-latestDate time')

            if title_elem:
                href = title_elem.get('href', '')
                thread_url = urljoin(self.base_url, href)
                last_activity = time_elem.get('datetime', '') if time_elem else ''

                threads.append({
                    'title': title_elem.get_text(strip=True),
                    'url': thread_url,
                    'last_activity': last_activity
                })

        return threads

    def scrape_thread(self, thread_url: str, category: str) -> Optional[Discussion]:
        """Scrape a single Chandoo thread."""
        soup = self.fetch_page(thread_url)

        if not soup:
            return None

        try:
            title_elem = soup.select_one('.p-title-value, h1')
            title = title_elem.get_text(strip=True) if title_elem else "Unknown"

            author_elem = soup.select_one('.message--post .message-userDetails a')
            author = author_elem.get_text(strip=True) if author_elem else "Unknown"

            date_elem = soup.select_one('.message--post time')
            date_created = date_elem.get('datetime', '') if date_elem else ""

            # Get last activity
            last_time_elem = soup.select_one('.message--post:last-child time')
            last_activity = last_time_elem.get('datetime', '') if last_time_elem else date_created

            replies = len(soup.select('.message--post')) - 1

            tag_elems = soup.select('.tagItem')
            tags = [t.get_text(strip=True) for t in tag_elems]

            is_solved = bool(soup.select_one('[class*="solved"], [class*="answered"]'))

            posts = []
            post_elems = soup.select('.message--post')[:50]

            for idx, post_elem in enumerate(post_elems, 1):
                post_author = post_elem.select_one('.message-userDetails a')
                post_content = post_elem.select_one('.message-body .bbWrapper')
                post_date = post_elem.select_one('time')

                if post_content:
                    posts.append(Post(
                        author=post_author.get_text(strip=True) if post_author else "Unknown",
                        content=post_content.get_text(strip=True)[:10000],
                        date=post_date.get('datetime', '') if post_date else "",
                        post_number=idx,
                        is_solution=bool(post_elem.select_one('[class*="solution"]'))
                    ))

            return Discussion(
                title=title,
                url=thread_url,
                author=author,
                date_created=date_created,
                category=category,
                tags=tags,
                replies=max(0, replies),
                posts=posts,
                forum_source=self.forum_name,
                is_solved=is_solved,
                last_activity=last_activity
            )

        except Exception as e:
            logger.error(f"Error parsing thread {thread_url}: {e}")
            return None

    def scrape_user_profile(self, username: str, profile_url: str = "") -> Optional[Dict[str, Any]]:
        """Scrape a Chandoo user profile."""
        if not profile_url:
            safe_username = username.lower().replace(' ', '-')
            profile_url = urljoin(self.base_url, f'members/{safe_username}/')

        soup = self.fetch_page(profile_url)
        if not soup:
            return None

        try:
            profile = {
                'username': username,
                'profile_url': profile_url,
                'forum_source': self.forum_name
            }

            # Similar structure to MrExcel (XenForo)
            pairs = soup.select('dl.pairs')
            for pair in pairs:
                dt = pair.select_one('dt')
                dd = pair.select_one('dd')
                if dt and dd:
                    label = dt.get_text(strip=True).lower()
                    value = dd.get_text(strip=True)

                    if 'joined' in label:
                        profile['join_date'] = value
                    elif 'message' in label:
                        try:
                            profile['post_count'] = int(value.replace(',', ''))
                        except ValueError:
                            pass
                    elif 'reaction' in label:
                        try:
                            profile['reputation'] = int(value.replace(',', ''))
                        except ValueError:
                            pass
                    elif 'location' in label:
                        profile['location'] = value
                    elif 'last seen' in label:
                        profile['last_seen'] = value

            title_elem = soup.select_one('.userTitle')
            if title_elem:
                profile['title'] = title_elem.get_text(strip=True)

            avatar_elem = soup.select_one('.avatar img')
            if avatar_elem:
                profile['avatar_url'] = avatar_elem.get('src', '')

            return profile

        except Exception as e:
            logger.error(f"Error parsing user profile {username}: {e}")
            return None


def get_scraper(forum: str, db: ForumDatabase, output_dir: str, delay: float) -> Optional[BaseForumScraper]:
    """Get scraper instance for a forum."""
    scrapers = {
        'mrexcel': MrExcelScraper,
        'excelforum': ExcelForumScraper,
        'chandoo': ChandooScraper
    }
    scraper_class = scrapers.get(forum.lower())
    if scraper_class:
        return scraper_class(db, output_dir, delay)
    return None


def cmd_scrape(args):
    """Handle scrape command."""
    db = ForumDatabase(args.database)

    if args.forum == 'all':
        forums = ['mrexcel', 'excelforum', 'chandoo']
    else:
        forums = [args.forum]

    all_discussions = []

    for forum in forums:
        scraper = get_scraper(forum, db, args.output, args.delay)
        if not scraper:
            continue

        try:
            discussions = scraper.scrape(
                max_pages=args.pages,
                max_threads_per_section=args.threads,
                incremental=not args.full,
                scrape_profiles=args.profiles
            )
            all_discussions.extend(discussions)

            if args.format in ('json', 'both'):
                scraper.save_to_json()
            if args.format in ('csv', 'both'):
                scraper.save_to_csv()

        except KeyboardInterrupt:
            logger.info("Scraping interrupted by user")
            break
        except Exception as e:
            logger.error(f"Error scraping {forum}: {e}")

    print(f"\nScraped {len(all_discussions)} total discussions")
    print(f"Database: {args.database}")


def cmd_update(args):
    """Handle update command - update watched threads."""
    db = ForumDatabase(args.database)

    if args.forum == 'all':
        forums = ['mrexcel', 'excelforum', 'chandoo']
    else:
        forums = [args.forum]

    for forum in forums:
        scraper = get_scraper(forum, db, args.output, args.delay)
        if not scraper:
            continue

        updated = scraper.update_watched_threads()

        if updated:
            print(f"\n{forum} - Updated threads:")
            for u in updated:
                status = " [SOLVED]" if u.get('newly_solved') else ""
                print(f"  {u['title']}: {u['old_replies']} -> {u['new_replies']} replies{status}")
        else:
            print(f"\n{forum} - No updates found")


def cmd_watch(args):
    """Handle watch command - add/remove threads from watchlist."""
    db = ForumDatabase(args.database)

    if args.action == 'add':
        # Find or create the discussion
        existing = db.get_discussion_by_url(args.url)
        if existing:
            db.watch_thread(existing['id'], args.frequency, args.notes or '')
            print(f"Now watching: {existing['title']}")
        else:
            print(f"Thread not in database. Scraping...")
            # Determine forum from URL
            if 'mrexcel.com' in args.url:
                forum = 'mrexcel'
            elif 'excelforum.com' in args.url:
                forum = 'excelforum'
            elif 'chandoo.org' in args.url:
                forum = 'chandoo'
            else:
                print("Unknown forum URL")
                return

            scraper = get_scraper(forum, db, args.output, 1.0)
            discussion = scraper.scrape_thread(args.url, 'Watched')
            if discussion:
                disc_id, _ = scraper.save_discussion_to_db(discussion)
                db.watch_thread(disc_id, args.frequency, args.notes or '')
                print(f"Now watching: {discussion.title}")
            else:
                print("Failed to scrape thread")

    elif args.action == 'remove':
        existing = db.get_discussion_by_url(args.url)
        if existing:
            db.unwatch_thread(existing['id'])
            print(f"Stopped watching: {existing['title']}")
        else:
            print("Thread not found")

    elif args.action == 'list':
        watched = db.get_watched_threads()
        if watched:
            print("\nWatched Threads:")
            print("-" * 80)
            for w in watched:
                solved = " [SOLVED]" if w.get('is_solved') else ""
                print(f"  {w['title']}{solved}")
                print(f"    URL: {w['url']}")
                print(f"    Replies: {w['replies']} | Frequency: {w['check_frequency']}")
                print()
        else:
            print("No watched threads")


def cmd_users(args):
    """Handle users command - list/export user profiles."""
    db = ForumDatabase(args.database)

    forum = args.forum if args.forum != 'all' else None
    users = db.get_users(forum, limit=args.limit)

    if args.export:
        filepath = os.path.join(args.output, f"users_{datetime.now().strftime('%Y%m%d')}.json")
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(users, f, indent=2, ensure_ascii=False)
        print(f"Exported {len(users)} users to {filepath}")
    else:
        print(f"\nTop Users (by post count):")
        print("-" * 60)
        for u in users[:20]:
            print(f"  {u['username']} ({u['forum_source']})")
            print(f"    Posts: {u.get('post_count', 0)} | Rep: {u.get('reputation', 0)}")
            if u.get('location'):
                print(f"    Location: {u['location']}")
            print()


def cmd_stats(args):
    """Handle stats command - show database statistics."""
    db = ForumDatabase(args.database)
    stats = db.get_statistics()

    print("\n=== Forum Database Statistics ===")
    print(f"Total discussions: {stats['total_discussions']:,}")
    print(f"Solved discussions: {stats['solved_discussions']:,}")
    print(f"Total posts: {stats['total_posts']:,}")
    print(f"Total users: {stats['total_users']:,}")
    print(f"Watched threads: {stats['watched_threads']}")
    print(f"Last update: {stats.get('last_update', 'Never')}")

    print("\nBy Forum:")
    for forum, count in stats.get('by_forum', {}).items():
        print(f"  {forum}: {count:,}")

    print("\nTop Categories:")
    for cat, count in list(stats.get('top_categories', {}).items())[:10]:
        print(f"  {cat}: {count:,}")


def cmd_search(args):
    """Handle search command."""
    db = ForumDatabase(args.database)
    results = db.search_discussions(args.query, limit=args.limit)

    if results:
        print(f"\nFound {len(results)} discussions matching '{args.query}':")
        print("-" * 80)
        for r in results:
            solved = " [SOLVED]" if r.get('is_solved') else ""
            print(f"  {r['title']}{solved}")
            print(f"    Forum: {r['forum_source']} | Category: {r['category']}")
            print(f"    Views: {r['views']:,} | Replies: {r['replies']}")
            print(f"    URL: {r['url']}")
            print()
    else:
        print(f"No discussions found matching '{args.query}'")


def cmd_export(args):
    """Handle export command."""
    db = ForumDatabase(args.database)

    forum = args.forum if args.forum != 'all' else None
    filepath = os.path.join(args.output, f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")

    count = db.export_to_json(filepath, forum)
    print(f"Exported {count} discussions to {filepath}")


def main():
    """Main entry point for the scraper."""
    parser = argparse.ArgumentParser(
        description='Excel Forum Scraper with incremental updates and user profiles',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    # Global options
    parser.add_argument('--database', '-db', default='./data/forum_data.db',
                        help='Database file path (default: ./data/forum_data.db)')
    parser.add_argument('--output', '-o', default='./data',
                        help='Output directory (default: ./data)')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Enable verbose logging')

    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # Scrape command
    scrape_parser = subparsers.add_parser('scrape', help='Scrape forums')
    scrape_parser.add_argument('--forum', choices=['mrexcel', 'excelforum', 'chandoo', 'all'],
                               default='all', help='Forum to scrape')
    scrape_parser.add_argument('--pages', '-p', type=int, default=5,
                               help='Max pages per section')
    scrape_parser.add_argument('--threads', '-t', type=int, default=50,
                               help='Max threads per section')
    scrape_parser.add_argument('--delay', '-d', type=float, default=1.0,
                               help='Delay between requests')
    scrape_parser.add_argument('--full', action='store_true',
                               help='Full scrape (not incremental)')
    scrape_parser.add_argument('--profiles', action='store_true',
                               help='Also scrape user profiles')
    scrape_parser.add_argument('--format', '-f', choices=['json', 'csv', 'both'],
                               default='both', help='Export format')

    # Update command
    update_parser = subparsers.add_parser('update', help='Update watched threads')
    update_parser.add_argument('--forum', choices=['mrexcel', 'excelforum', 'chandoo', 'all'],
                               default='all', help='Forum to update')
    update_parser.add_argument('--delay', '-d', type=float, default=1.0,
                               help='Delay between requests')

    # Watch command
    watch_parser = subparsers.add_parser('watch', help='Manage watched threads')
    watch_parser.add_argument('action', choices=['add', 'remove', 'list'],
                              help='Action to perform')
    watch_parser.add_argument('--url', help='Thread URL (for add/remove)')
    watch_parser.add_argument('--frequency', choices=['hourly', 'daily', 'weekly'],
                              default='daily', help='Check frequency')
    watch_parser.add_argument('--notes', help='Notes about the thread')

    # Users command
    users_parser = subparsers.add_parser('users', help='Manage user profiles')
    users_parser.add_argument('--forum', choices=['mrexcel', 'excelforum', 'chandoo', 'all'],
                              default='all', help='Filter by forum')
    users_parser.add_argument('--limit', type=int, default=100,
                              help='Max users to show')
    users_parser.add_argument('--export', action='store_true',
                              help='Export to JSON')

    # Stats command
    subparsers.add_parser('stats', help='Show database statistics')

    # Search command
    search_parser = subparsers.add_parser('search', help='Search discussions')
    search_parser.add_argument('query', help='Search query')
    search_parser.add_argument('--limit', type=int, default=50,
                               help='Max results')

    # Export command
    export_parser = subparsers.add_parser('export', help='Export database')
    export_parser.add_argument('--forum', choices=['mrexcel', 'excelforum', 'chandoo', 'all'],
                               default='all', help='Filter by forum')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    if not args.command:
        parser.print_help()
        return 1

    commands = {
        'scrape': cmd_scrape,
        'update': cmd_update,
        'watch': cmd_watch,
        'users': cmd_users,
        'stats': cmd_stats,
        'search': cmd_search,
        'export': cmd_export
    }

    try:
        commands[args.command](args)
    except KeyboardInterrupt:
        print("\nInterrupted")
        return 1
    except Exception as e:
        logger.error(f"Error: {e}")
        if args.verbose:
            raise
        return 1

    return 0


if __name__ == '__main__':
    sys.exit(main())
